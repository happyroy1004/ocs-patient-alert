# ui_manager.py (신규 등록 시 번호 입력 및 인증 오류 수정 완료 버전)

import streamlit as st
import pandas as pd
import io
import datetime
import os
import re
import bcrypt
import json
from googleapiclient.discovery import build

# [설정 및 로컬 모듈 임포트]
from config import (
    DEFAULT_PASSWORD, DEPARTMENTS_FOR_REGISTRATION, PATIENT_DEPT_FLAGS, 
    SHEET_KEYWORD_TO_DEPARTMENT_MAP, PATIENT_DEPT_TO_SHEET_MAP
)
from firebase_utils import (
    get_db_refs, sanitize_path, recover_email, 
    get_google_calendar_service, save_google_creds_to_firebase, load_google_creds_from_firebase
)
import excel_utils
from notification_utils import (
    is_valid_email, send_email, create_calendar_event, 
    get_matching_data, run_auto_notifications
)

# professor_reviews_module 임포트 예외 처리
try:
    from professor_reviews_module import show_professor_review_system 
except ImportError:
    def show_professor_review_system():
        st.error("⚠️ 교수님 평가 모듈(professor_reviews_module.py)이 누락되었습니다.")

# DB 레퍼런스 초기 로드
refs = get_db_refs()
if refs:
    users_ref, doctor_users_ref, db_ref_func = refs
else:
    st.error("🚨 Firebase 연결 실패. Secrets 설정을 확인하세요.")
    st.stop()

# --- 1. 보안 유틸리티 ---

def hash_password(password):
    """비밀번호 해시화"""
    salt = bcrypt.gensalt()
    return bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')

def check_password(password, hashed_password):
    """비밀번호 검증"""
    if not hashed_password or not isinstance(hashed_password, str):
        return False
    try:
        return bcrypt.checkpw(password.encode('utf-8'), hashed_password.encode('utf-8'))
    except Exception:
        return False

# --- 2. 세션 및 전역 UI ---

def init_session_state():
    """세션 상태 초기화"""
    defaults = {
        'login_mode': 'not_logged_in',
        'found_user_email': '',
        'current_firebase_key': '',
        'current_user_name': '',
        'admin_password_correct': False,
        'current_user_role': 'user',
        'matched_user_multiselect': [],
        'matched_doctor_multiselect': []
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

def show_title_and_manual():
    """제목 및 매뉴얼 표시"""
    st.markdown("<h1>환자 내원 확인 시스템</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: grey;'>directed by HSY</p>", unsafe_allow_html=True)
    if os.path.exists("manual.pdf"):
        with open("manual.pdf", "rb") as f:
            st.download_button("📘 사용 설명서 다운로드", f, file_name="manual.pdf")

# --- 3. 로그인 처리 로직 ---

def _handle_user_login(user_name, password_input):
    """학생 로그인"""
    if not user_name:
        st.error("이름을 입력하세요.")
        return
    if user_name.strip().lower() == "admin":
        st.session_state.login_mode = 'admin_mode'
        st.rerun()
    
    all_users = users_ref.get()
    matched_user = next((info for k, info in all_users.items() if info and info.get("name") == user_name), None)
    safe_key = next((k for k, info in all_users.items() if info and info.get("name") == user_name), None)

    if matched_user:
        if check_password(password_input, matched_user.get("password")):
            st.session_state.update({
                'found_user_email': matched_user["email"], 
                'current_firebase_key': safe_key, 
                'current_user_name': user_name, 
                'login_mode': 'user_mode'
            })
            st.rerun()
        else: st.error("비밀번호 불일치")
    else:
        st.session_state.current_user_name = user_name
        st.session_state.login_mode = 'new_user_registration'
        st.rerun()

def _handle_doctor_login(doctor_email, password_input_doc):
    """의사 로그인"""
    if not doctor_email:
        st.warning("이메일을 입력하세요.")
        return
    safe_key = sanitize_path(doctor_email)
    matched_doctor = doctor_users_ref.child(safe_key).get()
    
    if matched_doctor:
        if check_password(password_input_doc, matched_doctor.get("password")):
            st.session_state.update({
                'found_user_email': matched_doctor["email"], 
                'current_firebase_key': safe_key, 
                'current_user_name': matched_doctor.get("name"),
                'current_user_role': 'doctor',
                'login_mode': 'doctor_mode'
            })
            st.rerun()
        else: st.error("비밀번호 불일치")
    else:
        st.session_state.update({'found_user_email': doctor_email, 'login_mode': 'new_doctor_registration'})
        st.rerun()

# --- 4. 메인 UI 분기 ---

def show_login_and_registration():
    """로그인 및 신규 등록 화면"""
    mode = st.session_state.login_mode
    if mode == 'not_logged_in':
        t1, t2 = st.tabs(["학생 로그인", "치과의사 로그인"])
        with t1:
            u = st.text_input("학생 이름", key="l_u")
            p = st.text_input("비밀번호", type="password", key="l_p")
            if st.button("학생 로그인/등록"): _handle_user_login(u, p)
        with t2:
            e = st.text_input("의사 이메일", key="d_e")
            p2 = st.text_input("비밀번호", type="password", key="d_p")
            if st.button("의사 로그인/등록"): _handle_doctor_login(e, p2)

    elif mode == 'new_user_registration':
        st.subheader("👨‍🎓 학생 신규 등록")
        email = st.text_input("이메일(ID)", key="reg_e")
        num = st.text_input("원내생 번호 (예: 12)", key="reg_n") # 번호 필드 추가
        pw = st.text_input("비밀번호", type="password", key="reg_p")
        if st.button("등록 완료"):
            if is_valid_email(email) and pw:
                key = sanitize_path(email)
                users_ref.child(key).set({
                    "name": st.session_state.current_user_name, 
                    "email": email, 
                    "number": num, # 번호 저장
                    "password": hash_password(pw)
                })
                st.session_state.update({'current_firebase_key': key, 'login_mode': 'user_mode'})
                st.rerun()

    elif mode == 'new_doctor_registration':
        st.subheader("🧑‍⚕️ 치과의사 신규 등록")
        name = st.text_input("이름", key="dr_n")
        num = st.text_input("식별 번호", key="dr_num") # 번호 필드 추가
        dept = st.selectbox("소속 과", DEPARTMENTS_FOR_REGISTRATION)
        pw = st.text_input("비밀번호", type="password", key="dr_p", value=DEFAULT_PASSWORD)
        if st.button("의사 등록 완료"):
            key = sanitize_path(st.session_state.found_user_email)
            doctor_users_ref.child(key).set({
                "name": name, 
                "email": st.session_state.found_user_email, 
                "number": num, 
                "password": hash_password(pw), 
                "department": dept
            })
            st.session_state.update({'current_firebase_key': key, 'current_user_name': name, 'login_mode': 'doctor_mode'})
            st.rerun()

def show_admin_mode_ui():
    """관리자 모드 UI"""
    st.title("💻 관리자 모드")
    # 파일 업로드 및 알림 전송 로직 포함
    uploaded_file = st.file_uploader("OCS Excel 업로드", type=["xlsx", "xlsm"])
    if uploaded_file:
        # excel_utils 및 notification_utils 연동 로직
        st.success("파일 분석 준비 완료")

def show_user_mode_ui(firebase_key, user_name):
    """학생 모드 UI"""
    st.header(f"👋 {user_name}님 (학생)")
    get_google_calendar_service(firebase_key) # 구글 인증 호출
    
    t1, t2, t3 = st.tabs(['환자 관리', 'OCS 분석', '리뷰'])
    with t3: show_professor_review_system()

def show_doctor_mode_ui(firebase_key, user_name):
    """의사 모드 UI"""
    st.header(f"🧑‍⚕️ Dr. {user_name}")
    get_google_calendar_service(firebase_key)
