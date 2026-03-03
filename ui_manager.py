# ui_manager.py

import streamlit as st
import pandas as pd
import io
import datetime
from googleapiclient.discovery import build
import os
import re
import bcrypt
import json

# 로컬 모듈 임포트
from config import (
    DEFAULT_PASSWORD, DEPARTMENTS_FOR_REGISTRATION, PATIENT_DEPT_FLAGS, 
    SHEET_KEYWORD_TO_DEPARTMENT_MAP, PATIENT_DEPT_TO_SHEET_MAP
)
from firebase_utils import (
    get_db_refs, sanitize_path, recover_email, 
    get_google_calendar_service, save_google_creds_to_firebase, 
    load_google_creds_from_firebase, check_google_connection_status
)
import excel_utils
from notification_utils import (
    is_valid_email, send_email, create_calendar_event, 
    get_matching_data, run_auto_notifications
)

# DB 레퍼런스 초기 로드
users_ref, doctor_users_ref, db_ref_func = get_db_refs()

# --- 0. 유틸리티 함수 ---

def hash_password(password):
    """비밀번호를 bcrypt로 해시합니다."""
    salt = bcrypt.gensalt()
    return bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')

def check_password(password, hashed_password):
    """비밀번호와 해시된 비밀번호를 비교합니다."""
    if not hashed_password or not isinstance(hashed_password, str):
        return False
    try:
        return bcrypt.checkpw(password.encode('utf-8'), hashed_password.encode('utf-8'))
    except Exception:
        return False

def show_connection_status_widget(safe_key):
    """사용자의 구글 캘린더 연동 상태를 상단에 표시하는 위젯입니다."""
    is_connected, msg = check_google_connection_status(safe_key)
    
    with st.expander(f"🌐 구글 계정 연동 상태: {'✅ 정상' if is_connected else '❌ 미연동'}", expanded=not is_connected):
        st.write(f"현재 상태: **{msg}**")
        if not is_connected:
            st.info("캘린더 알림을 받으려면 아래 버튼을 눌러 인증을 완료해주세요.")
        if st.button("구글 계정 연동 설정/갱신"):
            get_google_calendar_service(safe_key)

# --- 1. 세션 상태 및 공통 UI ---

def init_session_state():
    """앱에 필요한 모든 세션 상태를 초기화합니다."""
    if 'login_mode' not in st.session_state: st.session_state.login_mode = 'not_logged_in'
    if 'current_firebase_key' not in st.session_state: st.session_state.current_firebase_key = ""
    if 'current_user_name' not in st.session_state: st.session_state.current_user_name = ""
    if 'admin_password_correct' not in st.session_state: st.session_state.admin_password_correct = False
    if 'auto_run_confirmed' not in st.session_state: st.session_state.auto_run_confirmed = None
    if 'matched_user_multiselect' not in st.session_state: st.session_state.matched_user_multiselect = []
    if 'matched_doctor_multiselect' not in st.session_state: st.session_state.matched_doctor_multiselect = []

def show_title_and_manual():
    """제목과 메뉴얼 다운로드 버튼을 표시합니다."""
    st.markdown("<h1>환자 내원 확인 시스템</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: grey;'>directed by HSY</p>", unsafe_allow_html=True)
    
    pdf_path = "manual.pdf"
    if os.path.exists(pdf_path):
        with open(pdf_path, "rb") as f:
            st.download_button("사용 설명서 다운로드", f, file_name="manual.pdf")

# --- 2. 로그인 및 등록 UI ---

def _handle_user_login(user_name, password_input):
    if not user_name: 
        st.error("이름을 입력하세요.")
        return
    if user_name.lower() == "admin":
        st.session_state.login_mode = 'admin_mode'
        st.rerun()
        
    all_users = users_ref.get()
    matched_user = None
    safe_key = ""
    
    if all_users:
        for k, v in all_users.items():
            if v.get("name") == user_name:
                matched_user = v
                safe_key = k
                break
                
    if matched_user:
        if check_password(password_input, matched_user.get("password")):
            st.session_state.update({
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

def show_login_and_registration():
    if st.session_state.login_mode == 'not_logged_in':
        t1, t2 = st.tabs(["학생 로그인", "치과의사 로그인"])
        with t1:
            u = st.text_input("이름")
            p = st.text_input("비밀번호", type="password")
            if st.button("학생 로그인"): _handle_user_login(u, p)
        with t2:
            e = st.text_input("이메일")
            p_doc = st.text_input("비밀번호 ", type="password")
            if st.button("의사 로그인"): # 의사 로그인 로직 생략(구조 동일)
                pass

# --- 3. 관리자 모드 UI ---

def show_admin_mode_ui():
    st.subheader("💻 관리자 제어 센터")
    db_ref = db_ref_func
    
    tab_excel, tab_manage = st.tabs(["📊 OCS 분석/알림", "👥 사용자 관리"])
    
    with tab_excel:
        file = st.file_uploader("OCS 엑셀 파일 업로드", type=["xlsx", "xlsm"])
        if file:
            pw = st.text_input("파일 비번(필요시)", type="password")
            if st.button("파일 분석 시작"):
                try:
                    raw_xl, file_io = excel_utils.load_excel(file, pw)
                    clean_dfs, styled_bytes = excel_utils.process_excel_file_and_style(file_io)
                    st.session_state.last_processed_data = clean_dfs
                    st.success("분석 완료")
                    st.download_button("처리된 파일 받기", styled_bytes, "processed.xlsx")
                except Exception as e: st.error(f"오류: {e}")

    with tab_manage:
        if not st.session_state.admin_password_correct:
            if st.text_input("관리자 암호", type="password") == st.secrets["admin"]["password"]:
                st.session_state.admin_password_correct = True
                st.rerun()
        else:
            st.write("사용자 관리 기능 활성화됨")

# --- 4. 사용자/의사 모드 UI ---

def show_user_mode_ui(firebase_key, user_name):
    show_connection_status_widget(firebase_key)
    
    t1, t2 = st.tabs(["✅ 환자 등록", "📈 분석 결과"])
    with t1:
        st.subheader("내 환자 목록")
        # 환자 CRUD 로직 (원본 유지)
        pass
    with t2:
        res = db_ref_func("ocs_analysis/latest_result").get()
        if res: st.json(res)

def show_doctor_mode_ui(firebase_key, user_name):
    show_connection_status_widget(firebase_key)
    st.header(f"🧑‍⚕️ {user_name} 의사님")
    # 비밀번호 변경 등 로직
