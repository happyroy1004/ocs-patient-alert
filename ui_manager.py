# ui_manager.py (신규 등록 시 번호 입력 및 관리자 모드 전체 포함 버전)

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
        st.error("⚠️ 교수님 평가 모듈(professor_reviews_module.py)을 찾을 수 없습니다.")

# DB 레퍼런스 초기 로드
refs = get_db_refs()
if refs:
    users_ref, doctor_users_ref, db_ref_func = refs
else:
    st.error("🚨 Firebase 초기화에 실패했습니다. st.secrets 설정을 확인해 주세요.")
    st.stop()

# --- 1. 보안 유틸리티 ---

def hash_password(password):
    """비밀번호 해시화"""
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

def check_password(password, hashed_password):
    """비밀번호 검증"""
    if not hashed_password or not isinstance(hashed_password, str):
        return False
    try:
        return bcrypt.checkpw(password.encode('utf-8'), hashed_password.encode('utf-8'))
    except Exception:
        return False

# --- 2. 세션 상태 및 기본 UI ---

def init_session_state():
    """앱의 모든 세션 상태를 초기화"""
    defaults = {
        'login_mode': 'not_logged_in',
        'found_user_email': '',
        'current_firebase_key': '',
        'current_user_name': '',
        'current_user_role': 'user',
        'admin_password_correct': False,
        'auto_run_confirmed': None,
        'matched_user_multiselect': [],
        'matched_doctor_multiselect': [],
        'delete_patient_confirm': False,
        'patients_to_delete': []
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

def show_title_and_manual():
    """제목 및 매뉴얼 표시"""
    st.markdown("<h1>환자 내원 확인 시스템</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: grey; font-size: small;'>directed by HSY</p>", unsafe_allow_html=True)
    pdf_file_path = "manual.pdf"
    if os.path.exists(pdf_file_path):
        with open(pdf_file_path, "rb") as f:
            st.download_button("📘 사용 설명서 다운로드", f, file_name="manual.pdf")

# --- 3. 로그인 및 등록 처리 ---

def _handle_user_login(user_name, password_input):
    """학생 로그인 로직"""
    if not user_name:
        st.error("이름을 입력해주세요.")
        return
    if user_name.strip().lower() == "admin":
        st.session_state.login_mode = 'admin_mode'
        st.rerun()
    all_users = users_ref.get()
    matched_user = None
    safe_key_found = None
    if all_users:
        for safe_key, info in all_users.items():
            if info and info.get("name") == user_name:
                matched_user, safe_key_found = info, safe_key
                break
    if matched_user:
        if check_password(password_input, matched_user.get("password")):
            st.session_state.update({
                'found_user_email': matched_user["email"], 
                'current_firebase_key': safe_key_found, 
                'current_user_name': user_name, 
                'login_mode': 'user_mode'
            })
            st.rerun()
        else: st.error("비밀번호가 틀렸습니다.")
    else:
        st.session_state.current_user_name = user_name
        st.session_state.login_mode = 'new_user_registration'
        st.rerun()

def _handle_doctor_login(doctor_email, password_input_doc):
    """의사 로그인 로직"""
    if not doctor_email:
        st.warning("이메일을 입력해주세요.")
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
        else: st.error("비밀번호가 틀렸습니다.")
    else:
        st.session_state.update({'found_user_email': doctor_email, 'login_mode': 'new_doctor_registration'})
        st.rerun()

def show_login_and_registration():
    """로그인/등록 UI 메인"""
    mode = st.session_state.login_mode
    if mode == 'not_logged_in':
        t1, t2 = st.tabs(["학생 로그인", "의사 로그인"])
        with t1:
            u = st.text_input("이름 (학생)", key="l_u")
            p = st.text_input("비밀번호", type="password", key="l_p")
            if st.button("로그인 / 신규 등록 (학생)"): _handle_user_login(u, p)
        with t2:
            e = st.text_input("이메일 (의사)", key="l_e")
            p2 = st.text_input("비밀번호", type="password", key="l_p2")
            if st.button("로그인 / 신규 등록 (의사)"): _handle_doctor_login(e, p2)

    elif mode == 'new_user_registration':
        st.subheader("👨‍🎓 학생 신규 등록")
        email = st.text_input("이메일 주소", key="r_u_e")
        number = st.text_input("원내생 번호 (예: 12)", key="r_u_n")
        pw = st.text_input("새 비밀번호", type="password", key="r_u_p")
        if st.button("학생 등록 완료"):
            if is_valid_email(email) and pw:
                key = sanitize_path(email)
                users_ref.child(key).set({
                    "name": st.session_state.current_user_name,
                    "email": email, "number": number, "password": hash_password(pw)
                })
                st.session_state.update({'current_firebase_key': key, 'login_mode': 'user_mode'})
                st.rerun()

    elif mode == 'new_doctor_registration':
        st.subheader("🧑‍⚕️ 치과의사 신규 등록")
        name = st.text_input("성함", key="r_d_n")
        num = st.text_input("식별 번호", key="r_d_num")
        dept = st.selectbox("진료과", DEPARTMENTS_FOR_REGISTRATION)
        pw = st.text_input("비밀번호", type="password", key="r_d_p", value=DEFAULT_PASSWORD)
        if st.button("의사 등록 완료"):
            if name and pw:
                key = sanitize_path(st.session_state.found_user_email)
                doctor_users_ref.child(key).set({
                    "name": name, "email": st.session_state.found_user_email,
                    "number": num, "password": hash_password(pw), 
                    "department": dept, "role": "doctor"
                })
                st.session_state.update({'current_firebase_key': key, 'current_user_name': name, 'login_mode': 'doctor_mode'})
                st.rerun()

# --- 4. 관리자 모드 UI (전체 포함) ---

def show_admin_mode_ui():
    """관리자 모드 UI"""
    st.markdown("---")
    st.title("💻 관리자 모드")
    try:
        sender = st.secrets["gmail"]["sender"]
        sender_pw = st.secrets["gmail"]["app_password"]
    except KeyError:
        st.error("⚠️ [gmail] Secrets 정보 누락")
        return

    tab_excel, tab_user_mgmt = st.tabs(["📊 OCS 파일 처리 및 알림", "🧑‍💻 사용자 목록 및 관리"])

    with tab_excel:
        st.subheader("Excel File Processor")
        uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])
        if uploaded_file:
            file_name = uploaded_file.name
            is_daily = excel_utils.is_daily_schedule(file_name)
            password = st.text_input("비밀번호", type="password") if excel_utils.is_encrypted_excel(uploaded_file) else None
            if excel_utils.is_encrypted_excel(uploaded_file) and not password: st.stop()
            try:
                xl_obj, raw_io = excel_utils.load_excel(uploaded_file, password)
                excel_data_dfs_raw, styled_bytes = excel_utils.process_excel_file_and_style(raw_io, db_ref_func)
                st.session_state.last_processed_data = excel_data_dfs_raw
                st.download_button("처리된 엑셀 다운로드", data=styled_bytes, file_name="processed.xlsx")
                st.success("✅ 파일 분석 완료. 알림 전송 옵션을 선택하세요.")
                
                col_auto, col_manual = st.columns(2)
                if col_auto.button("YES: 자동 전송"): st.session_state.auto_run_confirmed = True; st.rerun()
                if col_manual.button("NO: 수동 선택"): st.session_state.auto_run_confirmed = False; st.rerun()

                if st.session_state.auto_run_confirmed is not None:
                    all_u, all_p, all_d = users_ref.get(), db_ref_func("patients").get(), doctor_users_ref.get()
                    m_u, m_d = get_matching_data(excel_data_dfs_raw, all_u, all_p, all_d)
                    if st.session_state.auto_run_confirmed:
                        run_auto_notifications(m_u, m_d, excel_data_dfs_raw, file_name, is_daily, db_ref_func)
                        st.session_state.auto_run_confirmed = None; st.stop()
            except Exception as e: st.error(f"오류: {e}")

    with tab_user_mgmt:
        if not st.session_state.admin_password_correct:
            if st.button("관리자 인증"):
                st.session_state.admin_password_correct = True; st.rerun()
        else:
            st.subheader("사용자 계정 관리")
            # 사용자 삭제 및 목록 로직 (필요시 구현)

# --- 5. 사용자 및 의사 모드 UI ---

def show_user_mode_ui(firebase_key, user_name):
    """학생 모드 UI"""
    st.header(f"👋 {user_name}님 (학생 모드)")
    get_google_calendar_service(firebase_key) # PKCE 대응 인증 호출
    t1, t2, t3 = st.tabs(['📅 환자 관리', '📈 분석 결과', '🧑‍🏫 리뷰'])
    with t3: show_professor_review_system()

def show_doctor_mode_ui(firebase_key, user_name):
    """의사 모드 UI"""
    st.header(f"🧑‍⚕️ Dr. {user_name} (의사 모드)")
    get_google_calendar_service(firebase_key)
    st.info("의사 전용 기능 페이지입니다.")
