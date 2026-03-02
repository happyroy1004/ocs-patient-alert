# ui_manager.py (전체 코드)

import streamlit as st
import pandas as pd
import datetime
import os
import re
import bcrypt
from googleapiclient.discovery import build

from config import (
    DEFAULT_PASSWORD, DEPARTMENTS_FOR_REGISTRATION, PATIENT_DEPT_FLAGS
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
    st.error("🚨 Firebase 초기화 실패")
    st.stop()

def hash_password(password):
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

def check_password(password, hashed):
    if not hashed: return False
    try: return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))
    except: return False

def init_session_state():
    """앱에 필요한 모든 세션 상태를 초기화합니다."""
    # 💡 런타임 에러 방지를 위한 핵심 초기화 로직
    defaults = {
        'login_mode': 'not_logged_in',
        'current_user_name': '',
        'current_firebase_key': '',
        'current_user_role': 'user',
        'admin_password_correct': False,
        'auto_run_confirmed': None,  # 👈 이 부분이 누락되면 에러가 납니다.
        'matched_user_multiselect': [],
        'matched_doctor_multiselect': []
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

def show_title_and_manual():
    st.title("환자 내원 확인 시스템")
    st.caption("directed by HSY")
    if os.path.exists("manual.pdf"):
        with open("manual.pdf", "rb") as f:
            st.download_button("📘 사용 설명서 다운로드", f, file_name="manual.pdf")

def _handle_user_login(user_name, password_input):
    if not user_name: return
    if user_name.strip().lower() == "admin":
        st.session_state.login_mode = 'admin_mode'
        st.rerun()
    
    all_users = users_ref.get()
    matched_user = next((info for k, info in all_users.items() if info and info.get("name") == user_name), None)
    safe_key = next((k for k, info in all_users.items() if info and info.get("name") == user_name), None)

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
    mode = st.session_state.login_mode
    if mode == 'not_logged_in':
        t1, t2 = st.tabs(["학생", "의사"])
        with t1:
            u = st.text_input("학생 이름")
            p = st.text_input("비밀번호", type="password")
            if st.button("학생 로그인"): _handle_user_login(u, p)
    
    elif mode == 'new_user_registration':
        st.subheader("👨‍🎓 학생 신규 등록")
        email = st.text_input("이메일(ID)")
        num = st.text_input("원내생 번호 (예: 12)")
        pw = st.text_input("새 비밀번호", type="password")
        if st.button("등록 완료"):
            key = sanitize_path(email)
            users_ref.child(key).set({
                "name": st.session_state.current_user_name, 
                "email": email, "number": num, "password": hash_password(pw)
            })
            st.session_state.update({'current_firebase_key': key, 'login_mode': 'user_mode'})
            st.rerun()

def show_admin_mode_ui():
    st.title("💻 관리자 모드")
    uploaded_file = st.file_uploader("OCS Excel 업로드", type=["xlsx", "xlsm"])
    
    if uploaded_file:
        file_name = uploaded_file.name
        is_daily = excel_utils.is_daily_schedule(file_name)
        
        try:
            xl_obj, raw_io = excel_utils.load_excel(uploaded_file, None)
            excel_data_dfs, styled_bytes = excel_utils.process_excel_file_and_style(raw_io, db_ref_func)
            st.session_state.last_processed_data = excel_data_dfs
            
            st.success("분석 완료. 전송 옵션을 선택하세요.")
            col1, col2 = st.columns(2)
            if col1.button("전체 자동 전송"): 
                st.session_state.auto_run_confirmed = True
                st.rerun()
            if col2.button("수동 선택"): 
                st.session_state.auto_run_confirmed = False
                st.rerun()
                
            if st.session_state.auto_run_confirmed is True:
                # 자동 알림 로직 호출...
                pass
        except Exception as e:
            st.error(f"오류: {e}")

def show_user_mode_ui(firebase_key, user_name):
    st.header(f"👋 {user_name}님")
    get_google_calendar_service(firebase_key) # 🔑 PKCE 대응 인증 호출
    t1, t2, t3 = st.tabs(['환자 관리', '분석 결과', '리뷰'])
    with t3: show_professor_review_system()

def show_doctor_mode_ui(firebase_key, user_name):
    st.header(f"🧑‍⚕️ Dr. {user_name}")
    get_google_calendar_service(firebase_key)
