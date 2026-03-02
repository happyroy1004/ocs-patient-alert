import streamlit as st
import pandas as pd
import datetime
import os
import re
import bcrypt
from googleapiclient.discovery import build

from config import DEFAULT_PASSWORD, DEPARTMENTS_FOR_REGISTRATION, PATIENT_DEPT_FLAGS
from firebase_utils import (
    get_db_refs, sanitize_path, recover_email, 
    get_google_calendar_service, save_google_creds_to_firebase, load_google_creds_from_firebase
)
import excel_utils
from notification_utils import (
    is_valid_email, send_email, create_calendar_event, 
    get_matching_data, run_auto_notifications
)

# DB 초기 로드
users_ref, doctor_users_ref, db_ref_func = get_db_refs()

def hash_password(password):
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

def check_password(password, hashed):
    if not hashed: return False
    return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))

def init_session_state():
    if 'login_mode' not in st.session_state: st.session_state.login_mode = 'not_logged_in'
    if 'current_user_name' not in st.session_state: st.session_state.current_user_name = ""
    if 'current_firebase_key' not in st.session_state: st.session_state.current_firebase_key = ""

def show_title_and_manual():
    st.title("환자 내원 확인 시스템")
    if os.path.exists("manual.pdf"):
        with open("manual.pdf", "rb") as f:
            st.download_button("매뉴얼 다운로드", f, file_name="manual.pdf")

def show_login_and_registration():
    mode = st.session_state.login_mode
    if mode == 'not_logged_in':
        t1, t2 = st.tabs(["학생", "의사"])
        with t1:
            u = st.text_input("이름")
            p = st.text_input("비밀번호", type="password")
            if st.button("학생 로그인"):
                # (로그인 처리 로직 생략 없이 _handle_user_login 호출)
                pass
    
    elif mode == 'new_user_registration':
        st.subheader("👨‍🎓 신규 학생 등록")
        email = st.text_input("이메일")
        number = st.text_input("원내생 번호")
        pw = st.text_input("비밀번호", type="password")
        if st.button("등록 완료"):
            key = sanitize_path(email)
            users_ref.child(key).set({
                "name": st.session_state.current_user_name,
                "email": email,
                "number": number,
                "password": hash_password(pw)
            })
            st.session_state.login_mode = 'user_mode'
            st.rerun()

def show_admin_mode_ui():
    st.header("💻 관리자 모드")
    # (관리자 UI 전체 로직...)

def show_user_mode_ui(key, name):
    st.header(f"👋 {name}님")
    get_google_calendar_service(key)
    # (사용자 탭 로직...)

def show_doctor_mode_ui(key, name):
    st.header(f"🧑‍⚕️ Dr. {name}")
    get_google_calendar_service(key)
