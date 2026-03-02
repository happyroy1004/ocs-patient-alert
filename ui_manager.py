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
# professor_reviews_module은 실제 파일이 존재하는지 확인 필요
try:
    from professor_reviews_module import show_professor_review_system 
except ImportError:
    def show_professor_review_system(): st.error("리뷰 모듈을 찾을 수 없습니다.")

# DB 초기화
users_ref, doctor_users_ref, db_ref_func = get_db_refs()

def hash_password(password):
    salt = bcrypt.gensalt()
    return bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')

def check_password(password, hashed_password):
    if not hashed_password or not isinstance(hashed_password, str): return False
    try:
        return bcrypt.checkpw(password.encode('utf-8'), hashed_password.encode('utf-8'))
    except: return False

def init_session_state():
    defaults = {
        'login_mode': 'not_logged_in', 'current_user_name': '', 'current_firebase_key': '',
        'current_user_role': 'user', 'admin_password_correct': False,
        'matched_user_multiselect': [], 'matched_doctor_multiselect': []
    }
    for key, val in defaults.items():
        if key not in st.session_state: st.session_state[key] = val

def show_title_and_manual():
    st.title("환자 내원 확인 시스템")
    st.caption("directed by HSY")
    if os.path.exists("manual.pdf"):
        with open("manual.pdf", "rb") as f:
            st.download_button("사용 설명서 다운로드", f, "manual.pdf")

# ... (기존 _handle_user_login, _handle_doctor_login, show_login_and_registration 로직 유지) ...
# (중략: 기존에 제공해주신 UI 로직들은 문법적으로 큰 문제가 없으므로 그대로 포함하시면 됩니다.)

def show_admin_mode_ui():
    st.title("💻 관리자 모드")
    # 관리자 기능 구현부 (제공된 코드와 동일)
    # ...

def show_user_mode_ui(firebase_key, user_name):
    st.subheader(f"👋 {user_name}님 환영합니다.")
    # 일반 사용자 탭 구현부
    # ...

def show_doctor_mode_ui(firebase_key, user_name):
    st.header(f"🧑‍⚕️ Dr. {user_name}")
    # 의사 모드 탭 구현부
    # ...
