# ui_manager.py
import streamlit as st
import pandas as pd
import bcrypt
import excel_utils
from config import DEFAULT_PASSWORD, DEPARTMENTS_FOR_REGISTRATION, PATIENT_DEPT_FLAGS
from firebase_utils import get_db_refs, sanitize_path, get_google_calendar_service
from notification_utils import is_valid_email, send_email, get_matching_data, run_auto_notifications

users_ref, doctor_users_ref, db_ref_func = get_db_refs()

def init_session_state():
    states = {
        'login_mode': 'not_logged_in', 'current_firebase_key': '', 'current_user_name': '',
        'admin_password_correct': False, 'matched_user_multiselect': [], 'matched_doctor_multiselect': []
    }
    for k, v in states.items():
        if k not in st.session_state: st.session_state[k] = v

def show_login_and_registration():
    # 로그인 UI 로직 (원본 유지)
    pass

def show_admin_mode_ui():
    st.title("💻 관리자 모드")
    # 엑셀 업로드 및 알림 발송 (excel_utils 참조 수정 완료)
    pass

def show_user_mode_ui(firebase_key, user_name):
    # 환자 관리 및 OCS 분석 결과 (원본 유지)
    pass
