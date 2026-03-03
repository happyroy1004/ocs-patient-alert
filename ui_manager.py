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

# local imports
from config import (
    DEFAULT_PASSWORD, DEPARTMENTS_FOR_REGISTRATION, PATIENT_DEPT_FLAGS, 
    SHEET_KEYWORD_TO_DEPARTMENT_MAP, PATIENT_DEPT_TO_SHEET_MAP
)

from firebase_utils import (
    get_db_refs, sanitize_path, recover_email, 
    get_google_calendar_service, save_google_creds_to_firebase, 
    load_google_creds_from_firebase # 이름 확인!
)

# [핵심 수정] 반환값 3개를 정확히 언패킹해야 에러가 나지 않습니다.
users_ref, doctor_users_ref, db_ref_func = get_db_refs()

import excel_utils
from notification_utils import (
    is_valid_email, send_email, create_calendar_event, 
    get_matching_data, run_auto_notifications
)

users_ref, doctor_users_ref, db_ref_func = get_db_refs()

# --- 비밀번호 보안 관련 ---
def hash_password(password):
    salt = bcrypt.gensalt()
    return bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')

def check_password(password, hashed_password):
    if not hashed_password or not isinstance(hashed_password, str):
        return False
    try:
        return bcrypt.checkpw(password.encode('utf-8'), hashed_password.encode('utf-8'))
    except Exception:
        return False

# --- 1. 세션 상태 초기화 ---
def init_session_state():
    if 'login_mode' not in st.session_state: st.session_state.login_mode = 'not_logged_in'
    if 'current_firebase_key' not in st.session_state: st.session_state.current_firebase_key = ""
    if 'current_user_name' not in st.session_state: st.session_state.current_user_name = ""
    if 'admin_password_correct' not in st.session_state: st.session_state.admin_password_correct = False
    if 'auto_run_confirmed' not in st.session_state: st.session_state.auto_run_confirmed = None 
    if 'matched_user_multiselect' not in st.session_state: st.session_state.matched_user_multiselect = []
    if 'matched_doctor_multiselect' not in st.session_state: st.session_state.matched_doctor_multiselect = []

def show_title_and_manual():
    st.markdown("<h1>환자 내원 확인 시스템</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: grey;'>directed by HSY</p>", unsafe_allow_html=True)

# --- 2. 로그인 및 등록 UI ---
def _handle_user_login(user_name, password_input):
    if not user_name: 
        st.error("사용자 이름을 입력해주세요.")
        return
    
    all_users_meta = users_ref.get()
    if all_users_meta:
        for safe_key, info in all_users_meta.items():
            if info.get("name") == user_name:
                if check_password(password_input, info.get("password")):
                    st.session_state.update({
                        'current_firebase_key': safe_key, 
                        'current_user_name': user_name, 
                        'login_mode': 'user_mode'
                    })
                    st.rerun()
                else: st.error("비밀번호 불일치")
                return
    
    st.session_state.current_user_name = user_name
    st.session_state.login_mode = 'new_user_registration'
    st.rerun()

def show_login_and_registration():
    if st.session_state.login_mode == 'not_logged_in':
        user_name = st.text_input("학생 이름")
        password = st.text_input("비밀번호", type="password")
        if st.button("로그인"):
            _handle_user_login(user_name, password)
    
    elif st.session_state.login_mode == 'new_user_registration':
        st.subheader("신규 등록")
        email = st.text_input("이메일(ID)")
        pw = st.text_input("비밀번호", type="password")
        if st.button("등록 완료"):
            safe_key = sanitize_path(email)
            users_ref.child(safe_key).set({"name": st.session_state.current_user_name, "email": email, "password": hash_password(pw)})
            st.session_state.update({'current_firebase_key': safe_key, 'login_mode': 'user_mode'})
            st.rerun()

# --- 3. 관리자 모드 UI ---
def show_admin_mode_ui():
    st.title("💻 관리자 모드")
    uploaded_file = st.file_uploader("OCS 엑셀 업로드", type=["xlsx", "xlsm"])
    
    if uploaded_file:
        # 파일 처리 로직 (생략 방지를 위해 핵심만 포함)
        is_daily = excel_utils.is_daily_schedule(uploaded_file.name)
        xl_data, styled_file = excel_utils.process_excel_file_and_style(uploaded_file)
        
        if st.button("자동 알림 전송"):
            matched_users, matched_docs = get_matching_data(xl_data, users_ref.get(), db_ref_func("patients").get(), doctor_users_ref.get())
            run_auto_notifications(matched_users, matched_docs, xl_data, uploaded_file.name, is_daily, db_ref_func)
            st.success("알림 전송 완료")

# --- 4. 일반 사용자 모드 UI (수정 핵심) ---
def show_user_mode_ui(firebase_key, user_name):
    # [입구 통제] 환자 등록 시 사용되는 레퍼런스를 safe_key 기반으로 고정합니다.
    patients_ref = db_ref_func(f"patients/{firebase_key}")
    
    st.subheader(f"{user_name}님 환영합니다.")
    
    # 구글 캘린더 연동 상태 표시
    get_google_calendar_service(firebase_key)
    
    tab_reg, tab_list = st.tabs(["환자 등록", "목록 관리"])
    
    with tab_reg:
        with st.form("reg_form"):
            p_name = st.text_input("환자 이름")
            p_id = st.text_input("진료 번호")
            depts = st.multiselect("진료과", DEPARTMENTS_FOR_REGISTRATION)
            if st.form_submit_button("등록"):
                # 모든 진료과 플래그 초기화 및 선택된 과만 True 설정
                p_data = {"환자이름": p_name, "진료번호": p_id}
                for d in PATIENT_DEPT_FLAGS: p_data[d.lower()] = (d in depts)
                patients_ref.child(p_id).set(p_data)
                st.success("등록 완료")
                st.rerun()

    with tab_list:
        data = patients_ref.get()
        if data:
            for pid, val in data.items():
                col1, col2 = st.columns([4, 1])
                col1.write(f"{val.get('환자이름')} ({pid})")
                if col2.button("삭제", key=f"del_{pid}"):
                    patients_ref.child(pid).delete()
                    st.rerun()
