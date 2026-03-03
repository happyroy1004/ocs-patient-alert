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

# --- 0. 보안 및 유틸리티 함수 ---

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
    """통합된 google_calendar_creds 노드를 체크하여 상단에 표시합니다."""
    is_connected, msg = check_google_connection_status(safe_key)
    
    st.markdown("### 🌐 시스템 연결 상태")
    if is_connected:
        st.success(f"✅ **구글 캘린더 연동 정상** ({msg})")
        if st.session_state.get('google_calendar_service') is None:
            get_google_calendar_service(safe_key)
        
        with st.expander("연동 설정 관리"):
            if st.button("🔄 인증 정보 갱신 (재연동)"):
                get_google_calendar_service(safe_key)
    else:
        st.error(f"❌ **구글 캘린더 미연동** ({msg})")
        if st.button("🔗 구글 계정 연동하기", key="btn_auth_sync_main"):
            get_google_calendar_service(safe_key)

# --- 1. 세션 상태 및 공통 UI ---

def init_session_state():
    """앱 구동에 필요한 모든 세션 상태 초기화"""
    if 'login_mode' not in st.session_state: st.session_state.login_mode = 'not_logged_in'
    if 'current_firebase_key' not in st.session_state: st.session_state.current_firebase_key = ""
    if 'current_user_name' not in st.session_state: st.session_state.current_user_name = ""
    if 'admin_password_correct' not in st.session_state: st.session_state.admin_password_correct = False
    if 'google_calendar_service' not in st.session_state: st.session_state.google_calendar_service = None
    if 'auto_run_confirmed' not in st.session_state: st.session_state.auto_run_confirmed = None
    if 'matched_user_multiselect' not in st.session_state: st.session_state.matched_user_multiselect = []
    if 'matched_doctor_multiselect' not in st.session_state: st.session_state.matched_doctor_multiselect = []
    if 'last_processed_data' not in st.session_state: st.session_state.last_processed_data = None
    if 'delete_patient_confirm' not in st.session_state: st.session_state.delete_patient_confirm = False

def show_title_and_manual():
    st.markdown("<h1>🦷 환자 내원 확인 시스템</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: grey; font-size: 0.8em;'>directed by HSY</p>", unsafe_allow_html=True)
    if os.path.exists("manual.pdf"):
        with open("manual.pdf", "rb") as f:
            st.download_button("📘 사용 설명서 다운로드", f, file_name="manual.pdf", mime="application/pdf")

# --- 2. 로그인 및 등록 로직 ---

def _handle_user_login(user_name, password_input):
    if not user_name:
        st.error("이름을 입력해주세요.")
        return
    if user_name.lower() == "admin":
        st.session_state.login_mode = 'admin_mode'
        st.rerun()

    all_users = users_ref.get()
    if all_users:
        for k, v in all_users.items():
            if v.get("name") == user_name:
                if check_password(password_input, v.get("password")):
                    st.session_state.update({'current_firebase_key': k, 'current_user_name': user_name, 'login_mode': 'user_mode'})
                    st.rerun()
                else:
                    st.error("비밀번호가 일치하지 않습니다.")
                    return
    st.session_state.update({'current_user_name': user_name, 'login_mode': 'new_user_registration'})
    st.rerun()

def show_login_and_registration():
    if st.session_state.login_mode == 'not_logged_in':
        tab1, tab2 = st.tabs(["👨‍🎓 학생 로그인", "🧑‍⚕️ 치과의사 로그인"])
        with tab1:
            u_input = st.text_input("성함 (실명)", key="student_login_name")
            p_input = st.text_input("비밀번호", type="password", key="student_login_pw")
            if st.button("로그인 / 등록", key="btn_student_login"):
                _handle_user_login(u_input, p_input)
        with tab2:
            st.info("치과의사 로그인 기능은 현재 이메일 인증 기반으로 준비 중입니다.")

    elif st.session_state.login_mode == 'new_user_registration':
        st.subheader("🆕 신규 사용자 등록")
        new_email = st.text_input("아이디(이메일)")
        new_pw = st.text_input("비밀번호 설정", type="password")
        if st.button("등록 완료"):
            if is_valid_email(new_email) and new_pw:
                new_key = sanitize_path(new_email)
                users_ref.child(new_key).set({"name": st.session_state.current_user_name, "email": new_email, "password": hash_password(new_pw)})
                st.session_state.update({'current_firebase_key': new_key, 'login_mode': 'user_mode'})
                st.success("등록 성공!"); st.rerun()
            else: st.error("정보를 정확히 입력해주세요.")

# --- 3. 관리자 모드 UI (엑셀 & 알림) ---

def show_admin_mode_ui():
    st.title("💻 관리자 제어 센터")
    tab_excel, tab_manage = st.tabs(["📊 OCS 파일 처리 및 알림", "👥 사용자 관리"])
    
    with tab_excel:
        uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])
        if uploaded_file:
            file_name = uploaded_file.name
            is_daily = excel_utils.is_daily_schedule(file_name)
            password = st.text_input("파일 비밀번호", type="password") if excel_utils.is_encrypted_excel(uploaded_file) else None
            
            if st.button("파일 분석 및 스타일 적용"):
                try:
                    xl_obj, raw_io = excel_utils.load_excel(uploaded_file, password)
                    clean_dfs, styled_bytes = excel_utils.process_excel_file_and_style(raw_io)
                    st.session_state.last_processed_data = clean_dfs
                    analysis = excel_utils.run_analysis(clean_dfs)
                    db_ref_func("ocs_analysis/latest_result").set(analysis)
                    db_ref_func("ocs_analysis/latest_file_name").set(file_name)
                    st.download_button("📊 처리된 엑셀 다운로드", styled_bytes, file_name=f"processed_{file_name}")
                    st.success("분석 완료! 아래에서 알림 전송을 선택하세요.")
                except Exception as e: st.error(f"오류: {e}")

            if st.session_state.last_processed_data:
                st.divider()
                st.subheader("🚀 알림 전송")
                col1, col2 = st.columns(2)
                if col1.button("YES: 자동 전체 전송"): st.session_state.auto_run_confirmed = True
                if col2.button("NO: 수동 선택 전송"): st.session_state.auto_run_confirmed = False

                if st.session_state.auto_run_confirmed is True:
                    m_users, m_docs = get_matching_data(st.session_state.last_processed_data, users_ref.get(), db_ref_func("patients").get(), doctor_users_ref.get())
                    run_auto_notifications(m_users, m_docs, st.session_state.last_processed_data, db_ref_func("ocs_analysis/latest_file_name").get(), is_daily, db_ref_func)
                    st.session_state.auto_run_confirmed = None
                elif st.session_state.auto_run_confirmed is False:
                    st.info("수동 전송 모드입니다. 상세 로직을 구현하여 사용하세요.")

# --- 4. 사용자(학생) 모드 UI ---

def show_user_mode_ui(firebase_key, user_name):
    show_connection_status_widget(firebase_key)
    st.divider()
    tab_reg, tab_anal = st.tabs(["📋 환자 관리", "📊 분석 결과"])
    
    with tab_reg:
        p_ref = db_ref_func(f"patients/{firebase_key}")
        st.subheader("➕ 환자 신규 등록")
        with st.form("reg_patient"):
            p_name = st.text_input("환자명"); p_id = st.text_input("진료번호")
            depts = st.multiselect("진료과", DEPARTMENTS_FOR_REGISTRATION)
            if st.form_submit_button("등록"):
                if p_name and p_id and depts:
                    data = {"환자이름": p_name, "진료번호": p_id.zfill(8)}
                    for d in PATIENT_DEPT_FLAGS: data[d.lower()] = d in depts
                    p_ref.child(p_id.zfill(8)).set(data)
                    st.success("등록 완료"); st.rerun()

        st.divider()
        st.subheader("📂 내 환자 목록")
        patients = p_ref.get()
        if patients:
            for pid, p_info in patients.items():
                with st.container(border=True):
                    c1, c2 = st.columns([4, 1])
                    c1.write(f"**{p_info.get('환자이름')}** ({pid})")
                    if c2.button("삭제", key=f"del_{pid}"):
                        p_ref.child(pid).delete(); st.rerun()
        else: st.info("등록된 환자가 없습니다.")

    with tab_anal:
        res = db_ref_func("ocs_analysis/latest_result").get()
        fname = db_ref_func("ocs_analysis/latest_file_name").get()
        if res:
            st.write(f"### 📑 {fname}")
            st.json(res)
        else: st.info("분석 결과가 없습니다.")

def show_doctor_mode_ui(firebase_key, user_name):
    show_connection_status_widget(firebase_key)
    st.header(f"🧑‍⚕️ Dr. {user_name} 진료실")
    st.write("의사 전용 기능 및 캘린더 동기화 설정이 표시됩니다.")
