# ui_manager.py (신규 등록 번호 입력 및 엑셀 암호화 처리 완전판)

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
    st.error("🚨 Firebase 연결 실패. Secrets 설정을 확인하세요.")
    st.stop()

# --- 1. 보안 유틸리티 ---

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

# --- 2. 세션 상태 초기화 ---

def init_session_state():
    """앱의 모든 세션 상태를 초기화하여 'no attribute' 에러를 방지합니다."""
    defaults = {
        'login_mode': 'not_logged_in',
        'found_user_email': '',
        'current_firebase_key': '',
        'current_user_name': '',
        'admin_password_correct': False,
        'auto_run_confirmed': None, # 💡 auto_run_confirmed 누락 에러 방지
        'current_user_role': 'user',
        'matched_user_multiselect': [],
        'matched_doctor_multiselect': [],
        'delete_patient_confirm': False,
        'patients_to_delete': []
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

def show_title_and_manual():
    st.markdown("<h1>환자 내원 확인 시스템</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: grey; font-size: small;'>directed by HSY</p>", unsafe_allow_html=True)
    if os.path.exists("manual.pdf"):
        with open("manual.pdf", "rb") as f:
            st.download_button("📘 사용 설명서 다운로드", f, file_name="manual.pdf")

# --- 3. 로그인 및 등록 처리 ---

def _handle_user_login(user_name, password_input):
    if not user_name:
        st.error("이름을 입력해주세요.")
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
        else: st.error("비밀번호 불일치")
    else:
        st.session_state.update({'found_user_email': doctor_email, 'login_mode': 'new_doctor_registration'})
        st.rerun()

def show_login_and_registration():
    mode = st.session_state.login_mode
    if mode == 'not_logged_in':
        t1, t2 = st.tabs(["학생", "의사"])
        with t1:
            u = st.text_input("학생 성함", key="l_u")
            p = st.text_input("비밀번호", type="password", key="l_p")
            if st.button("학생 로그인/등록"): _handle_user_login(u, p)
        with t2:
            e = st.text_input("의사 이메일", key="l_e")
            p2 = st.text_input("비밀번호", type="password", key="l_p2")
            if st.button("의사 로그인/등록"): _handle_doctor_login(e, p2)

    elif mode == 'new_user_registration':
        st.subheader("👨‍🎓 학생 신규 등록")
        email = st.text_input("이메일(ID)", key="r_u_e")
        number = st.text_input("원내생 번호 (예: 12)", key="r_u_n") # 💡 번호 필드 추가
        pw = st.text_input("비밀번호", type="password", key="r_u_p")
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
        num = st.text_input("식별 번호", key="r_d_num") # 💡 번호 필드 추가
        dept = st.selectbox("소속 과", DEPARTMENTS_FOR_REGISTRATION)
        pw = st.text_input("비밀번호", type="password", key="r_d_p", value=DEFAULT_PASSWORD)
        if st.button("의사 등록 완료"):
            key = sanitize_path(st.session_state.found_user_email)
            doctor_users_ref.child(key).set({
                "name": name, "email": st.session_state.found_user_email,
                "number": num, "password": hash_password(pw), 
                "department": dept, "role": "doctor"
            })
            st.session_state.update({'current_firebase_key': key, 'current_user_name': name, 'login_mode': 'doctor_mode'})
            st.rerun()

# --- 4. 관리자 모드 UI (엑셀 비밀번호 처리 로직 포함) ---

def show_admin_mode_ui():
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
            
            # 💡 엑셀 암호화 여부 체크 및 입력 받기
            password = None
            if excel_utils.is_encrypted_excel(uploaded_file):
                password = st.text_input("⚠️ 암호화된 파일입니다. 엑셀 비밀번호를 입력해주세요.", type="password", key="excel_file_pw")
                if not password:
                    st.info("비밀번호를 입력해야 분석을 시작할 수 있습니다.")
                    st.stop()

            try:
                # 💡 입력된 비밀번호를 사용하여 엑셀 로드
                xl_obj, raw_io = excel_utils.load_excel(uploaded_file, password)
                excel_data_dfs_raw, styled_bytes = excel_utils.process_excel_file_and_style(raw_io, db_ref_func)
                st.session_state.last_processed_data = excel_data_dfs_raw
                
                st.success("✅ 파일 분석 및 로드 성공!")
                st.download_button("처리된 엑셀 다운로드", data=styled_bytes, file_name=f"processed_{file_name}")
                
                st.markdown("---")
                st.subheader("🚀 알림 전송 옵션")
                col_auto, col_manual = st.columns(2)
                if col_auto.button("YES: 모든 사용자 자동 전송"): st.session_state.auto_run_confirmed = True; st.rerun()
                if col_manual.button("NO: 수동으로 사용자 선택"): st.session_state.auto_run_confirmed = False; st.rerun()

                if st.session_state.auto_run_confirmed is not None:
                    all_u, all_p, all_d = users_ref.get(), db_ref_func("patients").get(), doctor_users_ref.get()
                    m_u, m_d = get_matching_data(excel_data_dfs_raw, all_u, all_p, all_d)
                    if st.session_state.auto_run_confirmed:
                        run_auto_notifications(m_u, m_d, excel_data_dfs_raw, file_name, is_daily, db_ref_func)
                        st.session_state.auto_run_confirmed = None
                        st.success("알림 전송 완료!")
            except Exception as e:
                st.error(f"❌ 오류 발생: {e}")

    with tab_user_mgmt:
        if not st.session_state.admin_password_correct:
            admin_pw_input = st.text_input("관리자 인증 비밀번호", type="password", key="admin_auth_pw")
            if st.button("인증"):
                if admin_pw_input == st.secrets.get("admin", {}).get("password", DEFAULT_PASSWORD):
                    st.session_state.admin_password_correct = True; st.rerun()
                else: st.error("인증 실패")
        else:
            st.subheader("사용자 계정 관리")
            # (사용자 목록 및 삭제 로직...)

# --- 5. 사용자/의사 모드 UI ---

def show_user_mode_ui(firebase_key, user_name):
    st.header(f"👋 {user_name}님 (학생 모드)")
    get_google_calendar_service(firebase_key) # 🔑 PKCE 대응 인증 호출
    t1, t2, t3 = st.tabs(['📅 환자 관리', '📈 분석 결과', '🧑‍🏫 리뷰'])
    with t3: show_professor_review_system()

def show_doctor_mode_ui(firebase_key, user_name):
    st.header(f"🧑‍⚕️ Dr. {user_name} (의사 모드)")
    get_google_calendar_service(firebase_key)
    st.info("의사 전용 기능 페이지입니다.")
