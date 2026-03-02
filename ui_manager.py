# ui_manager.py

import streamlit as st
import pandas as pd
import io
import datetime
from googleapiclient.discovery import build
import os
import re
import bcrypt
import json # json 임포트 추가 (secrets 처리에 필요)

# local imports: 상대 경로 임포트(.)를 절대 경로 임포트로 수정
from config import (
    DEFAULT_PASSWORD, DEPARTMENTS_FOR_REGISTRATION, PATIENT_DEPT_FLAGS, 
    SHEET_KEYWORD_TO_DEPARTMENT_MAP, PATIENT_DEPT_TO_SHEET_MAP
)
from firebase_utils import (
    get_db_refs, sanitize_path, recover_email, 
    get_google_calendar_service, save_google_creds_to_firebase, load_google_creds_from_firebase
)

# 💡 수정: excel_utils 전체를 import하여 순환 참조 문제를 회피
import excel_utils
from notification_utils import (
    is_valid_email, send_email, create_calendar_event, 
    get_matching_data, run_auto_notifications
)

# DB 레퍼런스 초기 로드 (전역에서 사용할 수 있도록 설정)
# @st.cache_resource 덕분에 앱 시작 시 단 한번 안전하게 초기화됩니다.
users_ref, doctor_users_ref, db_ref_func = get_db_refs()

# 🔑 비밀번호 암호화 및 확인 유틸리티 함수
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
    except ValueError:
        return False
    except Exception:
        return False


# --- 1. 세션 상태 초기화 및 전역 UI ---

def init_session_state():
    """앱에 필요한 모든 세션 상태를 초기화합니다."""
    # Note: 이 함수는 streamlit_app.py에서 호출되어야 합니다.
    if 'login_mode' not in st.session_state: st.session_state.login_mode = 'not_logged_in'
    if 'email_change_mode' not in st.session_state: st.session_state.email_change_mode = False
    if 'user_id_input_value' not in st.session_state: st.session_state.user_id_input_value = ""
    if 'found_user_email' not in st.session_state: st.session_state.found_user_email = ""
    if 'current_firebase_key' not in st.session_state: st.session_state.current_firebase_key = ""
    if 'current_user_name' not in st.session_state: st.session_state.current_user_name = ""
    if 'logged_in_as_admin' not in st.session_state: st.session_state.logged_in_as_admin = False
    if 'admin_password_correct' not in st.session_state: st.session_state.admin_password_correct = False
    if 'select_all_users' not in st.session_state: st.session_state.select_all_users = False
    if 'google_calendar_auth_needed' not in st.session_state: st.session_state.google_calendar_auth_needed = False
    if 'google_creds' not in st.session_state: st.session_state['google_creds'] = {}
    if 'auto_run_confirmed' not in st.session_state: st.session_state.auto_run_confirmed = None 
    if 'current_user_role' not in st.session_state: st.session_state.current_user_role = 'user'
    if 'current_user_dept' not in st.session_state: st.session_state.current_user_dept = None
    if 'delete_patient_confirm' not in st.session_state: st.session_state.delete_patient_confirm = False
    if 'patients_to_delete' not in st.session_state: st.session_state.patients_to_delete = []
    if 'select_all_mode' not in st.session_state: st.session_state.select_all_mode = False
    # Admin 모드에서 필요한 추가 세션 상태 (excel_utils에서 추출된 날짜를 저장)
    if 'reservation_date_excel' not in st.session_state: st.session_state.reservation_date_excel = "날짜_미정"
    # Admin 모드에서 필요한 추가 세션 상태 (매칭 UI에 필요)
    if 'matched_user_multiselect' not in st.session_state: st.session_state.matched_user_multiselect = []
    if 'matched_doctor_multiselect' not in st.session_state: st.session_state.matched_doctor_multiselect = []


def show_title_and_manual():
    """제목과 사용 설명서 버튼을 표시합니다."""
    st.markdown("""
        <style> .title-link { text-decoration: none; color: inherit; } </style>
        <h1> <a href="." class="title-link">환자 내원 확인 시스템</a> </h1>
    """, unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("<p style='text-align: left; color: grey; font-size: small;'>directed by HSY</p>", unsafe_allow_html=True)

    pdf_file_path = "manual.pdf"
    if os.path.exists(pdf_file_path):
        with open(pdf_file_path, "rb") as pdf_file:
            st.download_button(
                label="사용 설명서 다운로드", data=pdf_file, file_name=pdf_file_path, mime="application/pdf"
            )
    else: st.warning(f"⚠️ 사용 설명서 파일을 찾을 수 없습니다. (경로: {pdf_file_path})")

# --- 2. 로그인 및 등록 UI ---

def _handle_user_login(user_name, password_input):
    """학생 로그인 로직을 처리합니다."""
    # 💡 DB 연결 오류 방어 로직
    if users_ref is None:
        st.error("🚨 데이터베이스 연결에 문제가 있습니다. 관리자에게 문의하세요.")
        return
        
    if not user_name: st.error("사용자 이름을 입력해주세요.")
    elif user_name.strip().lower() == "admin": 
        # 'admin' 입력 시 비밀번호 없이 바로 관리자 모드 진입 (Admin 우회 접속)
        st.session_state.login_mode = 'admin_mode'; st.rerun()
    else:
        all_users_meta = users_ref.get()
        matched_user = None
        safe_key_found = None

        if all_users_meta:
            for safe_key, user_info in all_users_meta.items():
                if user_info and user_info.get("name") == user_name:
                    matched_user = user_info
                    safe_key_found = safe_key
                    break

        if matched_user:
            user_password_db = matched_user.get("password")

            # 💡 비밀번호 인증 및 마이그레이션 로직
            login_success = check_password(password_input, user_password_db)
            is_plaintext_or_default = False
            
            # 마이그레이션 로직: 저장된 비밀번호가 해시가 아닌 평문이거나 기본 비밀번호일 경우 평문 비교 시도
            if not login_success:
                if password_input == user_password_db:
                    login_success = True
                    is_plaintext_or_default = True
                elif (not user_password_db or user_password_db == DEFAULT_PASSWORD) and password_input == DEFAULT_PASSWORD:
                    login_success = True
                    is_plaintext_or_default = True
            
            if login_success:
                st.session_state.update({
                    'found_user_email': matched_user["email"], 
                    'current_firebase_key': safe_key_found, 
                    'current_user_name': user_name, 
                    'login_mode': 'user_mode'
                })
                # 🚨 평문 로그인 성공 시, 즉시 bcrypt로 해시하여 업데이트 (마이그레이션)
                if is_plaintext_or_default:
                    hashed_pw = hash_password(password_input if password_input else DEFAULT_PASSWORD)
                    users_ref.child(safe_key_found).update({"password": hashed_pw})
                    st.warning("⚠️ 보안 강화를 위해 비밀번호가 자동으로 암호화되었습니다. 다음 로그인부터는 변경된 비밀번호로 로그인됩니다.")

                st.info(f"**{user_name}**님으로 로그인되었습니다.")
                st.rerun()
            else: st.error("비밀번호가 일치하지 않습니다. 신규 등록 시 이름에 알파벳이나 숫자를 붙여주세요.")
        else:
            st.session_state.current_user_name = user_name
            st.session_state.login_mode = 'new_user_registration'
            st.rerun()

def _handle_doctor_login(doctor_email, password_input_doc):
    """치과의사 로그인 로직을 처리합니다."""
    # 💡 DB 연결 오류 방어 로직
    if doctor_users_ref is None:
        st.error("🚨 데이터베이스 연결에 문제가 있습니다. 관리자에게 문의하세요.")
        return

    if not doctor_email: st.warning("치과의사 이메일 주소를 입력해주세요.")
    else:
        safe_key = sanitize_path(doctor_email)
        matched_doctor = doctor_users_ref.child(safe_key).get()
        
        if matched_doctor:
            doctor_password_db = matched_doctor.get("password")
            
            # 💡 비밀번호 인증 및 마이그레이션 로직
            login_success = check_password(password_input_doc, doctor_password_db)
            is_plaintext_or_default = False
            
            # 마이그레이션 로직:
            if not login_success:
                if password_input_doc == doctor_password_db:
                    login_success = True
                    is_plaintext_or_default = True
                elif (not doctor_password_db or doctor_password_db == DEFAULT_PASSWORD) and password_input_doc == DEFAULT_PASSWORD:
                    login_success = True
                    is_plaintext_or_default = True

            if login_success:
                st.session_state.update({
                    'found_user_email': matched_doctor["email"], 
                    'current_firebase_key': safe_key, 
                    'current_user_name': matched_doctor.get("name"),
                    'current_user_dept': matched_doctor.get("department"),
                    'current_user_role': 'doctor',
                    'login_mode': 'doctor_mode'
                })
                # 🚨 평문 로그인 성공 시, 즉시 bcrypt로 해시하여 업데이트 (마이그레이션)
                if is_plaintext_or_default:
                    hashed_pw = hash_password(password_input_doc if password_input_doc else DEFAULT_PASSWORD)
                    doctor_users_ref.child(safe_key).update({"password": hashed_pw})
                    st.warning("⚠️ 보안 강화를 위해 비밀번호가 자동으로 암호화되었습니다. 다음 로그인부터는 변경된 비밀번호로 로그인됩니다.")

                st.info(f"치과의사 **{st.session_state.current_user_name}**님으로 로그인되었습니다.")
                st.rerun()
            else: st.error("비밀번호가 일치하지 않습니다. 다시 확인해주세요.")
        else:
            st.session_state.update({
                'found_user_email': doctor_email, 
                'current_firebase_key': "",
                'current_user_name': None,
                'current_user_role': 'doctor',
                'current_user_dept': None,
                'login_mode': 'new_doctor_registration'
            })
            if password_input_doc == DEFAULT_PASSWORD:
                st.info("💡 새로운 치과의사 계정으로 인식되었습니다. 초기 비밀번호로 등록을 진행합니다.")
            st.rerun()


def show_login_and_registration():
    """학생/치과의사 로그인 및 신규 등록 폼을 표시합니다."""
    
    if st.session_state.get('login_mode') == 'not_logged_in':
        tab1, tab2 = st.tabs(["학생 로그인", "치과의사 로그인"])

        with tab1:
            st.subheader("👨‍🎓 학생 로그인")
            user_name = st.text_input("사용자 이름을 입력하세요 (예시: 홍길동)", key="login_username_tab1")
            password_input = st.text_input("비밀번호를 입력하세요", type="password", key="login_password_tab1")
            if st.button("로그인/등록", key="login_button_tab1"):
                _handle_user_login(user_name, password_input)

        with tab2:
            st.subheader("🧑‍⚕️ 치과의사 로그인")
            doctor_email = st.text_input("치과의사 이메일 주소를 입력하세요", key="doctor_email_input_tab2")
            password_input_doc = st.text_input("비밀번호를 입력하세요", type="password", key="doctor_password_input_tab2")
            if st.button("로그인/등록", key="doctor_login_button_tab2"):
                _handle_doctor_login(doctor_email, password_input_doc)

    elif st.session_state.get('login_mode') == 'new_user_registration':
        st.info(f"'{st.session_state.current_user_name}'님은 새로운 사용자입니다. 아래에 정보를 입력하여 등록을 완료하세요.")
        st.subheader("👨‍⚕️ 신규 사용자 등록")
        new_email_input = st.text_input("아이디(이메일)를 입력하세요", key="new_user_email_input")
        password_input = st.text_input("새로운 비밀번호를 입력하세요", type="password", key="new_user_password_input")
        
        if st.button("사용자 등록 완료", key="new_user_reg_button"):
            if is_valid_email(new_email_input) and password_input:
                new_firebase_key = sanitize_path(new_email_input)
                
                # 중복 이메일 검사 및 DB 연결 방어
                if users_ref is None:
                    st.error("🚨 데이터베이스 연결 오류로 등록할 수 없습니다.")
                elif users_ref.child(new_firebase_key).get():
                    st.error("이미 등록된 이메일 주소입니다. 다른 주소를 사용해주세요.")
                else:
                    # 🔑 비밀번호를 해시하여 저장
                    hashed_pw = hash_password(password_input)

                    users_ref.child(new_firebase_key).set({
                        "name": st.session_state.current_user_name,
                        "email": new_email_input,
                        "password": hashed_pw
                    })
                    st.session_state.update({
                        'current_firebase_key': new_firebase_key, 
                        'found_user_email': new_email_input, 
                        'login_mode': 'user_mode'
                    })
                    st.success(f"새로운 사용자 **{st.session_state.current_user_name}**님 ({new_email_input}) 정보가 등록되었습니다.")
                    st.rerun()
            else: st.error("올바른 이메일 주소와 비밀번호를 입력해주세요.")

    elif st.session_state.get('login_mode') == 'new_doctor_registration':
        st.info(f"아래에 정보를 입력하여 등록을 완료하세요.")
        st.subheader("👨‍⚕️ 새로운 치과의사 등록")
        new_doctor_name_input = st.text_input("이름을 입력하세요 (원내생이라면 '홍길동95'과 같은 형태로 등록바랍니다)", key="new_doctor_name_input")
        password_input = st.text_input("새로운 비밀번호를 입력하세요", type="password", key="new_doctor_password_input", value=DEFAULT_PASSWORD)
        user_id_input = st.text_input("아이디(이메일)를 입력하세요", key="new_doctor_email_input", value=st.session_state.get('found_user_email', ''))
        department = st.selectbox("등록 과", DEPARTMENTS_FOR_REGISTRATION, key="new_doctor_dept_selectbox")

        if st.button("치과의사 등록 완료", key="new_doc_reg_button"):
            if new_doctor_name_input and is_valid_email(user_id_input) and password_input and department:
                new_firebase_key = sanitize_path(user_id_input)
                
                # DB 연결 방어
                if doctor_users_ref is None:
                    st.error("🚨 데이터베이스 연결 오류로 등록할 수 없습니다.")
                else:
                    # 🔑 비밀번호를 해시하여 저장
                    hashed_pw = hash_password(password_input)

                    doctor_users_ref.child(new_firebase_key).set({
                        "name": new_doctor_name_input, "email": user_id_input, "password": hashed_pw, 
                        "role": 'doctor', "department": department
                    })
                    st.session_state.update({
                        'current_firebase_key': new_firebase_key, 
                        'found_user_email': user_id_input, 
                        'current_user_name': new_doctor_name_input,
                        'current_user_dept': department,
                        'login_mode': 'doctor_mode'
                    })
                    st.success(f"새로운 치과의사 **{new_doctor_name_input}**님 ({user_id_input}) 정보가 등록되었습니다.")
                    st.rerun()
            else: st.error("이름, 올바른 이메일 주소, 비밀번호, 그리고 등록 과를 입력해주세요.")

# --- 3. 관리자 모드 UI (Excel 및 알림) ---

def show_admin_mode_ui():
    """관리자 모드 (엑셀 업로드, 알림 전송) UI를 표시합니다."""
    
    st.markdown("---")
    st.title("💻 관리자 모드")
    
    # DB 레퍼런스 및 Gmail 정보 로드
    db_ref = db_ref_func
    # secrets 로드 시 에러 방지용 try-except 추가
    try:
        sender = st.secrets["gmail"]["sender"]; sender_pw = st.secrets["gmail"]["app_password"]
    except KeyError:
        st.error("⚠️ secrets.toml 파일에 [gmail] 정보가 없습니다. 관리자에게 문의하세요.")
        sender = "error@example.com"; sender_pw = "none" # 더미 값 설정

    # 탭 분리: OCS 파일 처리 (비번 없이 접근) vs 사용자 관리 (비번 필요)
    tab_excel, tab_user_mgmt = st.tabs(["📊 OCS 파일 처리 및 알림", "🧑‍💻 사용자 목록 및 관리"])
    
    # -----------------------------------------------------
    # 탭 1: OCS 파일 처리 및 알림 로직 (Admin 이름 입력 후 즉시 접근 가능)
    # -----------------------------------------------------
    with tab_excel:
        st.subheader("💻 Excel File Processor")
        uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])
        
        if uploaded_file:
            file_name = uploaded_file.name; 
            # 💡 수정: excel_utils 모듈을 통해 함수 호출
            is_daily = excel_utils.is_daily_schedule(file_name) 
            
            # 1. 파일 비밀번호 처리
            password = None
            # 💡 수정: excel_utils 모듈을 통해 함수 호출
            if excel_utils.is_encrypted_excel(uploaded_file): 
                password = st.text_input("⚠️ 암호화된 파일입니다. 비밀번호를 입력해주세요.", type="password", key="admin_password_file")
                if not password: st.info("비밀번호 입력 대기 중..."); st.stop()

            # 2. 파일 처리 및 분석 실행
            try:
                # 💡 수정: excel_utils 모듈을 통해 함수 호출
                xl_object, raw_file_io = excel_utils.load_excel(uploaded_file, password)
                # excel_data_dfs_raw는 컬럼명이 표준화(공백 제거)된 DF 딕셔너리를 반환합니다.
                excel_data_dfs_raw, styled_excel_bytes = excel_utils.process_excel_file_and_style(raw_file_io)
                
                # run_analysis는 excel_utils.py에서 정의된 함수를 사용해야 하지만, 
                # 현재 ui_manager.py는 notification_utils.py의 run_auto_notifications만 참조하고 있으므로
                # run_analysis 함수를 notification_utils.py에 있다고 가정하고 호출합니다.
                # (단, 이전 코드와 달리 excel_utils 모듈에서 run_analysis를 가져와야 합니다.)
                analysis_results = excel_utils.run_analysis(excel_data_dfs_raw)
                
                # 💡 수정: 분석 결과가 유효할 때만 Firebase에 저장
                if analysis_results and any(analysis_results.values()): # 결과가 비어있지 않은지 확인
                    today_date_str = datetime.datetime.now().strftime("%Y-%m-%d")
                    db_ref("ocs_analysis/latest_result").set(analysis_results)
                    db_ref("ocs_analysis/latest_date").set(today_date_str)
                    db_ref("ocs_analysis/latest_file_name").set(file_name)
                else:
                    st.warning("⚠️ 분석 결과가 비어 있어 Firebase에 저장하지 않았습니다.")
                
                st.session_state.last_processed_data = excel_data_dfs_raw; st.session_state.last_processed_file_name = file_name

                if styled_excel_bytes:
                    output_filename = uploaded_file.name.replace(".xlsx", "_processed.xlsx").replace(".xlsm", "_processed.xlsm")
                    st.download_button("처리된 엑셀 다운로드", data=styled_excel_bytes, file_name=output_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    st.success("✅ 파일 처리 및 분석이 완료되었습니다. 이제 알림 전송 방법을 선택하세요.")
                else: st.warning("엑셀 파일 처리 중 문제가 발생했거나 처리할 데이터가 없습니다.")
                    
            except ValueError as ve: st.error(f"파일 처리 실패: {ve}"); st.stop()
            except Exception as e: st.error(f"예상치 못한 오류 발생: {e}"); st.stop()
            
            # 3. 알림 전송 옵션
            st.markdown("---")
            st.subheader("🚀 알림 전송 옵션")
            col_auto, col_manual = st.columns(2)

            with col_auto:
                if st.button("YES: 자동으로 모든 사용자에게 전송", key="auto_run_yes"):
                    st.session_state.auto_run_confirmed = True; st.rerun()
            with col_manual:
                if st.button("NO: 수동으로 사용자 선택", key="auto_run_no"):
                    st.session_state.auto_run_confirmed = False; st.rerun()
                    
            # 4. 실행 로직 분기
            if 'last_processed_data' in st.session_state and st.session_state.last_processed_data:
                
                all_users_meta = users_ref.get(); all_patients_data = db_ref("patients").get()
                all_doctors_meta = doctor_users_ref.get()
                excel_data_dfs = st.session_state.last_processed_data
                
                # 매칭 데이터 준비
                matched_users, matched_doctors_data = get_matching_data(
                    excel_data_dfs, all_users_meta, all_patients_data, all_doctors_meta
                )

                # A. 자동 실행 로직 (YES 클릭 시)
                if st.session_state.auto_run_confirmed:
                    st.markdown("---")
                    st.warning("자동으로 모든 매칭 사용자에게 알림(메일/캘린더)을 전송합니다.")
                    run_auto_notifications(matched_users, matched_doctors_data, excel_data_dfs, file_name, is_daily, db_ref_func)
                    st.session_state.auto_run_confirmed = None; st.stop() # None으로 변경하여 재실행 후 UI 리셋
                    
                # B. 수동 실행 로직 (NO 클릭 시)
                elif st.session_state.auto_run_confirmed is False:
                    st.markdown("---")
                    st.info("아래 탭에서 전송할 사용자 목록을 확인하고, 원하는 사용자에게 수동으로 알림을 전송해주세요.")

                    student_admin_tab, doctor_admin_tab = st.tabs(['📚 학생 수동 전송', '🧑‍⚕️ 치과의사 수동 전송'])
                    
                    # --- 학생 수동 전송 탭 ---
                    with student_admin_tab:
                        st.subheader("📚 학생 수동 전송 (매칭 결과)");
                        if matched_users:
                            st.success(f"매칭된 환자가 있는 **{len(matched_users)}명의 사용자**를 발견했습니다.")
                            matched_user_list_for_dropdown = [f"{user['name']} ({user['email']})" for user in matched_users]
                            
                            # 💡 수정: 버튼 클릭 시 세션 상태 토글 및 즉시 재실행 요청
                            if st.button("매칭된 사용자 모두 선택/해제", key="select_all_matched_btn"):
                                current_selection_count = len(st.session_state.matched_user_multiselect)
                                total_options_count = len(matched_user_list_for_dropdown)
                                
                                if current_selection_count == total_options_count:
                                    st.session_state.matched_user_multiselect = []
                                else:
                                    st.session_state.matched_user_multiselect = matched_user_list_for_dropdown
                                
                                st.rerun()
                            
                            # 💡 수정: 멀티셀렉트의 value를 session state로 직접 지정
                            selected_users_to_act_values = st.multiselect(
                                "액션을 취할 사용자 선택", 
                                matched_user_list_for_dropdown, 
                                default=st.session_state.matched_user_multiselect, 
                                key="matched_user_multiselect" 
                            )

                            selected_matched_users_data = [user for user in matched_users if f"{user['name']} ({user['email']})" in selected_users_to_act_values]
                            
                            for user_match_info in selected_matched_users_data:
                                st.markdown(f"**수신자:** {user_match_info['name']} ({user_match_info['email']})")
                                st.dataframe(user_match_info['data'])
                            
                            mail_col, calendar_col = st.columns(2)
                            with mail_col:
                                if st.button("선택된 사용자에게 메일 보내기", key="manual_send_mail_student"):
                                    for user_match_info in selected_matched_users_data:
                                        real_email = user_match_info['email']; df_matched = user_match_info['data']; user_name = user_match_info['name']
                                        email_cols = ['환자명', '진료번호', '예약의사', '진료내역', '예약일시', '예약시간', '등록과']
                                        df_for_mail = df_matched[[col for col in email_cols if col in df_matched.columns]]
                                        rows_as_dict = df_for_mail.to_dict('records')
                                        df_html = df_for_mail.to_html(index=False, escape=False)
                                        email_body = f"""<p>안녕하세요, {user_name}님.</p><p>{file_name} 분석 결과, 내원 예정인 환자 진료 정보입니다.</p>{df_html}<p>확인 부탁드립니다.</p>"""
                                        try: send_email(real_email, rows_as_dict, sender, sender_pw, custom_message=email_body, date_str=file_name); st.success(f"**{user_name}**님 ({real_email})에게 예약 정보 이메일 전송 완료!")
                                        except Exception as e: st.error(f"**{user_name}**님 ({real_email})에게 이메일 전송 실패: {e}")

                            with calendar_col:
                                if st.button("선택된 사용자에게 Google Calendar 일정 추가", key="manual_send_calendar_student"):
                                    for user_match_info in selected_matched_users_data:
                                        user_safe_key = user_match_info['safe_key']; user_name = user_match_info['name']; df_matched = user_match_info['data']
                                        creds = load_google_creds_from_firebase(user_safe_key) 
                                        
                                        if creds and creds.valid and not creds.expired:
                                            successful_adds = 0
                                            try:
                                                service = build('calendar', 'v3', credentials=creds)
                                                
                                                for index, row in df_matched.iterrows():
                                                    reservation_date_raw = row.get('예약일시', ''); reservation_time_raw = row.get('예약시간', '')
                                                    
                                                    if reservation_date_raw and reservation_time_raw:
                                                        try:
                                                            full_datetime_str = f"{str(reservation_date_raw).strip()} {str(reservation_time_raw).strip()}"
                                                            # 🚨 주의: 날짜 포맷이 엑셀에서 넘어올 때 일관적인지 확인 필요
                                                            reservation_datetime = datetime.datetime.strptime(full_datetime_str, '%Y/%m/%d %H:%M')
                                                            
                                                            success = create_calendar_event(service, row.get('환자명', 'N/A'), row.get('진료번호', ''), row.get('등록과', ''), reservation_datetime, row.get('예약의사', 'N/A'), row.get('진료내역', ''), is_daily)
                                                            
                                                            if success:
                                                                successful_adds += 1
                                                            
                                                        except ValueError as ve:
                                                            st.error(f"❌ [데이터 형식 오류] {user_name} (환자 {row.get('환자명')}): 날짜 포맷({full_datetime_str}) 오류: {ve}")
                                                        except Exception as api_e:
                                                            st.error(f"❌ [API/기타 오류] {user_name} (환자 {row.get('환자명')}): 일정 추가 실패: {api_e}")

                                                if successful_adds > 0:
                                                    st.success(f"**{user_name}**님의 캘린더에 총 **{successful_adds}건**의 일정을 추가했습니다.")
                                                elif successful_adds == 0:
                                                    st.warning(f"**{user_name}**님의 캘린더에 추가된 일정이 없습니다. 상세 오류 메시지를 확인하세요.")

                                            except Exception as e: 
                                                st.error(f"❌ **치명적 서비스 오류:** {user_name} (API 서비스 구축 실패): 인증 파일이나 권한을 확인하세요. (오류: {e})")
                                        
                                        else: st.warning(f"**{user_name}**님은 Google Calendar 계정이 연동되어 있지 않거나 인증이 만료되었습니다.")
                        else: st.info("매칭된 환자가 없습니다.")

                    # --- 치과의사 수동 전송 탭 ---
                    with doctor_admin_tab:
                        st.subheader("🧑‍⚕️ 치과의사 수동 전송 (매칭 결과)");
                        if matched_doctors_data:
                            st.success(f"등록된 진료가 있는 **{len(matched_doctors_data)}명의 치과의사**를 발견했습니다.")
                            doctor_list_for_multiselect = [f"{res['name']} ({res['email']})" for res in matched_doctors_data]

                            # 💡 수정: 버튼 클릭 시 세션 상태 토글 및 즉시 재실행 요청
                            if st.button("등록된 치과의사 모두 선택/해제", key="select_all_matched_res_btn"):
                                current_selection_count = len(st.session_state.matched_doctor_multiselect)
                                total_options_count = len(doctor_list_for_multiselect)

                                if current_selection_count == total_options_count:
                                    st.session_state.matched_doctor_multiselect = []
                                else:
                                    st.session_state.matched_doctor_multiselect = doctor_list_for_multiselect

                                st.rerun()

                            # 💡 수정: 멀티셀렉트의 value를 session state로 직접 지정
                            selected_doctors_str = st.multiselect(
                                "액션을 취할 치과의사 선택", 
                                doctor_list_for_multiselect, 
                                default=st.session_state.matched_doctor_multiselect, 
                                key="matched_doctor_multiselect" 
                            )
                            selected_doctors_to_act = [res for res in matched_doctors_data if f"{res['name']} ({res['email']})" in selected_doctors_str]
                            
                            for res in selected_doctors_to_act:
                                st.markdown(f"**수신자:** Dr. {res['name']} ({res['email']})")
                                st.dataframe(res['data'])

                            mail_col_doc, calendar_col_doc = st.columns(2)
                            with mail_col_doc:
                                if st.button("선택된 치과의사에게 메일 보내기", key="manual_send_mail_doctor"):
                                    for res in selected_doctors_to_act:
                                        df_matched = res['data']; latest_file_name = db_ref("ocs_analysis/latest_file_name").get()
                                        email_cols = ['환자명', '진료번호', '예약의사', '진료내역', '예약일시', '예약시간']; 
                                        df_for_mail = df_matched[[col for col in email_cols if col in df_matched.columns]]
                                        df_html = df_for_mail.to_html(index=False, border=1); rows_as_dict = df_for_mail.to_dict('records')
                                        email_body = f"""<p>안녕하세요, {res['name']} 치과의사님.</p><p>{latest_file_name}에서 가져온 내원할 환자 정보입니다.</p>{df_html}<p>확인 부탁드립니다.</p>"""
                                        try: send_email(res['email'], rows_as_dict, sender, sender_pw, custom_message=email_body, date_str=latest_file_name); st.success(f"**Dr. {res['name']}**에게 메일 전송 완료!")
                                        except Exception as e: st.error(f"**Dr. {res['name']}**에게 메일 전송 실패: {e}")

                            with calendar_col_doc:
                                if st.button("선택된 치과의사에게 Google Calendar 일정 추가", key="manual_send_calendar_doctor"):
                                    for res in selected_doctors_to_act:
                                        user_safe_key = res['safe_key']; user_name = res['name']; df_matched = res['data']
                                        creds = load_google_creds_from_firebase(user_safe_key) 
                                        
                                        if creds and creds.valid and not creds.expired:
                                            successful_adds = 0
                                            try:
                                                service = build('calendar', 'v3', credentials=creds)
                                                
                                                for index, row in df_matched.iterrows():
                                                    reservation_date_raw = row.get('예약일시', ''); reservation_time_raw = row.get('예약시간', '')
                                                    
                                                    if reservation_date_raw and reservation_time_raw:
                                                        try:
                                                            full_datetime_str = f"{str(reservation_date_raw).strip()} {str(reservation_time_raw).strip()}"
                                                            reservation_datetime = datetime.datetime.strptime(full_datetime_str, '%Y/%m/%d %H:%M')
                                                            
                                                            success = create_calendar_event(service, row.get('환자명', 'N/A'), row.get('진료번호', ''), res.get('department', 'N/A'), reservation_datetime, row.get('예약의사', 'N/A'), row.get('진료내역', ''), is_daily)
                                                            
                                                            if success:
                                                                successful_adds += 1
                                                            
                                                        except ValueError as ve:
                                                            st.error(f"❌ [데이터 형식 오류] Dr. {user_name} (환자 {row.get('환자명')}): 날짜 포맷({full_datetime_str}) 오류: {ve}")
                                                        except Exception as api_e:
                                                            st.error(f"❌ [API/기타 오류] Dr. {user_name} (환자 {row.get('환자명')}): 일정 추가 실패: {api_e}")

                                                if successful_adds > 0:
                                                    st.success(f"**Dr. {user_name}**님의 캘린더에 총 **{successful_adds}건**의 일정을 추가했습니다.")
                                                elif successful_adds == 0:
                                                    st.warning(f"**Dr. {user_name}**님의 캘린더에 추가된 일정이 없습니다. 상세 오류 메시지를 확인하세요.")

                                            except Exception as e: 
                                                st.error(f"❌ **치명적 서비스 오류:** Dr. {user_name} (API 서비스 구축 실패): 인증 파일이나 권한을 확인하세요. (오류: {e})")
                                                
                                        else: st.warning(f"⚠️ **Dr. {res['name']}**님은 Google Calendar 계정이 연동되지 않았습니다.")
                        else: st.info("매칭된 치과의사 계정이 없습니다.")
    
    # -----------------------------------------------------
    # 탭 2: 사용자 목록 및 관리 로직 복원 (인증 필요) 🚨
    # -----------------------------------------------------
    with tab_user_mgmt:
        # 🚨 Admin 비밀번호 확인 로직
        if not st.session_state.admin_password_correct:
            st.subheader("🔑 사용자 관리 권한 인증")
            admin_password_input = st.text_input("관리자 비밀번호를 입력하세요.", type="password", key="admin_password_check_tab2")
            
            try:
                # secrets.toml에서 직접 해시된 비밀번호를 가져옴
                admin_pw_hash = st.secrets["admin"]["password"] 
            except KeyError:
                # secrets에 설정이 없으면 기본 비밀번호 사용
                admin_pw_hash = DEFAULT_PASSWORD
                st.warning("⚠️ secrets.toml 파일에 'admin.password' 설정이 없습니다. 기본 비밀번호를 사용합니다.")
            
            if st.button("사용자 관리 인증", key="admin_auth_button_tab2"):
                # 비밀번호 확인 로직 (bcrypt 해시와 평문 모두 고려)
                if check_password(admin_password_input, admin_pw_hash) or \
                   (admin_password_input == admin_pw_hash and not admin_pw_hash.startswith('$2b')):
                    st.session_state.admin_password_correct = True
                    # 평문 비교에 성공했으나 해시가 아닌 경우, 해시화하여 업데이트를 시도해야 함 (생략)
                    st.success("✅ 사용자 관리 인증 성공! 기능을 로드합니다.")
                    st.rerun()
                else:
                    st.error("비밀번호가 일치하지 않습니다. 관리자 계정을 확인하세요.")
            
            # 인증 전에는 아래 기능들을 표시하지 않고 여기서 함수 종료
            return 
        
        # --- 인증 성공 후 사용자 관리 기능 실행 ---
        st.subheader("👥 사용자 목록 및 계정 관리")
        
        tab_student, tab_doctor, tab_test_mail = st.tabs(["📚 학생 사용자 관리", "🧑‍⚕️ 치과의사 사용자 관리", "📧 테스트 메일 발송"])

        # DB 사용자 데이터 로드
        user_meta = users_ref.get()
        user_list = [{"name": u.get('name'), "email": u.get('email'), "key": k} for k, u in user_meta.items() if u and isinstance(u, dict)] if user_meta else []
        doctor_meta = doctor_users_ref.get()
        doctor_list = [{"name": d.get('name'), "email": d.get('email'), "key": k, "dept": d.get('department')} for k, d in doctor_meta.items() if d and isinstance(d, dict)] if doctor_meta else []

        # --- 탭 2-1: 학생 사용자 관리 ---
        with tab_student:
            st.markdown("#### 학생 사용자 목록")
            if user_list:
                df_users = pd.DataFrame(user_list)
                st.dataframe(df_users[['name', 'email']], use_container_width=True)

                st.markdown("---")
                
                # 1-1. 학생 사용자 선택 (Multiselect)
                user_options = [f"{u['name']} ({u['email']})" for u in user_list]
                selected_users_to_act = st.multiselect(
                    "메일 발송 또는 삭제할 학생을 선택하세요:", 
                    options=user_options, 
                    key="student_multiselect_act"
                )
                
                selected_user_data = [u for u in user_list if f"{u['name']} ({u['email']})" in selected_users_to_act]
                
                if selected_user_data:
                    
                    # 1-2. 메일 발송 기능
                    with st.expander("📧 선택된 학생에게 메일 발송"):
                        mail_subject = st.text_input("메일 제목 (선택사항)", key="student_mail_subject")
                        mail_body = st.text_area("메일 내용", key="student_mail_body")
                        
                        if st.button(f"선택된 {len(selected_user_data)}명에게 메일 발송 실행", key="send_bulk_student_mail_btn"):
                            success_count = 0
                            for user_info in selected_user_data:
                                try:
                                    send_email(
                                        receiver=user_info['email'], 
                                        rows=[], 
                                        sender=sender, 
                                        password=sender_pw, 
                                        custom_message=f"<h4>{mail_subject}</h4><p>{mail_body}</p>",
                                        date_str="Admin 발송 테스트"
                                    )
                                    success_count += 1
                                except Exception as e:
                                    st.error(f"❌ {user_info['email']} 메일 발송 실패: {e}")
                            st.success(f"✅ 총 {success_count}명에게 메일 발송 완료!")

                    # 1-3. 일괄 삭제 기능
                    if st.session_state.get('student_delete_confirm', False) is False:
                        if st.button(f"선택된 {len(selected_user_data)}명 일괄 삭제 준비", key="init_student_delete_btn"):
                            st.session_state.student_delete_confirm = True
                            st.rerun()

                    if st.session_state.get('student_delete_confirm', False):
                        st.warning(f"⚠️ **{len(selected_user_data)}명**의 학생 계정을 영구적으로 삭제하시겠습니까?")
                        col_yes, col_no = st.columns(2)
                        if col_yes.button("예, 학생 계정 일괄 삭제", key="confirm_bulk_student_delete_btn"):
                            deleted_count = 0
                            for user_info in selected_user_data:
                                users_ref.child(user_info['key']).delete()
                                deleted_count += 1
                            st.session_state.student_delete_confirm = False
                            st.success(f"🎉 {deleted_count}명의 학생 계정이 삭제되었습니다.")
                            st.rerun()
                        if col_no.button("아니오, 취소", key="cancel_bulk_student_delete_btn"):
                            st.session_state.student_delete_confirm = False
                            st.rerun()
                            
            else:
                st.info("등록된 학생 사용자가 없습니다.")

        # --- 탭 2-2: 치과의사 사용자 관리 ---
        with tab_doctor:
            st.markdown("#### 치과의사 사용자 목록")
            if doctor_list:
                df_doctors = pd.DataFrame(doctor_list)
                st.dataframe(df_doctors[['name', 'email', 'dept']], use_container_width=True)

                st.markdown("---")
                
                # 2-1. 치과의사 사용자 선택 (Multiselect)
                doctor_options = [f"{d['name']} ({d['email']})" for d in doctor_list]
                selected_doctors_to_act = st.multiselect(
                    "메일 발송 또는 삭제할 치과의사를 선택하세요:", 
                    options=doctor_options, 
                    key="doctor_multiselect_act"
                )
                
                selected_doctor_data = [d for d in doctor_list if f"{d['name']} ({d['email']})" in selected_doctors_to_act]
                
                if selected_doctor_data:
                    
                    # 2-2. 메일 발송 기능
                    with st.expander("📧 선택된 치과의사에게 메일 발송"):
                        mail_subject = st.text_input("메일 제목 (선택사항)", key="doctor_mail_subject")
                        mail_body = st.text_area("메일 내용", key="doctor_mail_body")
                        
                        if st.button(f"선택된 {len(selected_doctor_data)}명에게 메일 발송 실행", key="send_bulk_doctor_mail_btn"):
                            success_count = 0
                            for doctor_info in selected_doctor_data:
                                try:
                                    send_email(
                                        receiver=doctor_info['email'], 
                                        rows=[], 
                                        sender=sender, 
                                        password=sender_pw, 
                                        custom_message=f"<h4>{mail_subject}</h4><p>{mail_body}</p>",
                                        date_str="Admin 발송 테스트"
                                    )
                                    success_count += 1
                                except Exception as e:
                                    st.error(f"❌ {doctor_info['email']} 메일 발송 실패: {e}")
                            st.success(f"✅ 총 {success_count}명에게 메일 발송 완료!")

                    # 2-3. 일괄 삭제 기능
                    if st.session_state.get('doctor_delete_confirm', False) is False:
                        if st.button(f"선택된 {len(selected_doctor_data)}명 일괄 삭제 준비", key="init_doctor_delete_btn"):
                            st.session_state.doctor_delete_confirm = True
                            st.rerun()

                    if st.session_state.get('doctor_delete_confirm', False):
                        st.warning(f"⚠️ **{len(selected_doctor_data)}명**의 치과의사 계정을 영구적으로 삭제하시겠습니까?")
                        col_yes, col_no = st.columns(2)
                        if col_yes.button("예, 치과의사 계정 일괄 삭제", key="confirm_bulk_doctor_delete_btn"):
                            deleted_count = 0
                            for doctor_info in selected_doctor_data:
                                doctor_users_ref.child(doctor_info['key']).delete()
                                deleted_count += 1
                            st.session_state.doctor_delete_confirm = False
                            st.success(f"🎉 {deleted_count}명의 치과의사 계정이 삭제되었습니다.")
                            st.rerun()
                        if col_no.button("아니오, 취소", key="cancel_bulk_doctor_delete_btn"):
                            st.session_state.doctor_delete_confirm = False
                            st.rerun()
                            
            else:
                st.info("등록된 치과의사 사용자가 없습니다.")
        
        # --- 탭 2-3: 테스트 메일 발송 ---
        with tab_test_mail:
            st.subheader("📧 테스트 메일 발송")
            test_email_recipient = st.text_input("테스트 메일 수신자 이메일 주소", key="test_email_recipient")
            
            if st.button("테스트 메일 발송", key="send_test_mail_btn"):
                if is_valid_email(test_email_recipient):
                    try:
                        send_email(
                            receiver=test_email_recipient, 
                            rows=[], 
                            sender=sender, 
                            password=sender_pw, 
                            custom_message="""<p>이 메일은 환자 내원 확인 시스템에서 발송한 테스트 메일입니다. 시스템 정상 작동을 확인해 주세요.</p>""",
                            date_str=datetime.datetime.now().strftime("%Y-%m-%d")
                        )
                        st.success(f"테스트 메일이 {test_email_recipient}에게 성공적으로 발송되었습니다.")
                    except Exception as e:
                        st.error(f"테스트 메일 발송 실패: {e}. secrets.toml의 [gmail] 정보를 확인해주세요.")
                else:
                    st.error("유효한 이메일 주소를 입력해주세요.")

# --- 4. 일반 사용자 모드 UI ---

def show_user_mode_ui(firebase_key, user_name):
    """일반 사용자 모드 (환자 등록 및 관리, 분석 결과) UI를 표시합니다."""
    patients_ref_for_user = db_ref_func(f"patients/{firebase_key}")

    registration_tab, analysis_tab = st.tabs(['✅ 환자 등록 및 관리', '📈 OCS 분석 결과'])

    # --- 환자 등록 및 관리 탭 ---
    with registration_tab:
        st.subheader("Google Calendar 연동")
        get_google_calendar_service(firebase_key) # 서비스 로드 시도
        if st.session_state.get('google_calendar_service'): st.success("✅ 캘린더 추가 기능이 허용되어 있습니다.")
        else: st.info("구글 캘린더 연동을 위해 인증이 필요합니다.")
        st.markdown("---")
        
        st.subheader(f"{user_name}님의 토탈 환자 목록")
        existing_patient_data = patients_ref_for_user.get()
        
        # 🚨 [수정] existing_patient_data가 None일 경우 빈 딕셔너리로 초기화 (오류 해결 핵심)
        if existing_patient_data is None:
            existing_patient_data = {}
        
        # 환자 목록 표시 로직
        if existing_patient_data:
            # ... (환자 정렬 및 표시 로직) ...
            patient_list = list(existing_patient_data.items())
            valid_patient_list = [item for item in patient_list if isinstance(item[1], dict)]
            sorted_patient_list = sorted(valid_patient_list, key=lambda item: (
                0 if item[1].get('소치', False) else 1 if item[1].get('외과', False) else 2 if item[1].get('내과', False) else 3 if item[1].get('교정', False) else 4 if item[1].get('보철', False) else 5 if item[1].get('원진실', False) else 6 if item[1].get('보존', False) else 7, 
                item[1].get('환자이름', 'zzz')
            ))
            cols_count = 3; cols = st.columns(cols_count)
            for idx, (pid_key, val) in enumerate(sorted_patient_list): 
                with cols[idx % cols_count]:
                    with st.container(border=True):
                         registered_depts = [dept.capitalize() for dept in PATIENT_DEPT_FLAGS if val.get(dept.lower()) is True or val.get(dept.lower()) == 'True']
                         depts_str = ", ".join(registered_depts) if registered_depts else "미지정"
                         info_col, btn_col = st.columns([4, 1])
                         with info_col: st.markdown(f"**{val.get('환자이름', '이름 없음')}** / {pid_key} / {depts_str}")
                         with btn_col:
                             # 개별 삭제 버튼
                             if st.button("X", key=f"delete_button_{pid_key}"):
                                 patients_ref_for_user.child(pid_key).delete(); st.rerun()

        else: st.info("등록된 환자가 없습니다.")
        st.markdown("---")

        ## 📋 환자 정보 대량 등록 섹션 (복원)
        st.subheader("📋 환자 정보 대량 등록")
        
        paste_area = st.text_area(
            "엑셀 또는 다른 곳에서 복사한 데이터를 여기에 붙여넣으세요 (환자명, 진료번호, 진료과를 탭/공백으로 구분).", 
            height=150, 
            key="bulk_paste_area",
            placeholder="예시: 홍길동\t12345678\t교정,보철\n김철수\t87654321\t소치\n(진료과는 쉼표로 구분 가능)"
        )
        bulk_submit = st.button("대량 등록 실행", key="bulk_reg_button")
        
        if bulk_submit and paste_area:
            lines = paste_area.strip().split('\n')
            success_count = 0
            
            for line in lines:
                parts = re.split(r'[\t\s]+', line.strip(), 2) # 탭, 공백 등으로 3부분 분리
                if len(parts) >= 3:
                    name, pid, depts_str = parts[0], parts[1], parts[2]
                    pid_key = pid.strip()
                    
                    # 진료과 목록 파싱 (쉼표로 구분된 경우)
                    selected_departments = [d.strip() for d in depts_str.replace(",", " ").split()]
                    
                    if name and pid_key and selected_departments:
                        # existing_patient_data가 딕셔너리이므로 안전하게 .get() 호출 가능
                        current_data = existing_patient_data.get(pid_key, {"환자이름": name, "진료번호": pid_key}) 
                        
                        # 진료과 플래그 업데이트
                        for dept_flag in PATIENT_DEPT_FLAGS + ['치주', '원진실']: current_data[dept_flag.lower()] = False
                        for dept in selected_departments: current_data[dept.lower()] = True
                        
                        patients_ref_for_user.child(pid_key).set(current_data)
                        success_count += 1
                    else:
                        st.warning(f"데이터 형식 오류로 건너뜀: {line}")
            
            if success_count > 0:
                st.success(f"🎉 총 {success_count}명의 환자 정보가 등록/업데이트되었습니다.")
                st.rerun()
            else:
                st.error("등록할 유효한 환자 정보가 없습니다. 형식을 확인해주세요.")

        st.markdown("---")

        ## 🗑️ 환자 정보 일괄 삭제 섹션 (복원)
        st.subheader("🗑️ 환자 정보 일괄 삭제")
        
        if existing_patient_data:
            patient_options = {
                f"{val.get('환자이름', '이름 없음')} ({pid_key})": pid_key
                for pid_key, val in existing_patient_data.items() 
                if isinstance(val, dict) # 유효한 데이터만 필터링
            }
            
            # 사용자에게 삭제할 환자 선택 요청
            selected_patients_str = st.multiselect(
                "삭제할 환자를 선택하세요:", 
                options=list(patient_options.keys()), 
                default=[], 
                key="delete_patient_multiselect"
            )
            
            # 실제 삭제할 환자 PID 목록 추출
            patients_to_delete = [patient_options[name_str] for name_str in selected_patients_str]

            if patients_to_delete:
                st.session_state.patients_to_delete = patients_to_delete
                st.session_state.delete_patient_confirm = True
            else:
                st.session_state.delete_patient_confirm = False
                
            
            
            # 삭제 확인 버튼 및 로직
            if st.session_state.delete_patient_confirm:
                st.warning(f"⚠️ **{len(st.session_state.patients_to_delete)}명**의 환자 정보를 영구적으로 삭제하시겠습니까?")
                
                if st.button("예, 선택된 환자 일괄 삭제", key="confirm_delete_button"):
                    deleted_count = 0
                    for pid_key in st.session_state.patients_to_delete:
                        patients_ref_for_user.child(pid_key).delete()
                        deleted_count += 1
                        
                    st.session_state.delete_patient_confirm = False
                    st.session_state.patients_to_delete = []
                    st.success(f"🎉 **{deleted_count}명**의 환자 정보가 성공적으로 삭제되었습니다.")
                    st.rerun()
            
        else:
            st.info("현재 등록된 환자가 없어 삭제할 항목이 없습니다.")

        st.markdown("---")

        # 단일 환자 등록 폼
        with st.form("register_form"):
            name = st.text_input("환자명")
            pid = st.text_input("진료번호")
            selected_departments = st.multiselect("등록할 진료과 (복수 선택 가능)", DEPARTMENTS_FOR_REGISTRATION)
            submitted = st.form_submit_button("등록")
            
            if submitted:
                if not name or not pid or not selected_departments: st.warning("환자명, 진료번호, 등록할 진료과를 모두 입력/선택해주세요.")
                else:
                    pid_key = pid.strip()
                    # existing_patient_data가 딕셔너리이므로 안전하게 .get() 호출 가능
                    new_patient_data = existing_patient_data.get(pid_key, {"환자이름": name, "진료번호": pid}) 
                    for dept_flag in PATIENT_DEPT_FLAGS + ['치주', '원진실']: new_patient_data[dept_flag.lower()] = False
                    for dept in selected_departments: new_patient_data[dept.lower()] = True
                        
                    patients_ref_for_user.child(pid_key).set(new_patient_data)
                    st.success(f"{name} ({pid}) [{', '.join(selected_departments)}] 환자 등록/업데이트 완료")
                    st.rerun()

    # --- OCS 분석 결과 탭 ---
    with analysis_tab:
        st.header("📈 OCS 분석 결과")
        analysis_results = db_ref_func("ocs_analysis/latest_result").get()
        latest_file_name = db_ref_func("ocs_analysis/latest_file_name").get()

        if analysis_results and latest_file_name:
            st.markdown(f"**<h3 style='text-align: left;'>{latest_file_name} 분석 결과</h3>**", unsafe_allow_html=True)
            st.markdown("---")
            
            # 분석 결과 표시 로직 (소치, 보존, 교정)
            for dept in ['소치', '보존', '교정']:
                if dept in analysis_results:
                    st.subheader(f"{dept} 현황 (오전/오후)")
                    st.info(f"오전: **{analysis_results[dept]['오전']}명**")
                    st.info(f"오후: **{analysis_results[dept]['오후']}명**")
                    st.markdown("---")
                else: st.warning(f"{dept} 데이터가 엑셀 파일에 없습니다.")
        else: st.info("💡 분석 결과가 없습니다. 관리자가 엑셀 파일을 업로드하면 표시됩니다.")
            
        st.divider(); st.header("🔑 비밀번호 변경")
        new_password = st.text_input("새 비밀번호를 입력하세요", type="password", key="user_new_password_input")
        confirm_password = st.text_input("새 비밀번호를 다시 입력하세요", type="password", key="user_confirm_password_input")
        
        if st.button("비밀번호 변경", key="user_password_change_btn"):
            if new_password and new_password == confirm_password:
                # 🔑 새 비밀번호를 해시하여 저장
                hashed_pw = hash_password(new_password)
                users_ref.child(firebase_key).update({"password": hashed_pw})
                st.success("🎉 비밀번호가 성공적으로 변경되었습니다!")
            else: st.error("새 비밀번호가 일치하지 않거나 입력되지 않았습니다.")

# --- 5. 치과의사 모드 UI ---

def show_doctor_mode_ui(firebase_key, user_name):
    """치과의사 모드 UI를 표시합니다."""
    st.header(f"🧑‍⚕️Dr. {user_name}")
    st.subheader("🗓️ Google Calendar 연동")
    get_google_calendar_service(firebase_key) 
    if st.session_state.get('google_calendar_service'): st.success("✅ 캘린더 추가 기능이 허용되어 있습니다.")
    else: st.info("구글 캘린더 연동을 위해 인증이 필요합니다.")
    
    st.markdown("---")
    st.header("🔑 비밀번호 변경")
    new_password = st.text_input("새 비밀번호를 입력하세요", type="password", key="res_new_password_input")
    confirm_password = st.text_input("새 비밀번호를 다시 입력하세요", type="password", key="res_confirm_password_input")

    if st.button("비밀번호 변경", key="res_password_change_btn"):
        if new_password and new_password == confirm_password:
            # 🔑 새 비밀번호를 해시하여 저장
            hashed_pw = hash_password(new_password)
            doctor_users_ref.child(firebase_key).update({"password": hashed_pw})
            st.success("🎉 비밀번호가 성공적으로 변경되었습니다!")
        else: st.error("새 비밀번호가 일치하지 않거나 입력되지 않았습니다.")
