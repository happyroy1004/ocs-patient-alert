# ui_manager.py (수정 전체 코드)

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
    get_google_calendar_service, save_google_creds_to_firebase, load_google_creds_from_firebase
)
import excel_utils
from notification_utils import (
    is_valid_email, send_email, create_calendar_event, 
    get_matching_data, run_auto_notifications
)
from professor_reviews_module import show_professor_review_system 

# DB 레퍼런스 초기 로드
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
    if 'reservation_date_excel' not in st.session_state: st.session_state.reservation_date_excel = "날짜_미정"
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
    if users_ref is None:
        st.error("🚨 데이터베이스 연결에 문제가 있습니다. 관리자에게 문의하세요.")
        return
        
    if not user_name: st.error("사용자 이름을 입력해주세요.")
    elif user_name.strip().lower() == "admin": 
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
            login_success = check_password(password_input, user_password_db)
            is_plaintext_or_default = False
            
            if not login_success:
                if password_input == user_password_db:
                    login_success = True; is_plaintext_or_default = True
                elif (not user_password_db or user_password_db == DEFAULT_PASSWORD) and password_input == DEFAULT_PASSWORD:
                    login_success = True; is_plaintext_or_default = True
            
            if login_success:
                st.session_state.update({
                    'found_user_email': matched_user["email"], 
                    'current_firebase_key': safe_key_found, 
                    'current_user_name': user_name, 
                    'login_mode': 'user_mode'
                })
                if is_plaintext_or_default:
                    hashed_pw = hash_password(password_input if password_input else DEFAULT_PASSWORD)
                    users_ref.child(safe_key_found).update({"password": hashed_pw})
                    st.warning("⚠️ 보안 강화를 위해 비밀번호가 자동으로 암호화되었습니다.")

                st.info(f"**{user_name}**님으로 로그인되었습니다.")
                st.rerun()
            else: st.error("비밀번호가 일치하지 않습니다.")
        else:
            st.session_state.current_user_name = user_name
            st.session_state.login_mode = 'new_user_registration'
            st.rerun()

def _handle_doctor_login(doctor_email, password_input_doc):
    """치과의사 로그인 로직을 처리합니다."""
    if doctor_users_ref is None:
        st.error("🚨 데이터베이스 연결에 문제가 있습니다. 관리자에게 문의하세요.")
        return

    if not doctor_email: st.warning("치과의사 이메일 주소를 입력해주세요.")
    else:
        safe_key = sanitize_path(doctor_email)
        matched_doctor = doctor_users_ref.child(safe_key).get()
        
        if matched_doctor:
            doctor_password_db = matched_doctor.get("password")
            login_success = check_password(password_input_doc, doctor_password_db)
            is_plaintext_or_default = False
            
            if not login_success:
                if password_input_doc == doctor_password_db:
                    login_success = True; is_plaintext_or_default = True
                elif (not doctor_password_db or doctor_password_db == DEFAULT_PASSWORD) and password_input_doc == DEFAULT_PASSWORD:
                    login_success = True; is_plaintext_or_default = True

            if login_success:
                st.session_state.update({
                    'found_user_email': matched_doctor["email"], 
                    'current_firebase_key': safe_key, 
                    'current_user_name': matched_doctor.get("name"),
                    'current_user_dept': matched_doctor.get("department"),
                    'current_user_role': 'doctor',
                    'login_mode': 'doctor_mode'
                })
                if is_plaintext_or_default:
                    hashed_pw = hash_password(password_input_doc if password_input_doc else DEFAULT_PASSWORD)
                    doctor_users_ref.child(safe_key).update({"password": hashed_pw})
                    st.warning("⚠️ 보안 강화를 위해 비밀번호가 자동으로 암호화되었습니다.")

                st.info(f"치과의사 **{st.session_state.current_user_name}**님으로 로그인되었습니다.")
                st.rerun()
            else: st.error("비밀번호가 일치하지 않습니다.")
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
            if st.button("로그인/등록", key="login_button_tab1"): _handle_user_login(user_name, password_input)
        with tab2:
            st.subheader("🧑‍⚕️ 치과의사 로그인")
            doctor_email = st.text_input("치과의사 이메일 주소를 입력하세요", key="doctor_email_input_tab2")
            password_input_doc = st.text_input("비밀번호를 입력하세요", type="password", key="doctor_password_input_tab2")
            if st.button("로그인/등록", key="doctor_login_button_tab2"): _handle_doctor_login(doctor_email, password_input_doc)

    elif st.session_state.get('login_mode') == 'new_user_registration':
        st.info(f"'{st.session_state.current_user_name}'님은 새로운 사용자입니다. 아래에 정보를 입력하여 등록을 완료하세요.")
        st.subheader("👨‍⚕️ 신규 사용자 등록")
        new_email_input = st.text_input("아이디(이메일)를 입력하세요", key="new_user_email_input")
        password_input = st.text_input("새로운 비밀번호를 입력하세요", type="password", key="new_user_password_input")
        
        if st.button("사용자 등록 완료", key="new_user_reg_button"):
            if is_valid_email(new_email_input) and password_input:
                new_firebase_key = sanitize_path(new_email_input)
                if users_ref is None: st.error("🚨 데이터베이스 연결 오류")
                elif users_ref.child(new_firebase_key).get(): st.error("이미 등록된 이메일입니다.")
                else:
                    hashed_pw = hash_password(password_input)
                    users_ref.child(new_firebase_key).set({"name": st.session_state.current_user_name, "email": new_email_input, "password": hashed_pw})
                    st.session_state.update({'current_firebase_key': new_firebase_key, 'found_user_email': new_email_input, 'login_mode': 'user_mode'})
                    st.success("등록 완료"); st.rerun()
            else: st.error("올바른 이메일과 비밀번호를 입력하세요.")

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
                if doctor_users_ref is None: st.error("🚨 데이터베이스 연결 오류")
                else:
                    hashed_pw = hash_password(password_input)
                    doctor_users_ref.child(new_firebase_key).set({"name": new_doctor_name_input, "email": user_id_input, "password": hashed_pw, "role": 'doctor', "department": department})
                    st.session_state.update({'current_firebase_key': new_firebase_key, 'found_user_email': user_id_input, 'current_user_name': new_doctor_name_input, 'current_user_dept': department, 'login_mode': 'doctor_mode'})
                    st.success("등록 완료"); st.rerun()
            else: st.error("모든 정보를 올바르게 입력해주세요.")

# --- 3. 관리자 모드 UI (Excel 및 알림) ---

def show_admin_mode_ui():
    """관리자 모드 (엑셀 업로드, 알림 전송) UI를 표시합니다."""
    
    st.markdown("---")
    st.title("💻 관리자 모드")
    
    db_ref = db_ref_func
    try: sender = st.secrets["gmail"]["sender"]; sender_pw = st.secrets["gmail"]["app_password"]
    except KeyError: st.error("⚠️ [gmail] 정보 누락"); sender = "error@example.com"; sender_pw = "none"

    tab_excel, tab_user_mgmt = st.tabs(["📊 OCS 파일 처리 및 알림", "🧑‍💻 사용자 목록 및 관리"])
    
    # -----------------------------------------------------
    # 탭 1: OCS 파일 처리 및 알림
    # -----------------------------------------------------
    with tab_excel:
        st.subheader("💻 Excel File Processor")
        uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])
        
        if uploaded_file:
            file_name = uploaded_file.name; 
            is_daily = excel_utils.is_daily_schedule(file_name) 
            
            password = None
            if excel_utils.is_encrypted_excel(uploaded_file): 
                password = st.text_input("⚠️ 암호화된 파일입니다. 비밀번호를 입력해주세요.", type="password", key="admin_password_file")
                if not password: st.info("비밀번호 입력 대기 중..."); st.stop()

            try:
                xl_object, raw_file_io = excel_utils.load_excel(uploaded_file, password)
                excel_data_dfs_raw, styled_excel_bytes = excel_utils.process_excel_file_and_style(raw_file_io, db_ref_func)
                analysis_results = excel_utils.run_analysis(excel_data_dfs_raw)
                
                if analysis_results and any(analysis_results.values()): 
                    today_date_str = datetime.datetime.now().strftime("%Y-%m-%d")
                    db_ref("ocs_analysis/latest_result").set(analysis_results)
                    db_ref("ocs_analysis/latest_date").set(today_date_str)
                    db_ref("ocs_analysis/latest_file_name").set(file_name)
                else: st.warning("⚠️ 분석 결과가 비어 있어 Firebase에 저장하지 않았습니다.")
                
                st.session_state.last_processed_data = excel_data_dfs_raw; st.session_state.last_processed_file_name = file_name

                if styled_excel_bytes:
                    output_filename = uploaded_file.name.replace(".xlsx", "_processed.xlsx").replace(".xlsm", "_processed.xlsm")
                    st.download_button("처리된 엑셀 다운로드", data=styled_excel_bytes, file_name=output_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    st.success("✅ 파일 처리 완료. 알림 전송 방법을 선택하세요.")
                else: st.warning("처리할 데이터가 없습니다.")
                    
            except ValueError as ve: st.error(f"파일 처리 실패: {ve}"); st.stop()
            except Exception as e: st.error(f"오류 발생: {e}"); st.stop()
            
            st.markdown("---")
            st.subheader("🚀 알림 전송 옵션")
            col_auto, col_manual = st.columns(2)

            with col_auto:
                if st.button("YES: 자동으로 모든 사용자에게 전송", key="auto_run_yes"):
                    st.session_state.auto_run_confirmed = True; st.rerun()
            with col_manual:
                if st.button("NO: 수동으로 사용자 선택", key="auto_run_no"):
                    st.session_state.auto_run_confirmed = False; st.rerun()
                    
            if 'last_processed_data' in st.session_state and st.session_state.last_processed_data:
                
                all_users_meta = users_ref.get(); all_patients_data = db_ref("patients").get()
                all_doctors_meta = doctor_users_ref.get()
                excel_data_dfs = st.session_state.last_processed_data
                
                # 매칭 실행 (get_matching_data는 number 정보도 가져옴)
                matched_users, matched_doctors_data = get_matching_data(
                    excel_data_dfs, all_users_meta, all_patients_data, all_doctors_meta
                )

                # A. 자동 실행 (notification_utils의 함수 이용 - 이미 수정됨)
                if st.session_state.auto_run_confirmed:
                    st.markdown("---")
                    st.warning("자동으로 모든 매칭 사용자에게 알림(메일/캘린더)을 전송합니다.")
                    run_auto_notifications(matched_users, matched_doctors_data, excel_data_dfs, file_name, is_daily, db_ref_func)
                    st.session_state.auto_run_confirmed = None; st.stop()
                    
                # B. 수동 실행 (여기서 직접 구현 - 수정 필요)
                elif st.session_state.auto_run_confirmed is False:
                    st.markdown("---")
                    st.info("수동으로 사용자를 선택하여 전송합니다.")

                    student_admin_tab, doctor_admin_tab = st.tabs(['📚 학생 수동 전송', '🧑‍⚕️ 치과의사 수동 전송'])
                    
                    # --- [수정] 학생 수동 전송 ---
                    with student_admin_tab:
                        st.subheader("📚 학생 수동 전송 (매칭 결과)");
                        if matched_users:
                            st.success(f"매칭된 환자가 있는 **{len(matched_users)}명의 사용자**를 발견했습니다.")
                            matched_user_list_for_dropdown = [f"{user['name']} ({user['email']})" for user in matched_users]
                            
                            if st.button("매칭된 사용자 모두 선택/해제", key="select_all_matched_btn"):
                                if len(st.session_state.matched_user_multiselect) == len(matched_user_list_for_dropdown):
                                    st.session_state.matched_user_multiselect = []
                                else: st.session_state.matched_user_multiselect = matched_user_list_for_dropdown
                                st.rerun()
                            
                            selected_users_to_act_values = st.multiselect(
                                "액션을 취할 사용자 선택", matched_user_list_for_dropdown, 
                                default=st.session_state.matched_user_multiselect, key="matched_user_multiselect" 
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
                                        # [추가] 번호 정보 가져오기
                                        user_number = user_match_info.get('number', '')

                                        email_cols = ['환자명', '진료번호', '예약의사', '진료내역', '예약일시', '예약시간', '등록과']
                                        df_for_mail = df_matched[[col for col in email_cols if col in df_matched.columns]]
                                        rows_as_dict = df_for_mail.to_dict('records')
                                        df_html = df_for_mail.to_html(index=False, escape=False)
                                        
                                        # [수정] 텍스트 데이터 생성 로직 직접 구현
                                        text_lines = []
                                        u_num = str(user_number).strip(); u_name = str(user_name).strip()
                                        for _, row in df_matched.iterrows():
                                            try:
                                                raw_date = str(row.get('예약일시', '')).strip().replace('-', '/').replace('.', '/')
                                                raw_time = str(row.get('예약시간', '')).strip()
                                                
                                                date_digits = re.sub(r'[^0-9]', '', raw_date)
                                                mmdd = date_digits[-4:] if len(date_digits) >= 4 else "0000"
                                                time_digits = re.sub(r'[^0-9]', '', raw_time)
                                                hhmm = time_digits.zfill(4) if len(time_digits) <= 4 else time_digits[:4]
                                                
                                                line = f"{row.get('예약의사','')},{mmdd},{hhmm},{row.get('환자명','')},{row.get('진료번호','')},{u_num},{u_name}"
                                                text_lines.append(line)
                                            except: continue
                                        formatted_text_html = "<br>".join(text_lines)

                                        # [수정] 본문 구성
                                        email_body = f"""<p>안녕하세요, {user_name}님.</p><p>{file_name} 분석 결과, 내원 예정인 환자 진료 정보입니다.</p>{df_html}
                                        <br><br><div style='font-family: sans-serif; font-size: 14px; line-height: 1.6; color: #333;'>{formatted_text_html}</div><p>확인 부탁드립니다.</p>"""
                                        
                                        try: send_email(real_email, rows_as_dict, sender, sender_pw, custom_message=email_body, date_str=file_name); st.success(f"**{user_name}**님에게 메일 전송 완료!")
                                        except Exception as e: st.error(f"**{user_name}**님에게 메일 전송 실패: {e}")

                            with calendar_col:
                                if st.button("선택된 사용자에게 Google Calendar 일정 추가", key="manual_send_calendar_student"):
                                    for user_match_info in selected_matched_users_data:
                                        user_safe_key = user_match_info['safe_key']; user_name = user_match_info['name']; df_matched = user_match_info['data']
                                        # [추가] 번호 정보
                                        user_number = user_match_info.get('number', '')
                                        
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
                                                            
                                                            # [수정] user_name, user_number 전달
                                                            success = create_calendar_event(
                                                                service, row.get('환자명', 'N/A'), row.get('진료번호', ''), row.get('등록과', ''), 
                                                                reservation_datetime, row.get('예약의사', 'N/A'), row.get('진료내역', ''), is_daily,
                                                                user_name=user_name, user_number=user_number
                                                            )
                                                            if success: successful_adds += 1
                                                        except: pass
                                                if successful_adds > 0: st.success(f"**{user_name}**님 캘린더에 {successful_adds}건 추가 완료.")
                                                else: st.warning(f"**{user_name}**님 캘린더에 추가된 일정 없음.")
                                            except Exception as e: st.error(f"❌ {user_name} 캘린더 오류: {e}")
                                        else: st.warning(f"**{user_name}**님 캘린더 미연동.")
                        else: st.info("매칭된 환자가 없습니다.")

                    # --- [수정] 치과의사 수동 전송 ---
                    with doctor_admin_tab:
                        st.subheader("🧑‍⚕️ 치과의사 수동 전송 (매칭 결과)");
                        if matched_doctors_data:
                            st.success(f"등록된 진료가 있는 **{len(matched_doctors_data)}명의 치과의사**를 발견했습니다.")
                            doctor_list_for_multiselect = [f"{res['name']} ({res['email']})" for res in matched_doctors_data]

                            if st.button("등록된 치과의사 모두 선택/해제", key="select_all_matched_res_btn"):
                                if len(st.session_state.matched_doctor_multiselect) == len(doctor_list_for_multiselect):
                                    st.session_state.matched_doctor_multiselect = []
                                else: st.session_state.matched_doctor_multiselect = doctor_list_for_multiselect
                                st.rerun()

                            selected_doctors_str = st.multiselect(
                                "액션을 취할 치과의사 선택", doctor_list_for_multiselect, 
                                default=st.session_state.matched_doctor_multiselect, key="matched_doctor_multiselect" 
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
                                        user_name = res['name']; user_number = res.get('number', '')
                                        
                                        email_cols = ['환자명', '진료번호', '예약의사', '진료내역', '예약일시', '예약시간']; 
                                        df_for_mail = df_matched[[col for col in email_cols if col in df_matched.columns]]
                                        df_html = df_for_mail.to_html(index=False, border=1); rows_as_dict = df_for_mail.to_dict('records')
                                        
                                        # [수정] 텍스트 데이터 생성 로직
                                        text_lines = []
                                        u_num = str(user_number).strip(); u_name = str(user_name).strip()
                                        for _, row in df_matched.iterrows():
                                            try:
                                                raw_date = str(row.get('예약일시', '')).strip().replace('-', '/').replace('.', '/')
                                                raw_time = str(row.get('예약시간', '')).strip()
                                                date_digits = re.sub(r'[^0-9]', '', raw_date)
                                                mmdd = date_digits[-4:] if len(date_digits) >= 4 else "0000"
                                                time_digits = re.sub(r'[^0-9]', '', raw_time)
                                                hhmm = time_digits.zfill(4) if len(time_digits) <= 4 else time_digits[:4]
                                                line = f"{row.get('예약의사','')},{mmdd},{hhmm},{row.get('환자명','')},{row.get('진료번호','')},{u_num},{u_name}"
                                                text_lines.append(line)
                                            except: continue
                                        formatted_text_html = "<br>".join(text_lines)

                                        email_body = f"""<p>안녕하세요, {res['name']} 치과의사님.</p><p>{latest_file_name}에서 가져온 내원할 환자 정보입니다.</p>{df_html}
                                        <br><br><div style='font-family: sans-serif; font-size: 14px; line-height: 1.6; color: #333;'>{formatted_text_html}</div><p>확인 부탁드립니다.</p>"""
                                        
                                        try: send_email(res['email'], rows_as_dict, sender, sender_pw, custom_message=email_body, date_str=latest_file_name); st.success(f"**Dr. {res['name']}**에게 메일 전송 완료!")
                                        except Exception as e: st.error(f"**Dr. {res['name']}**에게 메일 전송 실패: {e}")

                            with calendar_col_doc:
                                if st.button("선택된 치과의사에게 Google Calendar 일정 추가", key="manual_send_calendar_doctor"):
                                    for res in selected_doctors_to_act:
                                        user_safe_key = res['safe_key']; user_name = res['name']; df_matched = res['data']
                                        user_number = res.get('number', '')

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
                                                            # [수정] user_name, user_number 전달
                                                            success = create_calendar_event(
                                                                service, row.get('환자명', 'N/A'), row.get('진료번호', ''), res.get('department', 'N/A'), 
                                                                reservation_datetime, row.get('예약의사', ''), row.get('진료내역', ''), is_daily,
                                                                user_name=user_name, user_number=user_number
                                                            )
                                                            if success: successful_adds += 1
                                                        except: pass
                                                if successful_adds > 0: st.success(f"**Dr. {user_name}**님 캘린더에 {successful_adds}건 추가 완료.")
                                                else: st.warning(f"**Dr. {user_name}**님 캘린더에 추가된 일정 없음.")
                                            except Exception as e: st.error(f"❌ 오류: {e}")
                                        else: st.warning(f"⚠️ **Dr. {res['name']}**님은 캘린더 미연동.")
                        else: st.info("매칭된 치과의사 계정이 없습니다.")
    
    # -----------------------------------------------------
    # 탭 2: 사용자 목록 및 관리
    # -----------------------------------------------------
    with tab_user_mgmt:
        if not st.session_state.admin_password_correct:
            st.subheader("🔑 사용자 관리 권한 인증")
            admin_password_input = st.text_input("관리자 비밀번호를 입력하세요.", type="password", key="admin_password_check_tab2")
            try: admin_pw_hash = st.secrets["admin"]["password"] 
            except KeyError: admin_pw_hash = DEFAULT_PASSWORD; st.warning("⚠️ 기본 비밀번호 사용")
            if st.button("사용자 관리 인증", key="admin_auth_button_tab2"):
                if check_password(admin_password_input, admin_pw_hash) or (admin_password_input == admin_pw_hash and not admin_pw_hash.startswith('$2b')):
                    st.session_state.admin_password_correct = True; st.success("인증 성공"); st.rerun()
                else: st.error("비밀번호 불일치")
            return 
        
        st.subheader("👥 사용자 목록 및 계정 관리")
        tab_student, tab_doctor, tab_test_mail = st.tabs(["📚 학생 사용자 관리", "🧑‍⚕️ 치과의사 사용자 관리", "📧 테스트 메일 발송"])
        user_meta = users_ref.get(); user_list = [{"name": u.get('name'), "email": u.get('email'), "key": k} for k, u in user_meta.items() if u and isinstance(u, dict)] if user_meta else []
        doctor_meta = doctor_users_ref.get(); doctor_list = [{"name": d.get('name'), "email": d.get('email'), "key": k, "dept": d.get('department')} for k, d in doctor_meta.items() if d and isinstance(d, dict)] if doctor_meta else []

        with tab_student:
            st.markdown("#### 학생 사용자 목록")
            if user_list:
                df_users = pd.DataFrame(user_list); st.dataframe(df_users[['name', 'email']], use_container_width=True); st.markdown("---")
                user_options = [f"{u['name']} ({u['email']})" for u in user_list]
                selected_users_to_act = st.multiselect("메일 발송 또는 삭제할 학생:", options=user_options, key="student_multiselect_act")
                selected_user_data = [u for u in user_list if f"{u['name']} ({u['email']})" in selected_users_to_act]
                
                if selected_user_data:
                    with st.expander("📧 메일 발송"):
                        mail_subject = st.text_input("제목", key="student_mail_subject"); mail_body = st.text_area("내용", key="student_mail_body")
                        if st.button(f"전송 ({len(selected_user_data)}명)", key="send_bulk_student_mail_btn"):
                            success_count = 0
                            for user_info in selected_user_data:
                                try: send_email(user_info['email'], [], sender, sender_pw, custom_message=f"<h4>{mail_subject}</h4><p>{mail_body}</p>", date_str="Admin Test"); success_count += 1
                                except: pass
                            st.success(f"✅ {success_count}명 전송 완료")
                    if st.session_state.get('student_delete_confirm', False) is False:
                        if st.button("일괄 삭제 준비", key="init_student_delete_btn"): st.session_state.student_delete_confirm = True; st.rerun()
                    if st.session_state.get('student_delete_confirm', False):
                        st.warning(f"⚠️ **{len(selected_user_data)}명** 삭제?")
                        col_yes, col_no = st.columns(2)
                        if col_yes.button("예", key="confirm_bulk_student_delete_btn"):
                            for user_info in selected_user_data: users_ref.child(user_info['key']).delete()
                            st.session_state.student_delete_confirm = False; st.success("삭제 완료"); st.rerun()
                        if col_no.button("취소", key="cancel_bulk_student_delete_btn"): st.session_state.student_delete_confirm = False; st.rerun()
            else: st.info("등록된 학생 없음")

        with tab_doctor:
            st.markdown("#### 치과의사 사용자 목록")
            if doctor_list:
                df_doctors = pd.DataFrame(doctor_list); st.dataframe(df_doctors[['name', 'email', 'dept']], use_container_width=True); st.markdown("---")
                doctor_options = [f"{d['name']} ({d['email']})" for d in doctor_list]
                selected_doctors_to_act = st.multiselect("선택:", options=doctor_options, key="doctor_multiselect_act")
                selected_doctor_data = [d for d in doctor_list if f"{d['name']} ({d['email']})" in selected_doctors_to_act]
                
                if selected_doctor_data:
                    with st.expander("📧 메일 발송"):
                        mail_subject = st.text_input("제목", key="doctor_mail_subject"); mail_body = st.text_area("내용", key="doctor_mail_body")
                        if st.button(f"전송 ({len(selected_doctor_data)}명)", key="send_bulk_doctor_mail_btn"):
                            success_count = 0
                            for d in selected_doctor_data:
                                try: send_email(d['email'], [], sender, sender_pw, custom_message=f"<h4>{mail_subject}</h4><p>{mail_body}</p>", date_str="Admin Test"); success_count += 1
                                except: pass
                            st.success(f"✅ {success_count}명 전송 완료")
                    if st.session_state.get('doctor_delete_confirm', False) is False:
                        if st.button("일괄 삭제 준비", key="init_doctor_delete_btn"): st.session_state.doctor_delete_confirm = True; st.rerun()
                    if st.session_state.get('doctor_delete_confirm', False):
                        st.warning(f"⚠️ **{len(selected_doctor_data)}명** 삭제?")
                        col_yes, col_no = st.columns(2)
                        if col_yes.button("예", key="confirm_bulk_doctor_delete_btn"):
                            for d in selected_doctor_data: doctor_users_ref.child(d['key']).delete()
                            st.session_state.doctor_delete_confirm = False; st.success("삭제 완료"); st.rerun()
                        if col_no.button("취소", key="cancel_bulk_doctor_delete_btn"): st.session_state.doctor_delete_confirm = False; st.rerun()
            else: st.info("등록된 치과의사 없음")
        
        with tab_test_mail:
            st.subheader("📧 테스트 메일 발송")
            test_email_recipient = st.text_input("수신자 이메일", key="test_email_recipient")
            if st.button("발송", key="send_test_mail_btn"):
                if is_valid_email(test_email_recipient):
                    try: send_email(test_email_recipient, [], sender, sender_pw, custom_message="<p>테스트 메일입니다.</p>", date_str="Test"); st.success("성공")
                    except Exception as e: st.error(f"실패: {e}")
                else: st.error("이메일 형식 확인")

# --- 4. 일반 사용자 모드 UI ---

def show_user_mode_ui(firebase_key, user_name):
    """일반 사용자 모드 UI를 표시합니다."""
    patients_ref_for_user = db_ref_func(f"patients/{firebase_key}")
    registration_tab, analysis_tab, review_tab = st.tabs(['✅ 환자 등록 및 관리', '📈 OCS 분석 결과', '🧑‍🏫 케이스 방명록'])

    with registration_tab:
        st.subheader("Google Calendar 연동")
        get_google_calendar_service(firebase_key) 
        if st.session_state.get('google_calendar_service'): st.success("✅ 캘린더 추가 기능 허용됨")
        else: st.info("인증 필요")
        st.markdown("---")
        
        st.subheader(f"{user_name}님의 토탈 환자 목록")
        existing_patient_data = patients_ref_for_user.get()
        if existing_patient_data is None: existing_patient_data = {}
        
        if existing_patient_data:
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
                             if st.button("X", key=f"delete_button_{pid_key}"): patients_ref_for_user.child(pid_key).delete(); st.rerun()
        else: st.info("등록된 환자 없음")
        st.markdown("---")

        st.subheader("📋 환자 정보 대량 등록")
        paste_area = st.text_area("엑셀 붙여넣기 (이름 진료번호 진료과)", height=150, key="bulk_paste_area")
        if st.button("대량 등록 실행", key="bulk_reg_button") and paste_area:
            lines = paste_area.strip().split('\n'); success_count = 0
            for line in lines:
                parts = re.split(r'[\t\s]+', line.strip(), 2)
                if len(parts) >= 3:
                    name, pid, depts_str = parts[0], parts[1], parts[2]; pid_key = pid.strip()
                    selected_departments = [d.strip() for d in depts_str.replace(",", " ").split()]
                    current_data = existing_patient_data.get(pid_key, {"환자이름": name, "진료번호": pid_key}) 
                    for dept_flag in PATIENT_DEPT_FLAGS + ['치주', '원진실']: current_data[dept_flag.lower()] = False
                    for dept in selected_departments: current_data[dept.lower()] = True
                    patients_ref_for_user.child(pid_key).set(current_data); success_count += 1
            if success_count > 0: st.success(f"🎉 {success_count}명 등록 완료"); st.rerun()
            else: st.error("형식 오류")

        st.markdown("---")
        st.subheader("🗑️ 환자 정보 일괄 삭제")
        if existing_patient_data:
            patient_options = {f"{val.get('환자이름')} ({pid_key})": pid_key for pid_key, val in existing_patient_data.items() if isinstance(val, dict)}
            selected_patients_str = st.multiselect("삭제할 환자 선택:", list(patient_options.keys()), key="delete_patient_multiselect")
            patients_to_delete = [patient_options[name] for name in selected_patients_str]
            if patients_to_delete: st.session_state.patients_to_delete = patients_to_delete; st.session_state.delete_patient_confirm = True
            else: st.session_state.delete_patient_confirm = False
            
            if st.session_state.delete_patient_confirm:
                st.warning(f"⚠️ **{len(st.session_state.patients_to_delete)}명** 삭제?")
                if st.button("예, 삭제", key="confirm_delete_button"):
                    for pid_key in st.session_state.patients_to_delete: patients_ref_for_user.child(pid_key).delete()
                    st.session_state.delete_patient_confirm = False; st.session_state.patients_to_delete = []; st.success("삭제 완료"); st.rerun()

        st.markdown("---")
        with st.form("register_form"):
            name = st.text_input("환자명"); pid = st.text_input("진료번호"); selected_departments = st.multiselect("진료과", DEPARTMENTS_FOR_REGISTRATION)
            if st.form_submit_button("등록"):
                if name and pid and selected_departments:
                    pid_key = pid.strip(); new_patient_data = existing_patient_data.get(pid_key, {"환자이름": name, "진료번호": pid}) 
                    for dept_flag in PATIENT_DEPT_FLAGS + ['치주', '원진실']: new_patient_data[dept_flag.lower()] = False
                    for dept in selected_departments: new_patient_data[dept.lower()] = True
                    patients_ref_for_user.child(pid_key).set(new_patient_data); st.success("등록 완료"); st.rerun()
                else: st.warning("입력 확인")

    with analysis_tab:
        st.header("📈 OCS 분석 결과")
        analysis_results = db_ref_func("ocs_analysis/latest_result").get()
        latest_file_name = db_ref_func("ocs_analysis/latest_file_name").get()
        if analysis_results and latest_file_name:
            st.markdown(f"**<h3 style='text-align: left;'>{latest_file_name} 분석 결과</h3>**", unsafe_allow_html=True); st.markdown("---")
            for dept in ['소치', '보존', '교정']:
                if dept in analysis_results: st.subheader(f"{dept}"); st.info(f"오전: {analysis_results[dept]['오전']}명 / 오후: {analysis_results[dept]['오후']}명"); st.markdown("---")
        else: st.info("분석 결과 없음")
        st.divider(); st.header("🔑 비밀번호 변경")
        new_pw = st.text_input("새 비밀번호", type="password", key="u_new_pw"); cf_pw = st.text_input("확인", type="password", key="u_cf_pw")
        if st.button("변경", key="u_pw_chg_btn"):
            if new_pw and new_pw == cf_pw: users_ref.child(firebase_key).update({"password": hash_password(new_pw)}); st.success("변경 완료")
            else: st.error("불일치")

    with review_tab: show_professor_review_system()

# --- 5. 치과의사 모드 UI ---

def show_doctor_mode_ui(firebase_key, user_name):
    """치과의사 모드 UI를 표시합니다."""
    st.header(f"🧑‍⚕️Dr. {user_name}")
    st.subheader("🗓️ Google Calendar 연동")
    get_google_calendar_service(firebase_key) 
    if st.session_state.get('google_calendar_service'): st.success("✅ 캘린더 추가 기능 허용됨")
    else: st.info("인증 필요")
    
    st.markdown("---")
    st.header("🔑 비밀번호 변경")
    new_pw = st.text_input("새 비밀번호", type="password", key="d_new_pw"); cf_pw = st.text_input("확인", type="password", key="d_cf_pw")
    if st.button("변경", key="d_pw_chg_btn"):
        if new_pw and new_pw == cf_pw: doctor_users_ref.child(firebase_key).update({"password": hash_password(new_pw)}); st.success("변경 완료")
        else: st.error("불일치")
