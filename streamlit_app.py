#1. Imports, Validation Functions, and Firebase Initialization
import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, db
import io
import msoffcrypto
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from openpyxl import load_workbook
from openpyxl.styles import Font
import re
import json
import os
import time
import openpyxl 
import datetime 

# Google Calendar API 관련 라이브러리 추가
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import base64

# --- 파일명 유효성 검사 함수 (학생용 코드에서 가져옴) ---
def is_daily_schedule(file_name):
    """
    파일명이 'ocs_MMDD.xlsx' 또는 'ocs_MMDD.xlsm' 형식인지 확인합니다.
    """
    pattern = r'^ocs_\\d{4}\\.(?:xlsx|xlsm)$'
    return re.match(pattern, file_name, re.IGNORECASE) is not None
    

# --- 이메일 유효성 검사 함수 ---
def is_valid_email(email):
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$"
    return re.match(email_regex, email) is not None

# Firebase 초기화
if not firebase_admin._apps:
    try:
        firebase_credentials_json_str = st.secrets["firebase"]["FIREBASE_SERVICE_ACCOUNT_JSON"]
        firebase_credentials_dict = json.loads(firebase_credentials_json_str)

        cred = credentials.Certificate(firebase_credentials_dict)
        firebase_admin.initialize_app(cred, {
            'databaseURL': st.secrets["firebase"]["FIREBASE_DATABASE_URL"]
        })
    except Exception as e:
        st.error(f"Firebase 초기화 중 오류가 발생했습니다: {e}")
        st.stop()


#2. Session State Management and User Authentication
if "auth_status" not in st.session_state:
    st.session_state.auth_status = "unauthenticated"
if "current_user_email" not in st.session_state:
    st.session_state.current_user_email = ""
if "current_firebase_key" not in st.session_state:
    st.session_state.current_firebase_key = ""
if "uploaded_file" not in st.session_state:
    st.session_state.uploaded_file = None
if "user_role" not in st.session_state: # Add user_role state
    st.session_state.user_role = ""

def get_user_data(email, password):
    """Firebase에서 사용자 데이터를 조회합니다."""
    users_ref = db.reference("users")
    users = users_ref.order_by_child("email").equal_to(email).get()
    
    if not users:
        return None, None

    found_key = list(users.keys())[0]
    user_data = users[found_key]

    if user_data.get("password") == password:
        return user_data, found_key
    else:
        return None, None

def login():
    st.title("로그인")
    with st.form("login_form"):
        email = st.text_input("이메일", key="login_email")
        password = st.text_input("비밀번호", type="password", key="login_password")
        submitted = st.form_submit_button("로그인")
        
        if submitted:
            user_data, user_key = get_user_data(email, password)
            if user_data:
                st.session_state.auth_status = "authenticated"
                st.session_state.current_user_email = email
                st.session_state.current_firebase_key = user_key
                st.session_state.user_role = user_data.get("role", "일반 사용자") # 역할 정보 추가
                st.rerun()
            else:
                st.error("이메일 또는 비밀번호가 잘못되었습니다.")

def logout():
    """로그아웃 버튼을 클릭하면 세션을 초기화합니다."""
    if st.button("로그아웃"):
        for key in st.session_state.keys():
            del st.session_state[key]
        st.rerun()

#3. Main App Logic (If authenticated)
if st.session_state.auth_status == "authenticated":
    st.title(f"👋 환영합니다, {st.session_state.current_user_email}님!")
    st.write(f"현재 역할: {st.session_state.user_role}")
    logout()
    
    st.divider()

    # --- 엑셀 파일 업로드 섹션 (단일 업로드) ---
    st.header("엑셀 파일 업로드")
    uploaded_file = st.file_uploader("OCS 일일 스케줄 파일을 업로드하세요", type=['xlsx', 'xlsm'])
    
    if uploaded_file:
        st.session_state.uploaded_file = uploaded_file
        st.success(f"파일 '{uploaded_file.name}' 업로드 완료!")
    
    st.divider()

    # --- 탭을 이용한 분리된 기능 섹션 ---
    tab1, tab2 = st.tabs(["레지던트용 기능", "학생용 기능"])

    with tab1:
        st.header("레지던트용 기능")
        st.write("레지던트용 기능이 여기에 표시됩니다.")
        
        if st.session_state.uploaded_file:
            st.info(f"업로드된 파일: {st.session_state.uploaded_file.name}")
            # 여기에 레지던트용 파일 처리 및 분석 로직을 추가합니다.
            # 예: df = pd.read_excel(st.session_state.uploaded_file)
            #     st.dataframe(df.head())
        else:
            st.warning("파일을 먼저 업로드해주세요.")

    with tab2:
        st.header("학생용 기능")
        st.write("학생용 기능이 여기에 표시됩니다.")

        if st.session_state.uploaded_file:
            st.info(f"업로드된 파일: {st.session_state.uploaded_file.name}")
            # 여기에 학생용 파일 처리 및 분석 로직을 추가합니다.
            # 예: df = pd.read_excel(st.session_state.uploaded_file)
            #     st.dataframe(df.tail())
        else:
            st.warning("파일을 먼저 업로드해주세요.")
    
    st.divider()
    
    # --- 환자 등록 및 관리 기능 (요청 1) ---
    st.header("🏥 내 환자 관리")
    
    # 환자 추가 UI
    with st.expander("➕ 새 환자 등록", expanded=False):
        name = st.text_input("환자명", key="add_name")
        pid = st.text_input("진료번호", key="add_pid")
        selected_department = st.selectbox("등록과", ["외과", "내과", "소아과", "신경과"], key="add_department")

        if st.button("환자 등록"):
            if not name or not pid:
                st.error("환자명과 진료번호를 모두 입력해주세요.")
            else:
                patients_ref_for_user = db.reference(f"users/{st.session_state.current_firebase_key}/patients")
                existing_patient_data = patients_ref_for_user.get() or {}

                # 중복 확인
                is_duplicate = False
                for v in existing_patient_data.values():
                    if (v.get("환자명") == name and 
                        v.get("진료번호") == pid and 
                        v.get("등록과") == selected_department):
                        is_duplicate = True
                        break
                
                if is_duplicate:
                    st.error("이미 등록된 환자입니다.")
                else:
                    patients_ref_for_user.push().set({"환자명": name, "진료번호": pid, "등록과": selected_department})
                    st.success(f"{name} ({pid}) [{selected_department}] 환자 등록 완료")
                    st.rerun()
    
    # 환자 목록 표시 및 삭제 UI
    st.subheader("📋 등록된 환자 목록")
    patients_ref_for_user = db.reference(f"users/{st.session_state.current_firebase_key}/patients")
    existing_patient_data = patients_ref_for_user.get()

    if existing_patient_data:
        # 데이터프레임으로 변환하여 테이블로 보기 좋게 표시
        patient_list = []
        for key, value in existing_patient_data.items():
            value['key'] = key
            patient_list.append(value)
        
        # 컬럼 생성
        cols = st.columns([1, 1, 1, 0.2]) # 마지막 컬럼은 삭제 버튼용으로 작게 설정
        cols[0].write("**환자명**")
        cols[1].write("**진료번호**")
        cols[2].write("**등록과**")
        cols[3].write("") # 헤더 빈 칸

        for patient in patient_list:
            cols = st.columns([1, 1, 1, 0.2])
            cols[0].write(patient["환자명"])
            cols[1].write(patient["진료번호"])
            cols[2].write(patient["등록과"])
            
            # 삭제 버튼 (X 표시)
            if cols[3].button("❌", key=f"delete_{patient['key']}"):
                patients_ref_for_user.child(patient['key']).delete()
                st.success("환자 정보가 삭제되었습니다.")
                st.rerun()

    else:
        st.info("등록된 환자가 없습니다.")

    # --- 비밀번호 변경 기능 추가 ---
    if st.session_state.get("found_user_email"):
        st.divider()
        st.header("🔑 비밀번호 변경")
        
        new_password = st.text_input("새 비밀번호를 입력하세요", type="password", key="new_password_input")
        confirm_password = st.text_input("새 비밀번호를 다시 입력하세요", type="password", key="confirm_password_input")
        
        if st.button("비밀번호 변경"):
            if not new_password or not confirm_password:
                st.error("새 비밀번호와 확인용 비밀번호를 모두 입력해주세요.")
            elif new_password != confirm_password:
                st.error("새 비밀번호가 일치하지 않습니다. 다시 확인해주세요.")
            else:
                try:
                    users_ref = db.reference("users")
                    users_ref.child(st.session_state.current_firebase_key).update({"password": new_password})
                    st.success("비밀번호가 성공적으로 변경되었습니다.")
                except Exception as e:
                    st.error(f"비밀번호 변경 중 오류가 발생했습니다: {e}")


#3-1. App Entry Point
if st.session_state.auth_status == "unauthenticated":
    st.info("로그인이 필요합니다.")
    login()

#4. Excel Processing Constants and Functions
sheet_keyword_to_department_map = {
    '치과보철과': '보철', '보철과': '보철', '보철': '보철',
    '치과교정과' : '교정', '교정과': '교정', '교정': '교정',
    '구강 악안면외과' : '외과', '구강악안면외과': '외과', '외과': '외과',
    '구강 내과' : '내과', '구강내과': '내과', '내과': '내과',
    '치과보존과' : '보존', '보존과': '보존', '보존': '보존',
    '치주과' : '치주', '치주': '치주',
    '치과방사선과': '방사선', '방사선과': '방사선', '방사선': '방사선',
    '예방치과': '예방', '예방': '예방',
    '치과마취과': '마취', '마취과': '마취', '마취': '마취',
    '소아치과': '소치', '소아 치과': '소치', '소치': '소치'
}

# --- 메인 페이지 UI 구성 ---
st.title("👨‍💻 환자 내원 정보 관리")
users_ref = db.reference("users")
patients_ref = db.reference("patients")

# 로그인 폼
if not st.session_state.logged_in:
    st.subheader("로그인")
    user_name_input = st.text_input("사용자 이름")
    password_input = st.text_input("비밀번호", type="password")

    if st.button("로그인"):
        # 관리자 로그인 체크
        if user_name_input == "admin":
            st.session_state.user_role = "admin"
            st.session_state.logged_in = True
            st.session_state.found_user_email = "admin"
            st.success("관리자 모드로 로그인했습니다.")
            st.rerun()
        # 레지던트 로그인 체크
        elif user_name_input == "레지던트":
            st.session_state.user_role = "resident"
            st.session_state.logged_in = True
            st.session_state.found_user_email = "temp_resident_login" # 임시 이메일 할당
            st.session_state.current_firebase_key = "temp_resident_login"
            st.success("레지던트 전용 페이지로 이동합니다.")
            st.rerun()
        # 일반 사용자 로그인 체크
        else:
            try:
                user_data = users_ref.get()
                if not user_data:
                    st.error("등록된 사용자가 없습니다.")
                else:
                    found_user_key = None
                    for key, value in user_data.items():
                        if value.get("name") == user_name_input and value.get("password") == password_input:
                            found_user_key = key
                            st.session_state.found_user_email = value.get("email")
                            st.session_state.current_firebase_key = key
                            st.session_state.user_role = value.get("role", "student") # 역할 가져오기
                            break
                    
                    if found_user_key:
                        st.session_state.logged_in = True
                        st.success(f"{user_name_input}님, 로그인 성공!")
                        st.rerun()
                    else:
                        st.error("사용자 이름 또는 비밀번호가 올바르지 않습니다.")
            except Exception as e:
                st.error(f"로그인 중 오류 발생: {e}")
                
# --- 로그인 상태에 따른 페이지 분기 ---

# #9. 레지던트 전용 페이지
if st.session_state.logged_in and st.session_state.user_role == "resident":
    st.subheader("레지던트 정보 등록/수정")
    
    # 레지던트 이메일 입력 및 로그인 처리
    if st.session_state.found_user_email == "temp_resident_login":
        st.info("처음 로그인하셨습니다. 레지던트 계정을 등록해주세요.")
        resident_email_input = st.text_input("레지던트 이메일")
        resident_password_input = st.text_input("비밀번호", type="password")
        if st.button("레지던트 계정 등록/로그인"):
            if not resident_email_input or not is_valid_email(resident_email_input):
                st.error("유효한 이메일을 입력해주세요.")
            elif not resident_password_input:
                st.error("비밀번호를 입력해주세요.")
            else:
                user_key = sanitize_path(resident_email_input)
                user_data = users_ref.child(user_key).get()
                if user_data:
                    if user_data.get('password') == resident_password_input:
                        st.session_state.logged_in = True
                        st.session_state.user_role = "resident"
                        st.session_state.found_user_email = resident_email_input
                        st.session_state.current_firebase_key = user_key
                        st.success("레지던트 계정으로 로그인했습니다.")
                        st.rerun()
                    else:
                        st.error("비밀번호가 올바르지 않습니다.")
                else:
                    # 신규 등록
                    users_ref.child(user_key).set({
                        "email": resident_email_input,
                        "password": resident_password_input,
                        "role": "resident",
                        "name": "",
                        "department": ""
                    })
                    st.session_state.logged_in = True
                    st.session_state.user_role = "resident"
                    st.session_state.found_user_email = resident_email_input
                    st.session_state.current_firebase_key = user_key
                    st.success("새로운 레지던트 계정이 등록되었습니다. 정보를 입력해주세요.")
                    st.rerun()
    else:
        # 이미 로그인한 상태
        user_key = st.session_state.current_firebase_key
        user_data = users_ref.child(user_key).get()
        if user_data:
            st.session_state.resident_info['name'] = user_data.get('name', '')
            st.session_state.resident_info['department'] = user_data.get('department', '')
        
        resident_name_input = st.text_input("레지던트 이름", value=st.session_state.resident_info['name'])
        resident_dept_input = st.text_input("등록과", value=st.session_state.resident_info['department'])
        
        new_password = st.text_input("새 비밀번호 (변경 시)", type="password")
        confirm_new_password = st.text_input("새 비밀번호 확인", type="password")

        if st.button("정보 저장"):
            if not resident_name_input or not resident_dept_input:
                st.error("이름과 등록과는 필수 입력 항목입니다.")
            elif new_password and new_password != confirm_new_password:
                st.error("새 비밀번호가 일치하지 않습니다. 다시 확인해주세요.")
            else:
                update_data = {
                    "name": resident_name_input,
                    "department": resident_dept_input,
                }
                if new_password:
                    update_data["password"] = new_password
                
                users_ref.child(user_key).update(update_data)
                
                st.session_state.resident_info['name'] = resident_name_input
                st.session_state.resident_info['department'] = resident_dept_input
                st.success("레지던트 정보가 성공적으로 저장되었습니다.")
                st.rerun()

    st.divider()
    
    # 레지던트용 환자 등록
    st.subheader("레지던트 환자 등록")
    name = st.text_input("환자명", key="res_name")
    pid = st.text_input("진료번호 (PID)", key="res_pid")
    
    if st.button("환자 등록", key="res_register_patient"):
        if not st.session_state.resident_info['name'] or not st.session_state.resident_info['department']:
            st.error("환자 등록 전에 먼저 '레지던트 정보 등록/수정'에서 이름과 등록과를 입력해주세요.")
        elif not name or not pid:
            st.error("환자명과 진료번호를 모두 입력해주세요.")
        else:
            patients_ref_for_user = patients_ref.child(sanitize_path(st.session_state.found_user_email))
            existing_patient_data = patients_ref_for_user.get()
            if existing_patient_data is None: existing_patient_data = {}
            if any(v["환자명"] == name and v["진료번호"] == pid and v.get("등록과") == st.session_state.resident_info['department'] for v in existing_patient_data.values()):
                st.error("이미 등록된 환자입니다.")
            else:
                patients_ref_for_user.push().set({"환자명": name, "진료번호": pid, "등록과": st.session_state.resident_info['department']})
                st.success(f"{name} ({pid}) [{st.session_state.resident_info['department']}] 환자 등록 완료")
                st.rerun()

    # 등록된 환자 목록 보기
    st.subheader("등록된 환자 목록")
    if st.session_state.current_firebase_key:
        patients_ref_for_user = patients_ref.child(st.session_state.current_firebase_key)
        registered_patients_data = patients_ref_for_user.get()
        
        if registered_patients_data:
            patient_list = [{"환자명": v["환자명"], "진료번호": v["진료번호"], "등록과": v.get("등록과", "")} for v in registered_patients_data.values()]
            patient_df = pd.DataFrame(patient_list)
            st.dataframe(patient_df, use_container_width=True)
        else:
            st.info("아직 등록된 환자가 없습니다.")
    else:
        st.info("레지던트 계정을 등록하면 환자 목록이 여기에 표시됩니다.")

    # 구글 캘린더 연동
    st.divider()
    st.subheader("Google Calendar 연동")
    user_key = sanitize_path(st.session_state.found_user_email)
    google_calendar_service = get_google_calendar_service(user_key)
    if google_calendar_service:
        st.success("Google Calendar가 성공적으로 연동되었습니다.")
        st.info("엑셀 파일 업로드 시 일정이 자동으로 추가됩니다.")

# 학생 전용 페이지 (기존 기능 복원 및 유지)
elif st.session_state.logged_in and st.session_state.user_role == "student":
    st.subheader("환자 등록")
    name = st.text_input("환자명")
    pid = st.text_input("진료번호 (PID)")
    
    department_list = ["소치", "교정", "보존", "보철", "외과", "내과", "치주", "방사선", "예방", "마취"]
    selected_department = st.selectbox("등록과", department_list)
    
    if st.button("환자 등록"):
        if not name or not pid:
            st.error("환자명과 진료번호를 모두 입력해주세요.")
        else:
            patients_ref_for_user = patients_ref.child(st.session_state.current_firebase_key)
            existing_patient_data = patients_ref_for_user.get()
            if existing_patient_data is None:
                existing_patient_data = {}

            if any(v["환자명"] == name and v["진료번호"] == pid and v.get("등록과") == selected_department
                   for v in existing_patient_data.values()):
                st.error("이미 등록된 환자입니다.")
            else:
                patients_ref_for_user.push().set({"환자명": name, "진료번호": pid, "등록과": selected_department})
                st.success(f"{name} ({pid}) [{selected_department}] 환자 등록 완료")
                st.rerun()

    # 등록된 환자 목록 보기 (복원된 기능)
    st.subheader("등록된 환자 목록")
    if st.session_state.current_firebase_key:
        patients_ref_for_user = patients_ref.child(st.session_state.current_firebase_key)
        registered_patients_data = patients_ref_for_user.get()
        
        if registered_patients_data:
            patient_list = [{"환자명": v["환자명"], "진료번호": v["진료번호"], "등록과": v.get("등록과", "")} for v in registered_patients_data.values()]
            patient_df = pd.DataFrame(patient_list)
            st.dataframe(patient_df, use_container_width=True)
        else:
            st.info("아직 등록된 환자가 없습니다.")
    else:
        st.info("로그인하면 등록한 환자 목록이 여기에 표시됩니다.")
    
    # --- 비밀번호 변경 기능 추가 ---
    if st.session_state.get("found_user_email"):
        st.divider()
        st.header("🔑 비밀번호 변경")
        
        new_password = st.text_input("새 비밀번호를 입력하세요", type="password", key="new_password_input")
        confirm_password = st.text_input("새 비밀번호를 다시 입력하세요", type="password", key="confirm_password_input")
        
        if st.button("비밀번호 변경"):
            if not new_password or not confirm_password:
                st.error("새 비밀번호와 확인용 비밀번호를 모두 입력해주세요.")
            elif new_password != confirm_password:
                st.error("새 비밀번호가 일치하지 않습니다. 다시 확인해주세요.")
            else:
                try:
                    users_ref.child(st.session_state.current_firebase_key).update({"password": new_password})
                    st.success("비밀번호가 성공적으로 변경되었습니다.")
                except Exception as e:
                    st.error(f"비밀번호 변경 실패: {e}")

    # 구글 캘린더 연동
    st.divider()
    st.subheader("Google Calendar 연동")
    user_key = sanitize_path(st.session_state.found_user_email)
    google_calendar_service = get_google_calendar_service(user_key)
    
    if google_calendar_service:
        st.success("Google Calendar가 성공적으로 연동되었습니다.")
        st.info("엑셀 파일 업로드 시 일정이 자동으로 추가됩니다.")

# #7. 관리자 전용 페이지
elif st.session_state.logged_in and st.session_state.user_role == "admin":
    is_admin_input = True # 관리자 전용 페이지 진입을 위한 더미 변수
    
    # 두 가지 탭 생성
    student_tab, resident_tab = st.tabs(['학생 환자 관리', '레지던트 환자 관리'])

    with student_tab:
        st.subheader("💻 학생 환자 관리")
        uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])

        if uploaded_file:
            file_name = uploaded_file.name
            is_daily = is_daily_schedule(file_name)
            
            if is_daily: st.info("✔️ '확정된 당일 일정' 파일로 인식되었습니다. 기존 일정과 비교 후 업데이트합니다.")
            else: st.info("✔️ '예정된 전체 일정' 파일로 인식되었습니다. 모든 일정을 캘린더에 추가합니다.")
                
            uploaded_file.seek(0)
            password = st.text_input("엑셀 파일 비밀번호 입력", type="password", key="password_student") if is_encrypted_excel(uploaded_file) else None
            if is_encrypted_excel(uploaded_file) and not password:
                st.info("암호화된 파일입니다. 비밀번호를 입력해주세요.")
                st.stop()
            
            try:
                xl_object, raw_file_io = load_excel(uploaded_file, password)
                excel_data_dfs, styled_excel_bytes = process_excel_file_and_style(raw_file_io)
                
                professors_dict = {
                    '소치': ['김현태', '장기택', '김정욱', '현홍근', '김영재', '신터전', '송지수'],
                    '보존': ['이인복', '금기연', '이우철', '유연지', '서덕규', '이창하', '김선영', '손원준']
                }
                analysis_results = run_analysis(excel_data_dfs, professors_dict)
                today_date_str = datetime.datetime.now().strftime("%Y-%m-%d")
                db.reference("ocs_analysis/latest_result").set(analysis_results)
                db.reference("ocs_analysis/latest_date").set(today_date_str)
                db.reference("ocs_analysis/latest_file_name").set(file_name)
                
                st.session_state.last_processed_data = excel_data_dfs
                st.session_state.last_processed_file_name = file_name
                
                if excel_data_dfs is None or styled_excel_bytes is None:
                    st.warning("엑셀 파일 처리 중 문제가 발생했거나 처리할 데이터가 없습니다.")
                    st.stop()
                
                sender = st.secrets["gmail"]["sender"]
                sender_pw = st.secrets["gmail"]["app_password"]

                all_users_meta = users_ref.get()
                all_patients_data = patients_ref.get()

                matched_users = []
                
                if all_patients_data:
                    for uid_safe, registered_patients_for_this_user in all_patients_data.items():
                        user_meta = all_users_meta.get(uid_safe, {})
                        user_email = user_meta.get("email") or recover_email(uid_safe)
                        user_display_name = user_meta.get("name") or user_email
                        
                        registered_patients_data = []
                        if registered_patients_for_this_user:
                            for key, val in registered_patients_for_this_user.items():
                                registered_patients_data.append({
                                    "환자명": val.get("환자명", "").strip(),
                                    "진료번호": val.get("진료번호", "").strip().zfill(8),
                                    "등록과": val.get("등록과", "")
                                })
                        
                        matched_rows_for_user = []
                        for sheet_name_excel_raw, df_sheet in excel_data_dfs.items():
                            excel_sheet_name_lower = sheet_name_excel_raw.strip().lower()
                            excel_sheet_department = None
                            for keyword, department_name in sorted(sheet_keyword_to_department_map.items(), key=lambda item: len(item[0]), reverse=True):
                                if keyword.lower() in excel_sheet_name_lower:
                                    excel_sheet_department = department_name
                                    break
                            
                            if not excel_sheet_department: continue
                                
                            for _, excel_row in df_sheet.iterrows():
                                excel_patient_name = str(excel_row.get("환자명", "")).strip()
                                excel_patient_pid = str(excel_row.get("진료번호", "")).strip().zfill(8)
                                
                                for registered_patient in registered_patients_data:
                                    if (registered_patient["환자명"] == excel_patient_name and
                                            registered_patient["진료번호"] == excel_patient_pid and
                                            registered_patient["등록과"] == excel_sheet_department):
                                        
                                        matched_row_copy = excel_row.copy()
                                        matched_row_copy["시트"] = sheet_name_excel_raw
                                        matched_row_copy["등록과"] = excel_sheet_department
                                        matched_rows_for_user.append(matched_row_copy)
                                        break
                        if matched_rows_for_user:
                            combined_matched_df = pd.DataFrame(matched_rows_for_user)
                            matched_users.append({"email": user_email, "name": user_display_name, "data": combined_matched_df, "safe_key": uid_safe})

                if matched_users:
                    st.success(f"{len(matched_users)}명의 사용자(학생)와 일치하는 환자 발견됨.")
                    matched_user_list_for_dropdown = [f"{user['name']} ({user['email']})" for user in matched_users]
                    if 'select_all_matched_users_student' not in st.session_state: st.session_state.select_all_matched_users_student = False
                    select_all_matched_button = st.button("매칭된 사용자 모두 선택/해제", key="select_all_matched_btn_student")
                    if select_all_matched_button:
                        st.session_state.select_all_matched_users_student = not st.session_state.select_all_matched_users_student
                        st.rerun()
                    default_selection_matched = matched_user_list_for_dropdown if st.session_state.select_all_matched_users_student else []
                    selected_users_to_act = st.multiselect("액션을 취할 사용자 선택", matched_user_list_for_dropdown, default=default_selection_matched, key="matched_user_multiselect_student")
                    selected_matched_users_data = [user for user in matched_users if f"{user['name']} ({user['email']})" in selected_users_to_act]
                    
                    for user_match_info in selected_matched_users_data:
                        st.markdown(f"**수신자:** {user_match_info['name']} ({user_match_info['email']})")
                        st.dataframe(user_match_info['data'])
                    
                    mail_col, calendar_col = st.columns(2)
                    with mail_col:
                        if st.button("선택된 사용자에게 메일 보내기", key="mail_student"):
                            for user_match_info in selected_matched_users_data:
                                real_email = user_match_info['email']
                                df_matched = user_match_info['data']
                                user_name = user_match_info['name']
                                if not df_matched.empty:
                                    df_html = df_matched[['환자명', '진료번호', '예약의사', '진료내역', '예약시간']].to_html(index=False, escape=False)
                                    email_subject = "치과 예약 내원 정보"
                                    email_body = f"""<p>안녕하세요, {user_name}님.</p><p>오늘 예약된 환자 내원 정보입니다.</p>{df_html}<p>확인 부탁드립니다.</p>"""
                                    try:
                                        send_email(
                                            receiver=real_email, rows=df_matched, sender=sender, password=sender_pw, custom_message=email_body, date_str=today_date_str
                                        )
                                        st.success(f"**{user_name}**님 ({real_email})에게 예약 정보 이메일 전송 완료!")
                                    except Exception as e:
                                        st.error(f"**{user_name}**님 ({real_email})에게 이메일 전송 실패: {e}")
                                else:
                                    st.warning(f"**{user_name}**님에게 보낼 매칭 데이터가 없습니다.")

                    with calendar_col:
                        if st.button("선택된 사용자에게 Google Calendar 일정 추가", key="calendar_student"):
                            for user_match_info in selected_matched_users_data:
                                user_safe_key = user_match_info['safe_key']
                                user_email = user_match_info['email']
                                user_name = user_match_info['name']
                                df_matched = user_match_info['data']
                                creds = load_google_creds_from_firebase(user_safe_key)
                                if creds and creds.valid and not creds.expired:
                                    try:
                                        service = build('calendar', 'v3', credentials=creds)
                                        if not df_matched.empty:
                                            for _, row in df_matched.iterrows():
                                                patient_name = row.get('환자명', '')
                                                patient_pid = row.get('진료번호', '')
                                                department = row.get('등록과', '')
                                                doctor_name = row.get('예약의사', '')
                                                treatment_details = row.get('진료내역', '')
                                                reservation_date_raw = row.get('예약일시', '')
                                                reservation_time_raw = row.get('예약시간', '')
                                                is_datetime_invalid = (pd.isna(reservation_date_raw) or str(reservation_date_raw).strip() == "" or pd.isna(reservation_time_raw) or str(reservation_time_raw).strip() == "")
                                                if is_datetime_invalid:
                                                    st.warning(f"⚠️ {patient_name} 환자의 날짜/시간 데이터가 비어 있습니다. 일정 추가를 건너뜁니다.")
                                                    continue
                                                date_str_to_parse = str(reservation_date_raw).strip()
                                                time_str_to_parse = str(reservation_time_raw).strip()
                                                try:
                                                    full_datetime_str = f"{date_str_to_parse} {time_str_to_parse}"
                                                    reservation_datetime = datetime.datetime.strptime(full_datetime_str, '%Y/%m/%d %H:%M')
                                                except ValueError as e:
                                                    st.error(f"❌ {patient_name} 환자의 날짜/시간 형식 파싱 최종 실패: {e}. 일정 추가를 건너뜁니다.")
                                                    continue
                                                event_prefix = "별표 내원 : " if is_daily else "내원? : "
                                                event_title = f"{event_prefix}{patient_name} ({department}, {doctor_name})"
                                                event_description = f"환자명 : {patient_name}\n진료번호 : {patient_pid}\n진료내역 : {treatment_details}"
                                                create_calendar_event(service, event_title, patient_pid, department, reservation_datetime, doctor_name, event_description)
                                            st.success(f"**{user_name}**님의 캘린더에 일정을 추가했습니다.")
                                    except Exception as e:
                                        st.error(f"**{user_name}**님의 캘린더 일정 추가 실패: {e}")
                                else:
                                    st.warning(f"**{user_name}**님은 Google Calendar 계정이 연동되어 있지 않습니다. Google Calendar 탭에서 인증을 진행해주세요.")
                else:
                    st.info("엑셀 파일 처리 완료. 매칭된 환자가 없습니다.")
                    
                output_filename = uploaded_file.name.replace(".xlsx", "_processed.xlsx").replace(".xlsm", "_processed.xlsm")
                st.download_button("처리된 엑셀 다운로드", data=styled_excel_bytes, file_name=output_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            except ValueError as ve:
                st.error(f"파일 처리 실패: {ve}")
            except Exception as e:
                st.error(f"예상치 못한 오류 발생: {e}")
    
    with resident_tab:
        st.subheader("💻 레지던트 환자 관리")
        uploaded_file_res = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"], key="upload_res")
        
        if uploaded_file_res:
            file_name = uploaded_file_res.name
            is_daily = is_daily_schedule(file_name)
            
            if is_daily: st.info("✔️ '확정된 당일 일정' 파일로 인식되었습니다. 기존 일정과 비교 후 업데이트합니다.")
            else: st.info("✔️ '예정된 전체 일정' 파일로 인식되었습니다. 모든 일정을 캘린더에 추가합니다.")
                
            uploaded_file_res.seek(0)
            password = st.text_input("엑셀 파일 비밀번호 입력", type="password", key="password_res") if is_encrypted_excel(uploaded_file_res) else None
            if is_encrypted_excel(uploaded_file_res) and not password:
                st.info("암호화된 파일입니다. 비밀번호를 입력해주세요.")
                st.stop()
            
            try:
                xl_object, raw_file_io = load_excel(uploaded_file_res, password)
                excel_data_dfs, styled_excel_bytes = process_excel_file_and_style(raw_file_io)
                
                if excel_data_dfs is None or styled_excel_bytes is None:
                    st.warning("엑셀 파일 처리 중 문제가 발생했거나 처리할 데이터가 없습니다.")
                    st.stop()
                
                sender = st.secrets["gmail"]["sender"]
                sender_pw = st.secrets["gmail"]["app_password"]

                all_users_meta = users_ref.get()
                
                # 레지던트 역할 사용자 필터링
                resident_users = {
                    key: value for key, value in (all_users_meta.items() if all_users_meta else {}) 
                    if value.get('role') == 'resident' and value.get('name') and value.get('department')
                }
                
                matched_residents = []
                
                for uid_safe, resident_info in resident_users.items():
                    resident_name = resident_info.get("name")
                    resident_dept = resident_info.get("department")
                    resident_email = resident_info.get("email")

                    matched_rows_for_resident = []
                    
                    for sheet_name_excel_raw, df_sheet in excel_data_dfs.items():
                        df_sheet['예약의사'] = df_sheet['예약의사'].astype(str).str.strip()
                        
                        # 레지던트 이름과 진료과가 일치하는 행 필터링
                        matched_df = df_sheet[
                            (df_sheet['예약의사'] == resident_name) &
                            (df_sheet['진료과'].str.strip() == resident_dept)
                        ]
                        
                        if not matched_df.empty:
                            matched_df = matched_df.copy()
                            matched_df["시트"] = sheet_name_excel_raw
                            matched_df["등록과"] = resident_dept
                            matched_rows_for_resident.append(matched_df)
                    
                    if matched_rows_for_resident:
                        combined_matched_df = pd.concat(matched_rows_for_resident)
                        matched_residents.append({"email": resident_email, "name": resident_name, "data": combined_matched_df, "safe_key": uid_safe})
                
                if matched_residents:
                    st.success(f"{len(matched_residents)}명의 레지던트와 일치하는 환자 발견됨.")
                    matched_resident_list_for_dropdown = [f"{res['name']} ({res['email']})" for res in matched_residents]
                    selected_residents_to_act = st.multiselect("액션을 취할 레지던트 선택", matched_resident_list_for_dropdown, key="matched_res_multiselect")
                    selected_matched_residents_data = [res for res in matched_residents if f"{res['name']} ({res['email']})" in selected_residents_to_act]
                    
                    for res_match_info in selected_matched_residents_data:
                        st.markdown(f"**수신자:** {res_match_info['name']} ({res_match_info['email']})")
                        st.dataframe(res_match_info['data'])
                    
                    mail_col, calendar_col = st.columns(2)
                    with mail_col:
                        if st.button("선택된 레지던트에게 메일 보내기", key="mail_resident"):
                            for res_match_info in selected_matched_residents_data:
                                real_email = res_match_info['email']
                                df_matched = res_match_info['data']
                                res_name = res_match_info['name']
                                if not df_matched.empty:
                                    df_html = df_matched[['환자명', '진료번호', '예약의사', '진료내역', '예약시간']].to_html(index=False, escape=False)
                                    email_subject = "치과 예약 내원 정보 (레지던트용)"
                                    email_body = f"""<p>안녕하세요, {res_name} 레지던트님.</p><p>오늘 예약된 환자 내원 정보입니다.</p>{df_html}<p>확인 부탁드립니다.</p>"""
                                    try:
                                        send_email(receiver=real_email, rows=df_matched, sender=sender, password=sender_pw, custom_message=email_body, date_str=today_date_str)
                                        st.success(f"**{res_name}** 레지던트님 ({real_email})에게 예약 정보 이메일 전송 완료!")
                                    except Exception as e:
                                        st.error(f"**{res_name}** 레지던트님 ({real_email})에게 이메일 전송 실패: {e}")
                                else:
                                    st.warning(f"**{res_name}** 레지던트님에게 보낼 매칭 데이터가 없습니다.")

                    with calendar_col:
                        if st.button("선택된 레지던트에게 Google Calendar 일정 추가", key="calendar_resident"):
                            for res_match_info in selected_matched_residents_data:
                                user_safe_key = res_match_info['safe_key']
                                user_email = res_match_info['email']
                                user_name = res_match_info['name']
                                df_matched = res_match_info['data']
                                creds = load_google_creds_from_firebase(user_safe_key)
                                if creds and creds.valid and not creds.expired:
                                    try:
                                        service = build('calendar', 'v3', credentials=creds)
                                        if not df_matched.empty:
                                            for _, row in df_matched.iterrows():
                                                patient_name = row.get('환자명', '')
                                                patient_pid = row.get('진료번호', '')
                                                department = row.get('등록과', '')
                                                doctor_name = row.get('예약의사', '')
                                                treatment_details = row.get('진료내역', '')
                                                reservation_date_raw = row.get('예약일시', '')
                                                reservation_time_raw = row.get('예약시간', '')
                                                is_datetime_invalid = (pd.isna(reservation_date_raw) or str(reservation_date_raw).strip() == "" or pd.isna(reservation_time_raw) or str(reservation_time_raw).strip() == "")
                                                if is_datetime_invalid: continue
                                                date_str_to_parse = str(reservation_date_raw).strip()
                                                time_str_to_parse = str(reservation_time_raw).strip()
                                                try:
                                                    full_datetime_str = f"{date_str_to_parse} {time_str_to_parse}"
                                                    reservation_datetime = datetime.datetime.strptime(full_datetime_str, '%Y/%m/%d %H:%M')
                                                except ValueError as e: continue
                                                event_prefix = "별표 내원 : " if is_daily else "내원? : "
                                                event_title = f"{event_prefix}{patient_name} ({department}, {doctor_name})"
                                                event_description = f"환자명 : {patient_name}\n진료번호 : {patient_pid}\n진료내역 : {treatment_details}"
                                                create_calendar_event(service, event_title, patient_pid, department, reservation_datetime, doctor_name, event_description)
                                            st.success(f"**{user_name}** 레지던트님의 캘린더에 일정을 추가했습니다.")
                                    except Exception as e:
                                        st.error(f"**{user_name}** 레지던트님의 캘린더 일정 추가 실패: {e}")
                                else:
                                    st.warning(f"**{user_name}** 레지던트님은 Google Calendar 계정이 연동되어 있지 않습니다. Google Calendar 탭에서 인증을 진행해주세요.")
                else:
                    st.info("엑셀 파일 처리 완료. 매칭된 레지던트가 없습니다.")
                    
                output_filename = uploaded_file_res.name.replace(".xlsx", "_processed.xlsx").replace(".xlsm", "_processed.xlsm")
                st.download_button("처리된 엑셀 다운로드", data=styled_excel_bytes, file_name=output_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except ValueError as ve:
                st.error(f"파일 처리 실패: {ve}")
            except Exception as e:
                st.error(f"예상치 못한 오류 발생: {e}")
                
    # 일반 관리자 모드
    st.markdown("---")
    st.subheader("🛠️ Administer password")
    admin_password_input = st.text_input("관리자 비밀번호를 입력하세요", type="password", key="admin_password")
    try: secret_admin_password = st.secrets["admin"]["password"]
    except KeyError:
        secret_admin_password = None
        st.error("⚠️ secrets.toml 파일에 'admin.password' 설정이 없습니다. 개발자에게 문의하세요.")
    if admin_password_input and admin_password_input == secret_admin_password:
        st.session_state.admin_password_correct = True
        st.success("관리자 권한이 활성화되었습니다.")
        
        st.markdown("---")
        st.subheader("📦 메일 발송")
        all_users_meta = users_ref.get()
        user_list_for_dropdown = [f"{user_info.get('name', '이름 없음')} ({user_info.get('email', '이메일 없음')})" for user_info in (all_users_meta.values() if all_users_meta else [])]
        if 'select_all_users' not in st.session_state: st.session_state.select_all_users = False
        select_all_users_button = st.button("모든 사용자 선택/해제", key="select_all_btn")
        if select_all_users_button:
            st.session_state.select_all_users = not st.session_state.select_all_users
            st.rerun()
        default_selection = user_list_for_dropdown if st.session_state.select_all_users else []
        selected_users_for_mail = st.multiselect("보낼 사용자 선택", user_list_for_dropdown, default=default_selection, key="mail_multiselect")
        custom_message = st.text_area("보낼 메일 내용", height=200)
        if st.button("메일 보내기"):
            if custom_message:
                sender = st.secrets["gmail"]["sender"]
                sender_pw = st.secrets["gmail"]["app_password"]
                email_list = []
                if selected_users_for_mail:
                    for user_str in selected_users_for_mail:
                        match = re.search(r'\((.*?)\)', user_str)
                        if match: email_list.append(match.group(1))
                if email_list:
                    with st.spinner("메일 전송 중..."):
                        for email in email_list:
                            result = send_email(receiver=email, rows=None, sender=sender, password=sender_pw, date_str=None, custom_message=custom_message)
                            if result is True: st.success(f"{email}로 메일 전송 완료!")
                            else: st.error(f"{email}로 메일 전송 실패: {result}")
                else: st.warning("메일 내용을 입력했으나, 선택된 사용자가 없습니다. 전송이 진행되지 않았습니다.")
            else: st.warning("메일 내용을 입력해주세요.")
        
        st.markdown("---")
        st.subheader("🗑️ 사용자 삭제")
        if 'delete_confirm' not in st.session_state: st.session_state.delete_confirm = False
        if 'users_to_delete' not in st.session_state: st.session_state.users_to_delete = []
        if not st.session_state.delete_confirm:
            users_to_delete = st.multiselect("삭제할 사용자 선택", user_list_for_dropdown, key="delete_user_multiselect")
            if st.button("선택한 사용자 삭제"):
                if users_to_delete:
                    st.session_state.delete_confirm = True
                    st.session_state.users_to_delete = users_to_delete
                    st.rerun()
                else: st.warning("삭제할 사용자를 선택해주세요.")
        else:
            st.warning("정말로 선택한 사용자를 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("예, 삭제합니다"):
                    for user_to_del_str in st.session_state.users_to_delete:
                        match = re.search(r'\((.*?)\)', user_to_del_str)
                        if match:
                            email_to_del = match.group(1)
                            safe_key_to_del = sanitize_path(email_to_del)
                            db.reference(f"users/{safe_key_to_del}").delete()
                            db.reference(f"patients/{safe_key_to_del}").delete()
                    st.success(f"사용자 {', '.join(st.session_state.users_to_delete)} 삭제 완료.")
                    st.session_state.delete_confirm = False
                    st.session_state.users_to_delete = []
                    st.rerun()
            with col2:
                if st.button("아니오, 취소합니다"):
                    st.session_state.delete_confirm = False
                    st.session_state.users_to_delete = []
                    st.rerun()
    elif admin_password_input and admin_password_input != secret_admin_password:
        st.error("비밀번호가 틀렸습니다.")
        st.session_state.admin_password_correct = False
