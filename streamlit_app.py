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
import openpyxl # 추가
import datetime # 추가

# Google Calendar API 관련 라이브러리 추가
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import base64

# --- 파일명 유효성 검사 함수 ---
def is_daily_schedule(file_name):
    """
    파일명이 'ocs_MMDD.xlsx' 또는 'ocs_MMDD.xlsm' 형식인지 확인합니다.
    """
    pattern = r'^ocs_\d{4}\.(?:xlsx|xlsm)$'
    return re.match(pattern, file_name, re.IGNORECASE) is not None

# --- 이메일 유효성 검사 함수 ---
def is_valid_email(email):
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(email_regex, email) is not None

# Firebase 초기화
if not firebase_admin._apps:
    try:
        firebase_credentials_json_str = st.secrets["firebase"]["FIREBASE_SERVICE_ACCOUNT_JSON"]
        firebase_credentials_dict = json.loads(firebase_credentials_json_str)

        cred = credentials.Certificate(firebase_credentials_dict)
        firebase_admin.initialize_app(cred, {
            'databaseURL': st.secrets["firebase"]["database_url"]
        })
    except Exception as e:
        st.error(f"Firebase 초기화 오류: {e}")
        st.info("secrets.toml 파일의 Firebase 설정(FIREBASE_SERVICE_ACCOUNT_JSON 또는 database_url)을 [firebase] 섹션 아래에 올바르게 작성했는지 확인해주세요.")
        st.stop()

# Firebase-safe 경로 변환 (이메일을 Firebase 키로 사용하기 위해)
def sanitize_path(email):
    return email.replace(".", "_dot_").replace("@", "_at_")

# 이메일 주소 복원 (Firebase 안전 키에서 원래 이메일로)
def recover_email(safe_id: str) -> str:
    email = safe_id.replace("_at_", "@").replace("_dot_", ".").replace("_com", ".com")
    return email

# 구글 캘린더 인증 정보를 Firebase에 저장
def save_google_creds_to_firebase(user_id_safe, creds):
    try:
        creds_ref = db.reference(f"users/{user_id_safe}/google_creds")
        creds_ref.set({
            'token': creds.token,
            'refresh_token': creds.refresh_token,
            'token_uri': creds.token_uri,
            'client_id': creds.client_id,
            'client_secret': creds.client_secret,
            'scopes': creds.scopes,
            'id_token': creds.id_token
        })
        return True
    except Exception as e:
        st.error(f"Failed to save Google credentials: {e}")
        return False

# Firebase에서 구글 캘린더 인증 정보를 불러오기
def load_google_creds_from_firebase(user_id_safe):
    try:
        creds_ref = db.reference(f"users/{user_id_safe}/google_creds")
        creds_data = creds_ref.get()
        if creds_data and 'token' in creds_data:
            creds = Credentials(
                token=creds_data.get('token'),
                refresh_token=creds_data.get('refresh_token'),
                token_uri=creds_data.get('token_uri'),
                client_id=creds_data.get('client_id'),
                client_secret=creds_data.get('client_secret'),
                scopes=creds_data.get('scopes'),
                id_token=creds_data.get('id_token')
            )
            return creds
        return None
    except Exception as e:
        st.error(f"Failed to load Google credentials: {e}")
        return None

# --- OCS 분석 관련 함수 추가 ---
def is_encrypted_excel(file_path):
    try:
        with openpyxl.open(file_path, read_only=True) as wb:
            return False
    except openpyxl.utils.exceptions.InvalidFileException:
        return True
    except Exception:
        return False

def load_excel(uploaded_file, password=None):
    try:
        file_io = io.BytesIO(uploaded_file.getvalue())
        wb = load_workbook(file_io, data_only=True)
        return wb, file_io
    except Exception as e:
        st.error(f"엑셀 파일 로드 중 오류 발생: {e}")
        return None, None
    
def process_excel_file_and_style(file_io):
    try:
        raw_df = pd.read_excel(file_io)
        excel_data_dfs = pd.read_excel(file_io, sheet_name=None)
        return excel_data_dfs, raw_df.to_excel(index=False, header=True, engine='xlsxwriter')
    except Exception as e:
        st.error(f"엑셀 데이터 처리 및 스타일링 중 오류 발생: {e}")
        return None, None
    
def run_analysis(df_dict, professors_dict):
    analysis_results = {}
    sheet_department_map = {
        '소치': '소치', '소아치과': '소치', '소아 치과': '소치',
        '보존': '보존', '보존과': '보존', '치과보존과': '보존',
        '교정': '교정', '교정과': '교정', '치과교정과': '교정'
    }
    mapped_dfs = {}
    for sheet_name, df in df_dict.items():
        processed_sheet_name = sheet_name.replace(" ", "").lower()
        for key, dept in sheet_department_map.items():
            if processed_sheet_name == key.replace(" ", "").lower():
                mapped_dfs[dept] = df
                break
    
    if '소치' in mapped_dfs:
        df = mapped_dfs['소치']
        non_professors_df = df[~df['예약의사'].isin(professors_dict.get('소치', []))]
        non_professors_df['예약시간'] = non_professors_df['예약시간'].astype(str).str.strip()
        non_professors_df = non_professors_df[non_professors_df['예약시간'] != 'nan']
        morning_patients = non_professors_df[(non_professors_df['예약시간'] >= '08:00') & (non_professors_df['예약시간'] <= '12:50')].shape[0]
        afternoon_patients = non_professors_df[non_professors_df['예약시간'] >= '13:00'].shape[0]
        if afternoon_patients > 0: afternoon_patients -= 1
        analysis_results['소치'] = {'오전': morning_patients, '오후': afternoon_patients}

    if '보존' in mapped_dfs:
        df = mapped_dfs['보존']
        non_professors_df = df[~df['예약의사'].isin(professors_dict.get('보존', []))]
        non_professors_df['예약시간'] = non_professors_df['예약시간'].astype(str).str.strip()
        non_professors_df = non_professors_df[non_professors_df['예약시간'] != 'nan']
        morning_patients = non_professors_df[(non_professors_df['예약시간'] >= '08:00') & (non_professors_df['예약시간'] <= '12:30')].shape[0]
        afternoon_patients = non_professors_df[non_professors_df['예약시간'] >= '12:50'].shape[0]
        if afternoon_patients > 0: afternoon_patients -= 1
        analysis_results['보존'] = {'오전': morning_patients, '오후': afternoon_patients}

    if '교정' in mapped_dfs:
        df = mapped_dfs['교정']
        bonding_patients_df = df[
            df['진료내역'].str.contains('bonding|본딩', case=False, na=False) &
            ~df['진료내역'].str.contains('debonding', case=False, na=False)
        ]
        bonding_patients_df['예약시간'] = bonding_patients_df['예약시간'].astype(str).str.strip()
        morning_bonding_patients = bonding_patients_df[(bonding_patients_df['예약시간'] >= '08:00') & (bonding_patients_df['예약시간'] <= '12:30')].shape[0]
        afternoon_bonding_patients = bonding_patients_df[bonding_patients_df['예약시간'] >= '12:50'].shape[0]
        analysis_results['교정'] = {'오전': morning_bonding_patients, '오후': afternoon_bonding_patients}
        
    return analysis_results

# --- 세션 상태 초기화 ---
if "clear" in st.query_params and st.query_params["clear"] == "true":
    st.session_state.clear()
    st.query_params["clear"] = "false"
    st.rerun()

if 'user_role' not in st.session_state:
    st.session_state.user_role = 'guest'
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'found_user_email' not in st.session_state:
    st.session_state.found_user_email = None
if 'current_firebase_key' not in st.session_state:
    st.session_state.current_firebase_key = None
if 'google_creds' not in st.session_state:
    st.session_state['google_creds'] = {}
if 'last_processed_file_name' not in st.session_state:
    st.session_state.last_processed_file_name = None
if 'last_processed_data' not in st.session_state:
    st.session_state.last_processed_data = None
if 'resident_info' not in st.session_state:
    st.session_state.resident_info = {'name': '', 'department': '', 'email': ''}

users_ref = db.reference("users")
patients_ref = db.reference("patients")

#2. Excel and Email Processing Functions
def is_encrypted_excel(file):
    try:
        file.seek(0)
        return msoffcrypto.OfficeFile(file).is_encrypted()
    except Exception:
        return False

def load_excel(file, password=None):
    try:
        file.seek(0)
        office_file = msoffcrypto.OfficeFile(file)
        if office_file.is_encrypted():
            if not password:
                raise ValueError("암호화된 파일입니다. 비밀번호를 입력해주세요.")
            decrypted = io.BytesIO()
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            return pd.ExcelFile(decrypted), decrypted
        else:
            return pd.ExcelFile(file), file
    except Exception as e:
        raise ValueError(f"엑셀 로드 또는 복호화 실패: {e}")

def send_email(receiver, rows, sender, password, date_str=None, custom_message=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver
        if custom_message:
            msg['Subject'] = "단체 메일 알림"
            body = custom_message
        else:
            subject_prefix = ""
            if date_str:
                subject_prefix = f"{date_str}일에 내원하는 "
            msg['Subject'] = f"{subject_prefix}등록 환자 내원 알림"
            html_table = rows.to_html(index=False, escape=False)
            style = """
            <style>
                table {
                    width: 100%; max-width: 100%;
                    border-collapse: collapse;
                    font-family: Arial, sans-serif;
                    font-size: 14px;
                    table-layout: fixed;
                }
                th, td {
                    border: 1px solid #dddddd; text-align: left;
                    padding: 8px;
                    vertical-align: top;
                    word-wrap: break-word;
                    word-break: break-word;
                }
                th {
                    background-color: #f2f2f2; font-weight: bold;
                    white-space: nowrap;
                }
                tr:nth-child(even) {
                    background-color: #f9f9f9;
                }
                .table-container {
                    overflow-x: auto; -webkit-overflow-scrolling: touch;
                }
            </style>
            """
            body = f"다음 토탈 환자가 내일 내원예정입니다:<br><br><div class='table-container'>{style}{html_table}</div>"
        
        msg.attach(MIMEText(body, 'html'))
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        return str(e)


#3. Google Calendar API Functions
SCOPES = ["https://www.googleapis.com/auth/calendar.events"]

def get_google_calendar_service(user_id_safe):
    creds = st.session_state.get(f"google_creds_{user_id_safe}")
    if not creds:
        creds = load_google_creds_from_firebase(user_id_safe)
        if creds: st.session_state[f"google_creds_{user_id_safe}"] = creds

    client_config = {
        "web": {
            "client_id": st.secrets["google_calendar"]["client_id"],
            "client_secret": st.secrets["google_calendar"]["client_secret"],
            "redirect_uris": [st.secrets["google_calendar"]["redirect_uri"]],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs"
        }
    }
    flow = InstalledAppFlow.from_client_config(client_config, SCOPES, redirect_uri=st.secrets["google_calendar"]["redirect_uri"])
    
    if not creds:
        auth_code = st.query_params.get("code")
        if auth_code:
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            st.session_state[f"google_creds_{user_id_safe}"] = creds
            save_google_creds_to_firebase(user_id_safe, creds)
            st.success("Google Calendar 인증이 완료되었습니다.")
            st.query_params.clear()
            st.rerun()
        else:
            auth_url, _ = flow.authorization_url(prompt='consent')
            st.warning("Google Calendar 연동을 위해 인증이 필요합니다. 아래 링크를 클릭하여 권한을 부여하세요.")
            st.markdown(f"**[Google Calendar 인증 링크]({auth_url})**")
            return None

    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        st.session_state[f"google_creds_{user_id_safe}"] = creds
        save_google_creds_to_firebase(user_id_safe, creds)

    try:
        service = build('calendar', 'v3', credentials=creds)
        return service
    except HttpError as error:
        st.error(f'Google Calendar 서비스 생성 실패: {error}')
        st.session_state.pop(f"google_creds_{user_id_safe}", None)
        db.reference(f"users/{user_id_safe}/google_creds").delete()
        return None

def create_calendar_event(service, patient_name, pid, department, reservation_datetime, doctor_name, treatment_details):
    seoul_tz = datetime.timezone(datetime.timedelta(hours=9))
    event_start = reservation_datetime.replace(tzinfo=seoul_tz)
    event_end = event_start + datetime.timedelta(minutes=30)
    summary_text = f'{patient_name}'
    event = {
        'summary': summary_text,
        'location': pid,
        'description': f"{treatment_details}\n",
        'start': {
            'dateTime': event_start.isoformat(),
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': event_end.isoformat(),
            'timeZone': 'Asia/Seoul',
        },
    }
    try:
        service.events().insert(calendarId='primary', body=event).execute()
        st.success(f"'{patient_name}' 환자의 캘린더 일정이 추가되었습니다.")
    except HttpError as error:
        st.error(f"캘린더 이벤트 생성 중 오류 발생: {error}")
        st.warning("구글 캘린더 인증 권한을 다시 확인해주세요.")
    except Exception as e:
        st.error(f"알 수 없는 오류 발생: {e}")

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
