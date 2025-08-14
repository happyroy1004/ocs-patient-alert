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

# 구글 캘린더 API 관련 라이브러리
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import datetime

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
            'databaseURL': st.secrets["firebase"]["database_url"]
        })
    except Exception as e:
        st.error(f"Firebase 초기화 오류: {e}")
        st.info("secrets.toml 파일의 Firebase 설정(FIREBASE_SERVICE_ACCOUNT_JSON 또는 database_url)을 [firebase] 섹션 아래에 올바르게 작성했는지 확인해주세요.")

# --- 구글 캘린더 API 설정 ---
# 필요한 스코프 정의. 캘린더 이벤트를 생성하고 수정하기 위해 필요합니다.
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

def get_google_calendar_service(email):
    """
    사용자 이메일에 해당하는 구글 캘린더 API 서비스를 반환합니다.
    사용자의 'credentials.json' 및 'token.json' 파일이 필요합니다.
    """
    creds = None
    # 이메일별로 token.json 파일을 관리하여 다중 사용자 지원
    token_file = f'token_{email.replace("@", "_").replace(".", "_")}.json'

    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)
    
    # 인증 정보가 없거나 만료된 경우 새로고침
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # secrets.toml에 저장된 클라이언트 인증 정보를 사용
            # secrets.toml에 "google_calendar" 섹션 아래에 'client_id'와 'client_secret'이 필요합니다.
            client_config = {
                "installed": {
                    "client_id": st.secrets["google_calendar"]["client_id"],
                    "client_secret": st.secrets["google_calendar"]["client_secret"],
                    "redirect_uris": ["urn:ietf:wg:oauth:2.0:oob"],
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://accounts.google.com/o/oauth2/token"
                }
            }
            flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
            
            # 인증 URL 생성 및 사용자에게 안내
            auth_url, _ = flow.authorization_url(prompt='consent')
            st.info(f"""
                **구글 캘린더 API를 사용하려면 인증이 필요합니다.**
                1. 다음 URL을 브라우저에 복사하여 붙여넣고 로그인하세요.
                
                {auth_url}
                
                2. 인증 후 받은 코드를 아래 입력창에 붙여넣어주세요.
            """)
            auth_code = st.text_input("인증 코드 입력")
            if auth_code:
                flow.fetch_token(code=auth_code)
                creds = flow.credentials
                # 다음 실행을 위해 token.json에 인증 정보 저장
                with open(token_file, 'w') as token:
                    token.write(creds.to_json())

    if creds:
        try:
            service = build('calendar', 'v3', credentials=creds)
            return service
        except HttpError as error:
            st.error(f"구글 캘린더 서비스 빌드 중 오류 발생: {error}")
            return None
    return None

def create_calendar_event(service, patient_name, pid, department):
    """
    구글 캘린더에 이벤트를 생성합니다.
    """
    event_start_time = datetime.datetime.now(datetime.timezone.utc).isoformat()
    event_end_time = (datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(hours=1)).isoformat()
    
    event = {
        'summary': f'환자 등록: {patient_name}',
        'location': f'진료번호: {pid}',
        'description': f'등록 과: {department}',
        'start': {
            'dateTime': event_start_time,
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': event_end_time,
            'timeZone': 'Asia/Seoul',
        },
    }
    
    try:
        event = service.events().insert(calendarId='primary', body=event).execute()
        st.success(f"'{patient_name}' 환자 등록 일정이 캘린더에 추가되었습니다.")
    except HttpError as error:
        st.error(f"캘린더 이벤트 생성 중 오류 발생: {error}")

# 메인 Streamlit 앱 로직
st.title("환자 관리 시스템")
st.markdown("---")

# 로그인 및 사용자 이메일 입력
st.header("사용자 정보")
user_email = st.text_input("구글 계정 이메일(Google Calendar 연동용)")
if user_email and not is_valid_email(user_email):
    st.error("유효한 이메일 주소를 입력해주세요.")

# Firebase 데이터베이스 참조 설정 (사용자 이메일 기반)
if user_email and is_valid_email(user_email):
    user_id = user_email.replace(".", "_").replace("@", "_")
    patients_ref_for_user = db.reference(f'patients/{user_id}')
    
    # 구글 캘린더 서비스 초기화
    google_calendar_service = get_google_calendar_service(user_email)

    # 기존 환자 목록 불러오기
    st.header("등록된 환자 목록")
    existing_patient_data = patients_ref_for_user.get()
    
    if existing_patient_data:
        sorted_patients = sorted(existing_patient_data.items(), key=lambda item: item[1].get('등록시간', 0), reverse=True)
        for key, val in sorted_patients:
            with st.container(border=True):
                info_col, btn_col = st.columns([4, 1])
                
                with info_col:
                    st.markdown(f"**{val['환자명']}** / {val['진료번호']} / {val.get('등록과', '미지정')}")
                
                with btn_col:
                    if st.button("X", key=f"delete_button_{key}"):
                        patients_ref_for_user.child(key).delete()
                        st.rerun()
    else:
        st.info("등록된 환자가 없습니다.")
    st.markdown("---")

    # 환자 등록 폼
    st.header("환자 등록")
    with st.form("register_form"):
        name = st.text_input("환자명")
        pid = st.text_input("진료번호")
        departments_for_registration = sorted(["내과", "외과", "소아과", "안과"]) # 예시 과목
        selected_department = st.selectbox("등록 과", departments_for_registration)

        submitted = st.form_submit_button("등록")
        if submitted:
            if not name or not pid:
                st.warning("모든 항목을 입력해주세요.")
            else:
                if existing_patient_data and any(
                    v["환자명"] == name and v["진료번호"] == pid and v.get("등록과") == selected_department
                    for v in existing_patient_data.values()):
                    st.error("이미 등록된 환자입니다.")
                else:
                    new_patient_data = {
                        "환자명": name,
                        "진료번호": pid,
                        "등록과": selected_department,
                        "등록시간": int(time.time())
                    }
                    patients_ref_for_user.push(new_patient_data)
                    st.success(f"{name} 환자가 등록되었습니다.")
                    
                    # 캘린더 이벤트 생성 (새로운 기능)
                    if google_calendar_service:
                        create_calendar_event(google_calendar_service, name, pid, selected_department)

                    st.rerun()
