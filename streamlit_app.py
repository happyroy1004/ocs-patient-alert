# -*- coding: utf-8 -*-

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
import datetime
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

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
        st.stop()

# --- 이메일 전송 함수 ---
def send_email(to_email, subject, body):
    try:
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        sender_email = st.secrets["gmail"]["email"]
        sender_password = st.secrets["gmail"]["password"]

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        st.success(f"이메일이 {to_email}로 성공적으로 전송되었습니다.")
        return True
    except Exception as e:
        st.error(f"이메일 전송 실패: {e}")
        return False

# --- Google Calendar API 관련 설정 및 함수 ---
try:
    client_id = st.secrets["googlecalendar"]["client_id"]
    client_secret = st.secrets["googlecalendar"]["client_secret"]
    redirect_uri = st.secrets["googlecalendar"]["redirect_uri"]
except KeyError:
    st.error("`secrets.toml` 파일에 Google Calendar 설정이 누락되었습니다. 파일을 확인해 주세요.")
    st.stop()

SCOPES = ['https://www.googleapis.com/auth/calendar.events']

def get_google_calendar_service(refresh_token=None):
    creds = None
    if refresh_token:
        creds = Credentials(
            token=None,
            refresh_token=refresh_token,
            token_uri="https://oauth2.googleapis.com/token",
            client_id=client_id,
            client_secret=client_secret
        )

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_config(
                {
                    "installed": {
                        "client_id": client_id,
                        "client_secret": client_secret,
                        "redirect_uris": [redirect_uri],
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token"
                    }
                },
                SCOPES
            )
            authorization_url, _ = flow.authorization_url(prompt='consent')
            st.markdown(f"[Google 계정 연동하기]({authorization_url})")
            auth_code = st.text_input("위 링크를 클릭하여 인증 코드를 입력하세요.")
            if auth_code:
                flow.fetch_token(code=auth_code)
                creds = flow.credentials
                st.session_state["google_refresh_token"] = creds.refresh_token

    if creds:
        return build('calendar', 'v3', credentials=creds)
    return None

def create_event(service, start_time, end_time, summary, description):
    event = {
        'summary': summary,
        'description': description,
        'start': {
            'dateTime': start_time.isoformat(),
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': end_time.isoformat(),
            'timeZone': 'Asia/Seoul',
        },
    }
    try:
        event = service.events().insert(calendarId='primary', body=event).execute()
        st.success(f"Google Calendar에 이벤트가 생성되었습니다: {event.get('htmlLink')}")
    except Exception as e:
        st.error(f"이벤트 생성 실패: {e}")


# --- Excel 파일 처리 관련 함수 ---
def decrypt_and_read_excel(file, password):
    try:
        decrypted_file = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(file)
        office_file.decrypt(decrypted_file, password=password)
        df = pd.read_excel(decrypted_file, engine='openpyxl')
        return df
    except msoffcrypto.exceptions.InvalidKeyError:
        st.error("잘못된 비밀번호입니다.")
        return None
    except Exception as e:
        st.error(f"파일 처리 중 오류 발생: {e}")
        return None

def find_pid_in_dataframe(df, pid_list):
    pid_column = '진료번호' # 엑셀 파일의 '진료번호' 컬럼
    pid_in_df = df[df[pid_column].isin(pid_list)]
    return pid_in_df

# --- Streamlit UI ---
st.set_page_config(layout="wide")
st.title("OCS 환자 알림 시스템")
st.caption("환자 정보 관리 및 진료 알림을 위한 앱입니다.")

# --- Firebase에서 환자 데이터 가져오기 ---
ref = db.reference('/')
patients_ref_for_user = ref.child('patients')
patients_data = patients_ref_for_user.get()
existing_patient_data = patients_data if patients_data else {}
existing_pids = list(val['진료번호'] for val in existing_patient_data.values() if '진료번호' in val)

tab1, tab2, tab3 = st.tabs(["환자 관리", "환자 상태 확인 및 알림", "이메일 알림"])

with tab1:
    st.header("환자 등록/삭제")
    st.info("여기에 환자를 등록하거나 삭제할 수 있습니다.")

    with st.container(border=True):
        st.markdown("### 등록된 환자 목록")
        if existing_patient_data:
            for key, val in existing_patient_data.items():
                col1, col2 = st.columns([0.9, 0.1])
                with col1:
                    # '환자명'과 '진료번호' 키가 없을 경우 오류를 방지하기 위해 .get() 메서드를 사용
                    st.markdown(f"**{val.get('환자명', '미지정')}** / {val.get('진료번호', '미지정')} / {val.get('등록과', '미지정')}")
                with col2:
                    if st.button("X", key=f"delete_button_{key}"):
                        patients_ref_for_user.child(key).delete()
                        st.rerun()
        else:
            st.info("등록된 환자가 없습니다.")
    
    st.markdown("---")

    with st.form("register_form"):
        st.markdown("### 신규 환자 등록")
        name = st.text_input("환자명")
        pid = st.text_input("진료번호")
        selected_department = st.selectbox("등록 과", ["내과", "외과", "소아과", "미지정"])
        submitted = st.form_submit_button("등록")

        if submitted:
            if not name or not pid:
                st.warning("모든 항목을 입력해주세요.")
            elif any(v["진료번호"] == pid for v in existing_patient_data.values()):
                st.error("이미 등록된 진료번호입니다.")
            else:
                patients_ref_for_user.push().set({"환자명": name, "진료번호": pid, "등록과": selected_department})
                st.success(f"환자 '{name}'이(가) 등록되었습니다.")
                time.sleep(1)
                st.rerun()

with tab2:
    st.header("환자 상태 확인 및 알림")
    st.info("엑셀 파일에서 등록된 환자의 상태를 확인하고, 진료 일정을 캘린더에 추가할 수 있습니다.")

    uploaded_file = st.file_uploader("보호된 엑셀 파일(.xlsx)을 업로드하세요", type="xlsx")
    if uploaded_file:
        password = st.text_input("엑셀 파일 비밀번호를 입력하세요", type="password")
        if password:
            df = decrypt_and_read_excel(uploaded_file, password)
            if df is not None:
                st.write("### 업로드된 엑셀 파일 미리보기")
                st.dataframe(df.head())

                st.write("### 등록된 환자 진료 상태")
                pid_list = existing_pids
                if not pid_list:
                    st.warning("등록된 환자가 없습니다. '환자 관리' 탭에서 환자를 등록해주세요.")
                else:
                    found_patients_df = find_pid_in_dataframe(df, pid_list)
                    if not found_patients_df.empty:
                        st.dataframe(found_patients_df)
                    else:
                        st.info("업로드된 엑셀 파일에서 등록된 환자 정보를 찾을 수 없습니다.")
    
    st.markdown("---")
    
    st.markdown("### Google Calendar 연동")
    google_refresh_token = st.session_state.get("google_refresh_token", None)
    
    if st.button("Google Calendar 연동 시작"):
        service = get_google_calendar_service(refresh_token=None)
        if service:
            st.session_state["google_refresh_token"] = service._http.credentials.refresh_token
            st.success("Google Calendar에 성공적으로 연동되었습니다.")
            st.rerun()
    
    if google_refresh_token:
        st.success("Google Calendar와 연동되어 있습니다.")
        with st.form("calendar_event_form"):
            st.subheader("새로운 진료 이벤트 추가")
            event_summary = st.text_input("이벤트 제목", "환자 진료 일정")
            event_description = st.text_area("이벤트 설명", "진료 관련 내용")
            event_start_date = st.date_input("시작 날짜", datetime.date.today())
            event_start_time = st.time_input("시작 시간", datetime.time(9, 0))
            event_end_date = st.date_input("종료 날짜", datetime.date.today())
            event_end_time = st.time_input("종료 시간", datetime.time(10, 0))

            submitted_event = st.form_submit_button("캘린더에 추가")
            if submitted_event:
                try:
                    start_datetime = datetime.datetime.combine(event_start_date, event_start_time)
                    end_datetime = datetime.datetime.combine(event_end_date, event_end_time)
                    if start_datetime >= end_datetime:
                        st.error("종료 시간이 시작 시간보다 빨라야 합니다.")
                    else:
                        service = get_google_calendar_service(refresh_token=google_refresh_token)
                        if service:
                            create_event(service, start_datetime, end_datetime, event_summary, event_description)
                except Exception as e:
                    st.error(f"날짜/시간 입력 오류: {e}")

with tab3:
    st.header("이메일 알림")
    st.info("환자 상태에 대한 이메일 알림을 보낼 수 있습니다.")
    
    with st.form("email_form"):
        st.subheader("이메일 보내기")
        to_email = st.text_input("수신자 이메일 주소")
        subject = st.text_input("제목", "OCS 환자 알림")
        body = st.text_area("내용", "진료 상태 확인 부탁드립니다.")

        submitted_email = st.form_submit_button("이메일 전송")
        if submitted_email:
            if not is_valid_email(to_email):
                st.error("유효한 이메일 주소를 입력해주세요.")
            else:
                if send_email(to_email, subject, body):
                    st.success("이메일이 성공적으로 발송되었습니다.")
