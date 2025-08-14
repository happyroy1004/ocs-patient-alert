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
import calendar

# --- Google Calendar API 관련 import 및 설정 ---
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
        # Streamlit Secrets에서 Firebase 서비스 계정 정보를 가져옵니다.
        firebase_credentials_json_str = st.secrets["firebase"]["FIREBASE_SERVICE_ACCOUNT_JSON"]
        firebase_credentials_dict = json.loads(firebase_credentials_json_str)

        cred = credentials.Certificate(firebase_credentials_dict)
        firebase_admin.initialize_app(cred, {
            'databaseURL': st.secrets["firebase"]["database_url"]
        })
    except (KeyError, FileNotFoundError) as e:
        st.error(f"Firebase 초기화 중 오류가 발생했습니다: {e}")
        st.info("secrets.toml 파일에 Firebase 설정이 올바르게 되어 있는지 확인해주세요.")

# Firebase DB 참조
patients_ref_for_user = db.reference('/patients/user1')
sheet_keyword_ref = db.reference('/sheet_keyword')
department_keyword_map_ref = db.reference('/department_keyword_map')
calendar_settings_ref = db.reference('/calendar_settings')
patient_calendar_ref = db.reference('/patient_calendars')

# 전역 변수 초기화
sheet_keyword_to_department_map = None

# --- Google Calendar API 함수 ---
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

def get_google_calendar_credentials():
    creds = st.session_state.get('google_creds', None)
    
    # 만료된 인증 정보가 있으면 갱신
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        st.session_state['google_creds'] = creds
        st.info("Google Calendar 인증 정보가 갱신되었습니다.")
        return build('calendar', 'v3', credentials=creds)

    # 인증 정보가 없으면 인증 흐름 시작
    if not creds:
        # URL 쿼리 파라미터에서 'code'를 확인
        query_params = st.experimental_get_query_params()
        if "code" in query_params:
            code = query_params["code"][0]
            # 세션에 저장된 flow 객체를 사용하여 토큰을 가져옴
            if 'flow' in st.session_state:
                flow = st.session_state['flow']
                try:
                    flow.fetch_token(code=code)
                    st.session_state['google_creds'] = flow.credentials
                    st.experimental_set_query_params(code=[]) # URL에서 코드를 제거
                    st.info("Google Calendar 인증이 완료되었습니다.")
                except Exception as e:
                    st.error(f"인증 토큰을 가져오는 중 오류가 발생했습니다: {e}")
                    # 실패 시 flow 객체 삭제
                    if 'flow' in st.session_state:
                        del st.session_state['flow']
            else:
                st.error("인증 흐름을 복원할 수 없습니다. 다시 시도해주세요.")
            return None
        
        # flow 객체가 없으면 새로 생성
        if 'flow' not in st.session_state:
            try:
                google_calendar_secrets = st.secrets["googlecalendar"]
                client_config = {
                    "web": {
                        "client_id": google_calendar_secrets.get("client_id"),
                        "client_secret": google_calendar_secrets.get("client_secret"),
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token",
                        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                        "redirect_uris": [google_calendar_secrets.get("redirect_uri")]
                    }
                }
                flow = InstalledAppFlow.from_client_config(client_config, scopes=SCOPES)
                st.session_state['flow'] = flow

                authorization_url, _ = flow.authorization_url(prompt='consent')
                st.session_state['authorization_url'] = authorization_url
            except KeyError as e:
                st.error(f"Google Calendar API 설정 오류: secrets.toml 파일에 '[googlecalendar]' 섹션이 없거나 형식이 잘못되었습니다. {e}")
                return None
        
        if 'authorization_url' in st.session_state:
            st.warning("Google 계정 로그인 필요! 아래 링크를 클릭하여 로그인해주세요.")
            st.markdown(f"[{st.session_state['authorization_url']}]({st.session_state['authorization_url']})")
        
        return None
    
    # 인증 정보가 있으면 서비스 빌드
    try:
        service = build('calendar', 'v3', credentials=creds)
        st.session_state['calendar_service'] = service
        return service
    except Exception as e:
        st.error(f"Google Calendar 서비스 빌드 중 오류: {e}")
        return None

def create_google_calendar_event(service, calendar_id, event_data):
    try:
        event = service.events().insert(calendarId=calendar_id, body=event_data).execute()
        return event
    except Exception as e:
        st.error(f"Google Calendar 이벤트 생성 중 오류: {e}")
        return None

# --- Streamlit UI 구성 ---

st.set_page_config(
    page_title="진료 스케줄 등록 시스템",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🏥 진료 스케줄 등록 시스템")

# Google Calendar API 서비스 로드
calendar_service = get_google_calendar_credentials()

# 사이드바
with st.sidebar:
    st.header("설정 및 기능")
    st.markdown("---")

    # 환자 관리
    st.subheader("환자 관리")
    existing_patient_data = patients_ref_for_user.get()
    if existing_patient_data:
        st.info("등록된 환자 목록:")
        for key, val in existing_patient_data.items():
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
    with st.form("register_form"):
        st.subheader("환자 등록")
        name = st.text_input("환자명")
        pid = st.text_input("진료번호")
        
        # 진료과 리스트 로드
        department_data = department_keyword_map_ref.get()
        if department_data:
            departments_for_registration = sorted(list(set(department_data.values())))
        else:
            departments_for_registration = ["내과", "외과", "소아과", "기타"]

        selected_department = st.selectbox("등록 과", departments_for_registration)
        
        submitted = st.form_submit_button("등록")
        if submitted:
            if not name or not pid:
                st.warning("모든 항목을 입력해주세요.")
            elif existing_patient_data and any(
                v["환자명"] == name and v["진료번호"] == pid and v.get("등록과") == selected_department
                for v in existing_patient_data.values()):
                st.error("이미 등록된 환자입니다.")
            else:
                patients_ref_for_user.push({
                    "환자명": name,
                    "진료번호": pid,
                    "등록과": selected_department
                })
                st.success("환자가 성공적으로 등록되었습니다!")
                st.rerun()
    st.markdown("---")

    # 캘린더 설정
    st.subheader("Google Calendar 설정")
    google_creds_exist = 'google_creds' in st.session_state and st.session_state['google_creds'] is not None
    if google_creds_exist:
        st.success("Google Calendar에 연결되었습니다.")
        calendar_list = calendar_service.calendarList().list().execute().get('items', [])
        calendar_names = {c['summary']: c['id'] for c in calendar_list}
        
        with st.form("calendar_form"):
            selected_calendar_name = st.selectbox("일정을 추가할 캘린더", sorted(calendar_names.keys()))
            submitted_calendar = st.form_submit_button("캘린더 설정 저장")
            if submitted_calendar:
                calendar_id = calendar_names[selected_calendar_name]
                calendar_settings_ref.set({"calendarId": calendar_id, "calendarName": selected_calendar_name})
                st.success(f"'{selected_calendar_name}' 캘린더가 기본으로 설정되었습니다.")
                st.rerun()
    else:
        st.warning("Google Calendar에 연결되지 않았습니다.")
        if st.button("Google Calendar 로그인"):
            if 'authorization_url' in st.session_state:
                st.markdown(f"[{st.session_state['authorization_url']}]({st.session_state['authorization_url']})")
            else:
                st.warning("로그인 URL을 생성할 수 없습니다. 페이지를 새로고침 해주세요.")
    
# --- 메인 페이지 ---

st.header("엑셀 파일 업로드 및 스케줄 등록")
uploaded_file = st.file_uploader("암호화된 엑셀 파일(.xlsx) 업로드", type="xlsx")

if uploaded_file:
    # 엑셀 파일 암호 해제
    password = st.text_input("엑셀 파일 비밀번호", type="password")
    if password:
        try:
            decrypted_file = io.BytesIO()
            office_file = msoffcrypto.OfficeFile(uploaded_file)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted_file)

            # 엑셀 파일 읽기
            df = pd.read_excel(decrypted_file)
            st.write("엑셀 파일 미리보기:")
            st.dataframe(df)

            # DB에서 키워드-진료과 매핑 정보 가져오기
            sheet_keyword_to_department_map = department_keyword_map_ref.get()

            # 시트 키워드-컬럼 매핑 정보
            sheet_keyword_data = sheet_keyword_ref.get()
            
            # 메일 발송 기능
            st.markdown("---")
            st.subheader("이메일 발송")
            with st.form("email_form"):
                sender_email = st.text_input("보내는 사람 이메일")
                receiver_email = st.text_input("받는 사람 이메일")
                email_password = st.text_input("보내는 사람 이메일 비밀번호", type="password")

                email_submitted = st.form_submit_button("메일 발송")

                if email_submitted:
                    if not is_valid_email(sender_email) or not is_valid_email(receiver_email):
                        st.error("이메일 주소가 올바르지 않습니다.")
                    else:
                        try:
                            # 엑셀 시트 생성
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                df.to_excel(writer, index=False, sheet_name='Sheet1')
                            output.seek(0)

                            # 이메일 전송
                            msg = MIMEMultipart()
                            msg['From'] = sender_email
                            msg['To'] = receiver_email
                            msg['Subject'] = f"{df.iloc[0]['진료번호']} 환자의 진료 스케줄"
                            body = "안녕하세요. 환자 진료 스케줄 파일입니다."
                            msg.attach(MIMEText(body, 'plain'))
                            
                            part = MIMEText(output.getvalue(), _subtype="xlsx")
                            part.add_header('Content-Disposition', 'attachment', filename="진료스케줄.xlsx")
                            msg.attach(part)

                            server = smtplib.SMTP('smtp.gmail.com', 587)
                            server.starttls()
                            server.login(sender_email, email_password)
                            server.sendmail(sender_email, receiver_email, msg.as_string())
                            server.quit()
                            st.success("이메일이 성공적으로 발송되었습니다!")
                        except Exception as e:
                            st.error(f"이메일 발송 중 오류가 발생했습니다: {e}")
            
            # Google Calendar에 등록된 환자 스케줄 등록
            st.markdown("---")
            st.subheader("Google Calendar 일정 등록")

            if calendar_service and 'google_creds' in st.session_state:
                st.success("Google Calendar에 연결되었습니다.")
                calendar_list = calendar_service.calendarList().list().execute().get('items', [])
                calendar_names = {c['summary']: c['id'] for c in calendar_list}
                
                with st.form("calendar_form_main"):
                    selected_calendar_name = st.selectbox("일정을 추가할 캘린더", sorted(calendar_names.keys()), key="main_calendar_select")
                    submitted_calendar = st.form_submit_button("캘린더 설정 저장", key="main_calendar_submit")
                    if submitted_calendar:
                        calendar_id = calendar_names[selected_calendar_name]
                        calendar_settings_ref.set({"calendarId": calendar_id, "calendarName": selected_calendar_name})
                        st.success(f"'{selected_calendar_name}' 캘린더가 기본으로 설정되었습니다.")
                        st.rerun()
            else:
                st.warning("Google Calendar에 연결되지 않았습니다.")
                if st.button("Google Calendar 로그인", key="login_btn_main"):
                    if 'authorization_url' in st.session_state:
                        st.markdown(f"[{st.session_state['authorization_url']}]({st.session_state['authorization_url']})")
                    else:
                        st.warning("로그인 URL을 생성할 수 없습니다. 페이지를 새로고침 해주세요.")
        except Exception as e:
            st.error(f"엑셀 파일 처리 중 오류가 발생했습니다: {e}")
            st.warning("비밀번호가 올바른지, 또는 파일이 손상되지 않았는지 확인해주세요.")

