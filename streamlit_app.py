import streamlit as st
import os
import datetime

from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# --- 환경 설정 ---
# Streamlit secrets에 저장된 클라이언트 정보와 리디렉션 URI를 불러옵니다.
# secrets.toml 파일에 아래 내용이 있는지 확인해 주세요.
# [google_calendar]
# client_id = "YOUR_CLIENT_ID.apps.googleusercontent.com"
# client_secret = "YOUR_CLIENT_SECRET"
# redirect_uri = "https://ocs-patient-alert-etaaycuhlm7xfuzqqmub9a.streamlit.app"
try:
    client_id = st.secrets["google_calendar"]["client_id"]
    client_secret = st.secrets["google_calendar"]["client_secret"]
    redirect_uri = st.secrets["google_calendar"]["redirect_uri"]
except KeyError:
    st.error("`secrets.toml` 파일에 Google Calendar 설정이 누락되었습니다. 파일을 확인해 주세요.")
    st.stop()

# 사용 권한 범위 (Google Calendar events)
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

# --- 함수: 인증 및 토큰 관리 ---
def get_google_calendar_service():
    """
    세션에 저장된 인증 정보를 사용하여 Google Calendar 서비스 객체를 반환합니다.
    """
    creds = st.session_state.get("credentials")
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                st.error(f"토큰 갱신 실패: {e}")
                st.session_state.credentials = None
                return None
        else:
            return None # 인증 정보가 없으면 None 반환
    
    # 인증 정보가 유효하면 서비스 객체 생성
    service = build('calendar', 'v3', credentials=creds)
    return service

def get_authorization_url():
    """
    Google 인증 URL을 생성하여 반환합니다.
    """
    flow = InstalledAppFlow.from_client_config(
        {
            "installed": {
                "client_id": client_id,
                "client_secret": client_secret,
                "redirect_uris": [redirect_uri],
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
            }
        },
        SCOPES
    )
    auth_url, _ = flow.authorization_url(prompt='consent')
    return auth_url

def add_event_to_google_calendar(service, summary, start_time, end_time):
    """
    Google 캘린더에 일정을 추가하는 함수
    """
    event = {
        'summary': summary,
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
        st.success(f"캘린더에 일정을 추가했습니다: {event.get('htmlLink')}")
    except Exception as e:
        st.error(f"일정 추가 실패: {e}")

# --- 앱 실행 로직 ---
st.title("환자 관리 및 캘린더 연동")

# URL에서 인증 코드 확인 후 토큰 교환
query_params = st.query_params
auth_code = query_params.get("code")

if auth_code:
    flow = InstalledAppFlow.from_client_config(
        {
            "installed": {
                "client_id": client_id,
                "client_secret": client_secret,
                "redirect_uris": [redirect_uri],
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
            }
        },
        SCOPES
    )
    
    try:
        flow.fetch_token(code=auth_code[0])
        creds = flow.credentials
        st.session_state.credentials = creds
        st.query_params.clear()  # URL에서 코드 제거
        st.success("Google 캘린더 권한이 성공적으로 허용되었습니다!")
    except Exception as e:
        st.error(f"인증 실패: {e}")

# 관리자 모드 비밀번호 입력 (좌측 사이드바)
password = st.sidebar.text_input("관리자 비밀번호", type="password")
if password == "admin":  # 비밀번호는 "admin"으로 가정
    st.session_state.admin_mode = True
    st.sidebar.success("관리자 모드 활성화")
else:
    st.session_state.admin_mode = False
    st.sidebar.error("비밀번호를 입력하세요.")

# 관리자 모드에서만 보이는 기능
if st.session_state.admin_mode:
    st.subheader("매칭된 환자에게 메일 보내기")

    if st.button("캘린더에 일정 추가하기"):
        service = get_google_calendar_service()
        
        if service:
            st.info("권한이 확인되었습니다. 캘린더에 일정을 추가합니다.")
            # 여기에 환자 정보를 이용해 캘린더에 추가할 로직을 작성하세요.
            event_summary = "환자 진료 예약 (홍길동)"
            event_start = datetime.datetime.now(datetime.timezone.utc)
            event_end = event_start + datetime.timedelta(hours=1)
            add_event_to_google_calendar(service, event_summary, event_start, event_end)
        else:
            st.warning("캘린더 권한이 없습니다. 아래 링크를 눌러 권한을 허용해 주세요.")
            auth_url = get_authorization_url()
            st.markdown(f"[**클릭하여 권한 허용하기**]({auth_url})", unsafe_allow_html=True)
            st.caption("링크를 복사하여 권한이 필요한 사람에게 전달할 수 있습니다.")
