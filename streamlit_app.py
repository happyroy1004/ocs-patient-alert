import streamlit as st
import os
import pickle
import datetime

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# --- 환경 설정 ---
# Google Cloud Console에서 발급받은 클라이언트 정보
# Streamlit secrets에 저장된 정보를 불러옵니다.
client_id = st.secrets["google_calendar"]["client_id"]
client_secret = st.secrets["google_calendar"]["client_secret"]

# 사용 권한 범위 (Google Calendar events)
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

# 리디렉션 URI는 Streamlit 앱 URL과 동일해야 합니다.
redirect_uri = st.secrets["google_calendar"]["redirect_uri"]

# --- 상태 관리 ---
# 세션 상태 초기화 (관리자 모드, 인증 정보 등)
if "admin_mode" not in st.session_state:
    st.session_state.admin_mode = False
if "credentials" not in st.session_state:
    st.session_state.credentials = None

# --- 함수: 인증 및 토큰 관리 ---
def get_google_calendar_service():
    """
    인증 정보를 사용하여 Google Calendar 서비스 객체를 반환합니다.
    권한이 없으면 인증 페이지로 리디렉션하는 링크를 생성합니다.
    """
    creds = st.session_state.credentials
    
    # 인증 정보가 없거나 유효하지 않은 경우
    if not creds or not creds.valid:
        # 토큰 갱신
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # 인증 flow 시작
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
            
            # 인증 URL 생성
            auth_url, _ = flow.authorization_url(prompt='consent')
            
            st.error("캘린더 권한이 없습니다. 아래 링크를 눌러 권한을 허용해 주세요.")
            st.markdown(f"[**클릭하여 권한 허용하기**]({auth_url})", unsafe_allow_html=True)
            return None
    
    # 인증 정보가 유효하면 서비스 객체 생성
    st.session_state.credentials = creds
    service = build('calendar', 'v3', credentials=creds)
    return service

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

# 관리자 모드 비밀번호 입력
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

    # 버튼 클릭 시 권한 확인 및 일정 추가 로직 실행
    if st.button("캘린더 추가하기"):
        service = get_google_calendar_service()
        
        if service:
            # 권한이 있는 경우, 여기에 캘린더에 추가할 로직을 작성합니다.
            # 예시:
            event_summary = "환자 진료 예약"
            event_start = datetime.datetime.utcnow()
            event_end = event_start + datetime.timedelta(hours=1)
            add_event_to_google_calendar(service, event_summary, event_start, event_end)

# URL에서 인증 코드 확인 후 토큰 교환
query_params = st.experimental_get_query_params()
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
        st.session_state.credentials = flow.credentials
        st.experimental_set_query_params()  # URL에서 코드 제거
        st.success("Google 캘린더 권한이 성공적으로 허용되었습니다!")
    except Exception as e:
        st.error(f"인증 실패: {e}")
