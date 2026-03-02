import streamlit as st
import firebase_admin
from firebase_admin import credentials, db, auth
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import pickle

# SCOPES 설정
SCOPES = ['https://www.googleapis.com/auth/calendar']

# --- 0. Secrets 로드 ---
try:
    FIREBASE_CREDENTIALS = dict(st.secrets["firebase"]) 
    DB_URL = st.secrets["database_url"] 
    GOOGLE_CALENDAR_CLIENT_SECRET = dict(st.secrets["google_calendar"])
except Exception as e:
    st.error("🚨 Secrets.toml 설정을 확인해주세요.")
    GOOGLE_CALENDAR_CLIENT_SECRET = {}

# --- 1. DB 초기화 (생략/기존동일) ---
@st.cache_resource
def get_db_refs():
    if not firebase_admin._apps:
        creds_init = FIREBASE_CREDENTIALS.copy()
        if 'FIREBASE_DATABASE_URL' in creds_init: del creds_init['FIREBASE_DATABASE_URL']
        cred = credentials.Certificate(creds_init)
        firebase_admin.initialize_app(cred, {'databaseURL': DB_URL})
    base_ref = db.reference()
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

# --- 2. Credentials 관리 (기존동일) ---
def save_google_creds_to_firebase(safe_key, creds):
    db.reference(f'google_calendar_creds/{safe_key}').set({'creds': pickle.dumps(creds).hex()})

def load_google_creds_from_firebase(safe_key):
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    return pickle.loads(bytes.fromhex(data['creds'])) if data else None

# --- 3. Google Calendar Service (최종 수정 버전) ---

def get_google_calendar_service(safe_key):
    """
    (invalid_grant) Bad Request를 해결하기 위해 
    URL 파라미터 처리 로직을 최우선으로 배치합니다.
    """
    # 1. URL에 인증 코드가 있는지 먼저 확인 (가장 중요)
    auth_code = st.query_params.get("code")
    
    # 2. 이미 서비스가 빌드되어 있다면 반환
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 3. Flow 설정 준비
    redirect_uri = GOOGLE_CALENDAR_CLIENT_SECRET.get("redirect_uri")
    client_config = {
        "web": {
            "client_id": GOOGLE_CALENDAR_CLIENT_SECRET.get("client_id"),
            "project_id": GOOGLE_CALENDAR_CLIENT_SECRET.get("project_id"),
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "client_secret": GOOGLE_CALENDAR_CLIENT_SECRET.get("client_secret"),
            "redirect_uris": [redirect_uri]
        }
    }

    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=redirect_uri
        )

    # 4. [핵심] 코드가 있다면 즉시 토큰으로 교환 (함수 중간에 있으면 중복 실행 위험)
    if auth_code:
        try:
            # fetch_token 실행 전 코드를 변수에 담고 파라미터에서 미리 제거 시도
            # (Streamlit의 rerun 시 중복 호출 방지)
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            
            save_google_creds_to_firebase(safe_key, new_creds)
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
            
            st.query_params.clear() # URL 깨끗하게 청소
            if 'auth_flow' in st.session_state: del st.session_state.auth_flow
            
            st.success("✅ 인증 완료!")
            st.rerun()
        except Exception as e:
            # 여기서 Bad Request가 뜬다면 코드가 이미 만료되었거나 URI 불일치
            st.warning("인증 코드가 만료되었습니다. 다시 시도해주세요.")
            st.query_params.clear()
            if 'auth_flow' in st.session_state: del st.session_state.auth_flow
            st.rerun()

    # 5. 기존 토큰 로드 시도 (코드가 없을 때만 실행)
    creds = load_google_creds_from_firebase(safe_key)
    if creds:
        if creds.valid:
            service = build('calendar', 'v3', credentials=creds)
            st.session_state.google_calendar_service = service
            return service
        elif creds.refresh_token:
            try:
                creds.refresh(Request())
                save_google_creds_to_firebase(safe_key, creds)
                service = build('calendar', 'v3', credentials=creds)
                st.session_state.google_calendar_service = service
                return service
            except: pass

    # 6. 인증 링크 표시 (아무것도 없을 때)
    auth_url, _ = st.session_state.auth_flow.authorization_url(
        prompt='consent', access_type='offline', include_granted_scopes='true'
    )
    st.info("구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 Google Calendar 인증 링크]({auth_url})**")
    return None

def sanitize_path(email):
    return email.replace('.', '_')

def recover_email(safe_key):
    db_ref = db.reference()
    for path in [f'users/{safe_key}', f'doctor_users/{safe_key}', safe_key]:
        data = db_ref.child(path).get()
        if data and 'email' in data: return data['email']
    return None
