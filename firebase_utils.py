import streamlit as st
import firebase_admin
from firebase_admin import credentials, db, auth
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import pickle

# SCOPES 설정
SCOPES = ['https://www.googleapis.com/auth/calendar', 'https://www.googleapis.com/auth/userinfo.email', 'openid']

# --- 0. Secrets 로드 ---
try:
    FIREBASE_CREDENTIALS = dict(st.secrets["firebase"]) 
    DB_URL = st.secrets["database_url"] 
    GOOGLE_CALENDAR_CLIENT_SECRET = dict(st.secrets["google_calendar"])
except Exception as e:
    st.error("🚨 Secrets.toml 설정을 확인해주세요.")
    GOOGLE_CALENDAR_CLIENT_SECRET = {}

# --- 1. DB 초기화 ---
@st.cache_resource
def get_db_refs():
    if not firebase_admin._apps:
        creds_init = FIREBASE_CREDENTIALS.copy()
        if 'FIREBASE_DATABASE_URL' in creds_init: del creds_init['FIREBASE_DATABASE_URL']
        cred = credentials.Certificate(creds_init)
        firebase_admin.initialize_app(cred, {'databaseURL': DB_URL})
    base_ref = db.reference()
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

# --- 2. Credentials 관리 (계정별 독립 저장) ---
def save_google_creds_to_firebase(google_email_key, creds):
    """구글 인증 계정 이메일을 키로 사용하여 저장 (OCS 계정과 분리)"""
    db.reference(f'google_calendar_creds/{google_email_key}').set({
        'creds': pickle.dumps(creds).hex(),
        'email': google_email_key
    })

def load_google_creds_from_firebase(google_email_key):
    data = db.reference(f'google_calendar_creds/{google_email_key}').get()
    return pickle.loads(bytes.fromhex(data['creds'])) if data else None

# --- 3. Google Calendar Service (계정 매핑 로직 포함) ---

def get_google_calendar_service(safe_key):
    """
    OCS 로그인 계정(safe_key)과 구글 계정이 달라도 작동하도록 설계되었습니다.
    """
    # 1. 기존 세션 확인
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 2. 이 OCS 계정에 연결된 구글 계정이 있는지 확인
    # (연결 정보를 따로 저장하는 로직이 없다면, 일단 로그인된 계정 키로 로드 시도)
    creds = load_google_creds_from_firebase(safe_key)

    # 3. Flow 설정
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

    # 4. [핵심] 토큰 교환 및 실제 구글 계정 확인
    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            
            # 구글 API를 통해 실제 로그인한 구글 계정 이메일 가져오기
            user_info_service = build('oauth2', 'v2', credentials=new_creds)
            user_info = user_info_service.userinfo().get().execute()
            google_email = user_info.get('email')
            google_safe_key = google_email.replace('.', '_')
            
            # OCS 계정(safe_key)과 구글 계정(google_safe_key) 모두에 권한 저장 (매핑)
            save_google_creds_to_firebase(safe_key, new_creds) 
            save_google_creds_to_firebase(google_safe_key, new_creds)
            
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
            
            st.query_params.clear()
            if 'auth_flow' in st.session_state: del st.session_state.auth_flow
            
            st.success(f"✅ 인증 성공! ({google_email} 계정 연동됨)")
            st.rerun()
        except Exception as e:
            st.error(f"인증 오류: {e}")
            st.query_params.clear()
            if 'auth_flow' in st.session_state: del st.session_state.auth_flow
            return None

    # 5. 기존 토큰 갱신 로직
    if creds:
        if creds.valid:
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
            return st.session_state.google_calendar_service
        elif creds.refresh_token:
            try:
                creds.refresh(Request())
                save_google_creds_to_firebase(safe_key, creds)
                st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
                return st.session_state.google_calendar_service
            except: pass

    # 6. 인증 링크 표시
    auth_url, _ = st.session_state.auth_flow.authorization_url(
        prompt='consent', access_type='offline', include_granted_scopes='true'
    )
    st.info("구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 Google Calendar 인증 링크]({auth_url})**")
    return None

def sanitize_path(email):
    return email.replace('.', '_')
