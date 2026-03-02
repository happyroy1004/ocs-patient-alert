import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
import json

# 권한 범위 설정
SCOPES = [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid'
]

# --- 1. Firebase 초기화 ---
@st.cache_resource
def get_db_refs():
    if not firebase_admin._apps:
        try:
            creds_dict = dict(st.secrets["firebase"])
            db_url = st.secrets["database_url"]
            if 'FIREBASE_DATABASE_URL' in creds_dict: 
                del creds_dict['FIREBASE_DATABASE_URL']
            cred = credentials.Certificate(creds_dict)
            firebase_admin.initialize_app(cred, {'databaseURL': db_url})
        except Exception as e:
            st.error(f"❌ Firebase 초기화 실패: {e}")
            return None, None, None
    base_ref = db.reference()
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

# --- 2. 유틸리티 함수 (ImportError 방지용) ---
def sanitize_path(email):
    return email.replace('.', '_')

def recover_email(safe_key):
    """sanitize된 키를 다시 이메일로 복구"""
    return safe_key.replace('_', '.')

# --- 3. Google 인증 관리 (PKCE 오류 해결 버전) ---
def save_google_creds_to_firebase(safe_key, creds):
    ref = db.reference(f'google_calendar_creds/{safe_key}')
    ref.set({'creds': pickle.dumps(creds).hex()})

def load_google_creds_from_firebase(safe_key):
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    if data and 'creds' in data:
        return pickle.loads(bytes.fromhex(data['creds']))
    return None

def get_google_calendar_service(safe_key):
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    conf = dict(st.secrets["google_calendar"])
    client_config = {
        "web": {
            "client_id": conf.get("client_id"),
            "project_id": conf.get("project_id"),
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "client_secret": conf.get("client_secret"),
            "redirect_uris": [conf.get("redirect_uri")]
        }
    }

    # Flow 객체 세션 유지 (중요)
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf.get("redirect_uri")
        )

    auth_code = st.query_params.get("code")
    if auth_code and 'auth_flow' in st.session_state:
        try:
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            save_google_creds_to_firebase(safe_key, new_creds)
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
            st.query_params.clear()
            del st.session_state.auth_flow
            st.rerun()
        except Exception as e:
            st.error(f"⚠️ 인증 오류: {e}")
            st.query_params.clear()

    creds = load_google_creds_from_firebase(safe_key)
    if creds:
        if creds.valid:
            service = build('calendar', 'v3', credentials=creds)
            st.session_state.google_calendar_service = service
            return service
        elif creds.refresh_token:
            creds.refresh(Request())
            save_google_creds_to_firebase(safe_key, creds)
            service = build('calendar', 'v3', credentials=creds)
            st.session_state.google_calendar_service = service
            return service

    auth_url, _ = st.session_state.auth_flow.authorization_url(prompt='consent', access_type='offline')
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 인증 링크]({auth_url})**")
    return None
