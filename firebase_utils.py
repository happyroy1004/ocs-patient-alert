import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
import json

# --- [설정] 권한 범위 ---
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
            if "firebase" not in st.secrets or "database_url" not in st.secrets:
                st.error("🚨 Secrets.toml 설정 확인 필요")
                return None, None, None
                
            creds_dict = dict(st.secrets["firebase"])
            db_url = st.secrets["database_url"]
            
            cred = credentials.Certificate(creds_dict)
            firebase_admin.initialize_app(cred, {'databaseURL': db_url})
        except Exception as e:
            st.error(f"❌ Firebase 초기화 실패: {e}")
            return None, None, None

    base_ref = db.reference()
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

# --- 2. 유틸리티 함수 ---
def sanitize_path(email):
    return email.replace('.', '_')

def recover_email(safe_key):
    """sanitize된 키를 다시 이메일로 복구 (필요 시 사용)"""
    return safe_key.replace('_', '.')

# --- 3. Google Credentials 관리 ---
def save_google_creds_to_firebase(safe_key, creds):
    try:
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        encoded_creds = pickle.dumps(creds).hex()
        ref.set({'creds': encoded_creds})
        st.success("✅ 인증 정보가 저장되었습니다.")
    except Exception as e:
        st.error(f"❌ DB 저장 실패: {e}")

def load_google_creds_from_firebase(safe_key):
    try:
        data = db.reference(f'google_calendar_creds/{safe_key}').get()
        if data and 'creds' in data:
            return pickle.loads(bytes.fromhex(data['creds']))
    except:
        return None
    return None

def get_google_calendar_service(safe_key):
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    if "google_calendar" not in st.secrets:
        st.error("🚨 Secrets 설정 누락")
        return None
        
    conf = dict(st.secrets["google_calendar"])
    redirect_uri = conf.get("redirect_uri")
    
    client_config = {
        "web": {
            "client_id": conf.get("client_id"),
            "project_id": conf.get("project_id"),
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "client_secret": conf.get("client_secret"),
            "redirect_uris": [redirect_uri]
        }
    }

    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(client_config, scopes=SCOPES, redirect_uri=redirect_uri)

    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            save_google_creds_to_firebase(safe_key, new_creds)
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
            st.query_params.clear()
            if 'auth_flow' in st.session_state: del st.session_state.auth_flow
            st.rerun()
        except Exception as e:
            st.error(f"⚠️ 인증 오류: {str(e)}")
            st.query_params.clear()
            return None

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
            except:
                creds = None

    auth_url, _ = st.session_state.auth_flow.authorization_url(prompt='consent', access_type='offline')
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 구글 계정 인증 링크]({auth_url})**")
    return None
