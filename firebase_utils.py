import streamlit as st
import firebase_admin
from firebase_admin import credentials, db, auth
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import os
import io
import pickle
import json

# --- [해결] Config 임포트 에러 방지 로직 ---
try:
    from config import SCOPES
except ImportError:
    # config.py가 없거나 SCOPES가 없을 경우를 대비한 기본값
    SCOPES = ['https://www.googleapis.com/auth/calendar']

# --- 0. Secrets 로드 및 초기 설정 ---
try:
    # Streamlit Cloud의 Secrets에서 정보 가져오기
    FIREBASE_CREDENTIALS = dict(st.secrets["firebase"]) 
    DB_URL = st.secrets["database_url"] 
    GOOGLE_CALENDAR_CLIENT_SECRET = dict(st.secrets["google_calendar"])
except Exception as e:
    st.error(f"🚨 Secrets.toml 로드 실패: {e}. 'secrets.toml' 설정을 확인해주세요.")
    FIREBASE_CREDENTIALS = None
    DB_URL = None
    GOOGLE_CALENDAR_CLIENT_SECRET = {}

# --- 1. DB 레퍼런스 및 초기화 ---
@st.cache_resource
def get_db_refs():
    if not firebase_admin._apps:
        try:
            if FIREBASE_CREDENTIALS and DB_URL:
                creds_init = FIREBASE_CREDENTIALS.copy()
                if 'FIREBASE_DATABASE_URL' in creds_init: 
                    del creds_init['FIREBASE_DATABASE_URL']
                
                cred = credentials.Certificate(creds_init)
                firebase_admin.initialize_app(cred, {'databaseURL': DB_URL})
        except Exception as e:
            st.error(f"❌ Firebase 앱 초기화 실패: {e}")
            return None, None, None 

    if firebase_admin._apps:
        base_ref = db.reference()
        users_ref = base_ref.child('users')
        doctor_users_ref = base_ref.child('doctor_users')
        
        def db_ref_func(path):
            return base_ref.child(path)
            
        return users_ref, doctor_users_ref, db_ref_func
    return None, None, None

# --- 2. Google Calendar Credentials 관리 ---
def save_google_creds_to_firebase(safe_key, creds):
    """Credentials 객체를 pickle로 직렬화하여 Firebase에 저장합니다."""
    creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
    encoded_creds = pickle.dumps(creds).hex()
    creds_ref.set({'creds': encoded_creds})

def load_google_creds_from_firebase(safe_key):
    """Firebase에서 Credentials를 로드합니다."""
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    if data and 'creds' in data:
        return pickle.loads(bytes.fromhex(data['creds']))
    return None

# --- 3. Google Calendar Service 로직 (Bad Request & PKCE 해결) ---

def get_google_calendar_service(safe_key):
    user_id_safe = safe_key
    
    # 1. 이미 서비스가 빌드되어 있다면 반환
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 2. Firebase에서 기존 토큰 로드 및 갱신 시도
    creds = load_google_creds_from_firebase(user_id_safe)
    if creds:
        if creds.valid:
            service = build('calendar', 'v3', credentials=creds)
            st.session_state.google_calendar_service = service
            return service
        elif creds.refresh_token:
            try:
                creds.refresh(Request())
                save_google_creds_to_firebase(user_id_safe, creds)
                service = build('calendar', 'v3', credentials=creds)
                st.session_state.google_calendar_service = service
                return service
            except:
                creds = None

    # 3. 신규 인증 설정
    redirect_uri = GOOGLE_CALENDAR_CLIENT_SECRET.get("redirect_uri")
    if not redirect_uri:
        st.error("🚨 redirect_uri가 secrets에 없습니다.")
        return None

    # 구글 표준 Config 구성
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

    # Flow 객체를 세션에 고정 (중요: 에러 방지의 핵심)
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, 
            scopes=SCOPES, 
            redirect_uri=redirect_uri
        )

    # 4. 토큰 교환 처리
    auth_code = st.query_params.get("code")

    if auth_code:
        try:
            # 세션에 저장된 flow를 사용해 토큰 획득
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            
            save_google_creds_to_firebase(user_id_safe, new_creds)
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
            
            # 정리 및 리셋
            st.query_params.clear()
            if 'auth_flow' in st.session_state:
                del st.session_state.auth_flow
            
            st.success("✅ 구글 인증 성공!")
            st.rerun()
            
        except Exception as e:
            st.error(f"⚠️ 인증 처리 중 오류 발생: {e}")
            st.query_params.clear()
            if 'auth_flow' in st.session_state:
                del st.session_state.auth_flow
            return None
    
    else:
        # 인증 링크 표시
        auth_url, _ = st.session_state.auth_flow.authorization_url(
            prompt='consent', 
            access_type='offline',
            include_granted_scopes='true'
        )
        st.info("구글 캘린더 연동이 필요합니다.")
        st.markdown(f"**[🔗 Google Calendar 인증 링크]({auth_url})**")
        return None

    return None

def sanitize_path(email):
    return email.replace('.', '_')

def recover_email(safe_key):
    db_ref = db.reference()
    paths = [f'users/{safe_key}', f'doctor_users/{safe_key}', safe_key]
    for path in paths:
        try:
            data = db_ref.child(path).get()
            if data and 'email' in data: return data['email']
        except: continue
    return None
