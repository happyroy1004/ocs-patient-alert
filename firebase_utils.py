# firebase_utils.py

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

# local imports: config에서 순수한 상수(SCOPES)만 가져옵니다.
from config import SCOPES

# --- 0. Secrets 로드 및 초기 설정 ---
try:
    FIREBASE_CREDENTIALS = dict(st.secrets["firebase"]) 
    DB_URL = st.secrets["database_url"] 

    google_calendar_secrets = st.secrets.get("google_calendar")
    if google_calendar_secrets:
        GOOGLE_CALENDAR_CLIENT_SECRET = dict(google_calendar_secrets)
    else:
        st.error("🚨 Secrets.toml에 [google_calendar] 섹션이 누락되었습니다.")
        GOOGLE_CALENDAR_CLIENT_SECRET = {}
    
except KeyError as e:
    st.error(f"🚨 중요: Secrets.toml 설정 오류. '{e.args[0]}' 키를 찾을 수 없습니다.")
    FIREBASE_CREDENTIALS = None
    DB_URL = None
    GOOGLE_CALENDAR_CLIENT_SECRET = {}
except Exception as e:
    st.error(f"🚨 Secrets 로드 중 예상치 못한 오류 발생: {e}")
    FIREBASE_CREDENTIALS = None
    DB_URL = None
    GOOGLE_CALENDAR_CLIENT_SECRET = {}


# --- 1. DB 레퍼런스 및 초기화 ---

@st.cache_resource
def get_db_refs():
    if not firebase_admin._apps:
        try:
            if FIREBASE_CREDENTIALS is None or DB_URL is None:
                return None, None, None

            creds_for_init = FIREBASE_CREDENTIALS.copy()
            if 'FIREBASE_DATABASE_URL' in creds_for_init: 
                 del creds_for_init['FIREBASE_DATABASE_URL']
            
            cred = credentials.Certificate(creds_for_init)
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

def sanitize_path(email):
    return email.replace('.', '_')

def save_google_creds_to_firebase(safe_key, creds):
    """Credentials 객체를 pickle로 직렬화하여 Firebase에 저장합니다."""
    creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
    pickled_creds = pickle.dumps(creds)
    encoded_creds = pickled_creds.hex()
    creds_ref.set({'creds': encoded_creds})

def load_google_creds_from_firebase(safe_key):
    """Firebase에서 Credentials를 로드하고 필요한 경우 마이그레이션합니다."""
    creds_ref_new = db.reference(f'google_calendar_creds/{safe_key}')
    data_new = creds_ref_new.get()
    
    if data_new and 'creds' in data_new:
        return pickle.loads(bytes.fromhex(data_new['creds']))

    # 마이그레이션 레이어 (기존 형식 대응)
    db_ref = db.reference()
    paths = [f'{safe_key}/google_creds', f'users/{safe_key}/google_creds', f'doctor_users/{safe_key}/google_creds']
    for path in paths:
        data_old = db_ref.child(path).get()
        if data_old and data_old.get('refresh_token'):
            try:
                creds = Credentials(
                    token=data_old.get('token'),
                    refresh_token=data_old.get('refresh_token'),
                    token_uri=data_old.get('token_uri') or 'https://oauth2.googleapis.com/token',
                    client_id=data_old.get('client_id'),
                    client_secret=data_old.get('client_secret'),
                    scopes=data_old.get('scopes') if isinstance(data_old.get('scopes'), list) else SCOPES
                )
                save_google_creds_to_firebase(safe_key, creds)
                return creds
            except: continue
    return None


# --- 3. Google Calendar Service 로직 (핵심 수정 부분) ---

def get_google_calendar_service(safe_key):
    """
    OAuth2 인증 흐름을 관리하고 Google Calendar API 서비스 객체를 반환합니다.
    InvalidGrantError 방지를 위해 세션 상태와 URL 파라미터를 엄격히 관리합니다.
    """
    user_id_safe = safe_key
    
    # 1. 기존 유효한 세션이 있는지 확인
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 2. Firebase에서 Credentials 로드
    creds = load_google_creds_from_firebase(user_id_safe)

    # 3. 토큰 유효성 검사 및 갱신
    if creds:
        if creds.valid:
            service = build('calendar', 'v3', credentials=creds)
            st.session_state.google_calendar_service = service
            return service
        
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                save_google_creds_to_firebase(user_id_safe, creds)
                service = build('calendar', 'v3', credentials=creds)
                st.session_state.google_calendar_service = service
                return service
            except Exception as e:
                st.warning(f"토큰 갱신 실패: {e}. 다시 인증해 주세요.")
                creds = None

    # 4. 신규 인증 플로우 (Redirection)
    redirect_uri = GOOGLE_CALENDAR_CLIENT_SECRET.get("redirect_uri")
    if not redirect_uri:
        st.error("🚨 redirect_uri가 설정되지 않았습니다.")
        return None

    # Flow 객체 생성 (installed 구조 대응)
    client_config = {"web": GOOGLE_CALENDAR_CLIENT_SECRET} if "web" not in GOOGLE_CALENDAR_CLIENT_SECRET else GOOGLE_CALENDAR_CLIENT_SECRET
    if "installed" not in client_config and "web" not in client_config:
        client_config = {"installed": GOOGLE_CALENDAR_CLIENT_SECRET}

    flow = Flow.from_client_config(client_config, scopes=SCOPES, redirect_uri=redirect_uri)

    # URL 파라미터 확인
    auth_code = st.query_params.get("code")

    if auth_code and not st.session_state.get('oauth_in_progress'):
        try:
            # 중복 실행 방지 플래그 설정
            st.session_state.oauth_in_progress = True
            
            # 토큰 교환 (여기서 InvalidGrantError 발생 가능성 차단)
            flow.fetch_token(code=auth_code)
            new_creds = flow.credentials
            
            save_google_creds_to_firebase(user_id_safe, new_creds)
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
            
            st.success("✅ 인증 성공!")
            
            # 파라미터 정리 및 리런 (매우 중요)
            st.query_params.clear()
            del st.session_state.oauth_in_progress
            st.rerun()
            
        except Exception as e:
            st.error(f"인증 오류: {e}")
            st.query_params.clear()
            if 'oauth_in_progress' in st.session_state:
                del st.session_state.oauth_in_progress
            return None
    
    elif not auth_code:
        # 인증이 필요할 때 링크 표시
        auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
        st.info("구글 캘린더 연동이 필요합니다.")
        st.markdown(f"**[🔗 Google Calendar 인증 링크]({auth_url})**")
        return None

    return None

def recover_email(safe_key):
    db_ref = db.reference()
    paths = [f'users/{safe_key}', f'doctor_users/{safe_key}', safe_key]
    for path in paths:
        try:
            data = db_ref.child(path).get()
            if data and 'email' in data: return data['email']
        except: continue
    return None
