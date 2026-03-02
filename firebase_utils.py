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
    """Firebase에서 Credentials를 로드합니다."""
    creds_ref_new = db.reference(f'google_calendar_creds/{safe_key}')
    data_new = creds_ref_new.get()
    
    if data_new and 'creds' in data_new:
        return pickle.loads(bytes.fromhex(data_new['creds']))
    return None


# --- 3. Google Calendar Service 로직 (PKCE 오류 해결 버전) ---

def get_google_calendar_service(safe_key):
    """
    OAuth2 인증 흐름을 관리하고 Google Calendar API 서비스 객체를 반환합니다.
    Flow 객체를 세션에 보관하여 'Missing code verifier' 오류를 방지합니다.
    """
    user_id_safe = safe_key
    
    # 1. 기존 유효한 세션이 있는지 확인 (성능 최적화)
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 2. Firebase에서 기존 토큰 로드 시도
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
            except Exception:
                creds = None # 갱신 실패 시 재인증 유도

    # 4. 신규 인증 플로우 시작
    redirect_uri = GOOGLE_CALENDAR_CLIENT_SECRET.get("redirect_uri")
    if not redirect_uri:
        st.error("🚨 redirect_uri가 설정되지 않았습니다. Secrets.toml을 확인해주세요.")
        return None

    # Google Cloud Console 설정에 따른 유연한 대응
    client_config = {"web": GOOGLE_CALENDAR_CLIENT_SECRET} if "web" not in GOOGLE_CALENDAR_CLIENT_SECRET else GOOGLE_CALENDAR_CLIENT_SECRET

    # 핵심: Flow 객체를 세션에 보관하여 PKCE code_verifier 유실 방지
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, 
            scopes=SCOPES, 
            redirect_uri=redirect_uri
        )

    # URL에서 구글이 보내준 코드 확인
    auth_code = st.query_params.get("code")

    if auth_code:
        try:
            # 세션에 저장해둔 flow 객체를 사용하여 토큰 교환 (verifier가 유지됨)
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            
            # 성공 시 Firebase에 저장
            save_google_creds_to_firebase(user_id_safe, new_creds)
            
            # 서비스 객체 생성 및 세션 저장
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
            
            # 처리 완료 후 세션 데이터 및 URL 파라미터 정리
            st.query_params.clear()
            if 'auth_flow' in st.session_state:
                del st.session_state.auth_flow
            
            st.success("✅ Google Calendar 인증에 성공했습니다!")
            st.rerun() # URL을 깨끗하게 만들기 위해 재실행
            
        except Exception as e:
            st.error(f"인증 처리 중 오류 발생: {e}")
            st.query_params.clear()
            if 'auth_flow' in st.session_state:
                del st.session_state.auth_flow
            return None
    
    else:
        # 인증 코드가 없는 경우 사용자에게 인증 링크 제공
        auth_url, _ = st.session_state.auth_flow.authorization_url(
            prompt='consent', 
            access_type='offline',
            include_granted_scopes='true'
        )
        st.info("구글 캘린더 연동이 필요합니다.")
        st.markdown(f"**[🔗 Google Calendar 인증 링크]({auth_url})**")
        
        # 안내 문구
        with st.expander("인증 방법 안내"):
            st.write("1. 위 링크를 클릭하여 구글 로그인을 진행합니다.")
            st.write("2. '권한 허용' 화면에서 모든 항목을 선택한 후 승인합니다.")
            st.write("3. 자동으로 이 페이지로 돌아오며 연동이 완료됩니다.")
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
