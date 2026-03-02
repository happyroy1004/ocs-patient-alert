import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import pickle
import json

# local imports: config에서 SCOPES 상수를 가져옵니다.
from config import SCOPES

# --- 0. Secrets 로드 ---
try:
    FIREBASE_CREDENTIALS = dict(st.secrets["firebase"]) 
    DB_URL = st.secrets["database_url"] 
    google_calendar_secrets = st.secrets.get("google_calendar")
    GOOGLE_CALENDAR_CLIENT_SECRET = dict(google_calendar_secrets) if google_calendar_secrets else {}
except Exception as e:
    st.error(f"🚨 Secrets 로드 오류: {e}")
    FIREBASE_CREDENTIALS = None
    DB_URL = None
    GOOGLE_CALENDAR_CLIENT_SECRET = {}

# --- 1. DB 초기화 ---
@st.cache_resource
def get_db_refs():
    if not firebase_admin._apps:
        try:
            if FIREBASE_CREDENTIALS is None or DB_URL is None: return None, None, None
            creds_copy = FIREBASE_CREDENTIALS.copy()
            if 'FIREBASE_DATABASE_URL' in creds_copy: del creds_copy['FIREBASE_DATABASE_URL']
            cred = credentials.Certificate(creds_copy)
            firebase_admin.initialize_app(cred, {'databaseURL': DB_URL})
        except Exception as e:
            st.error(f"❌ Firebase 초기화 실패: {e}")
            return None, None, None 
    base_ref = db.reference()
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

# --- 2. 유틸리티 ---
def sanitize_path(email):
    return email.replace('.', '_')

def save_google_creds_to_firebase(safe_key, creds):
    try:
        creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
        encoded_creds = pickle.dumps(creds).hex()
        creds_ref.set({'creds': encoded_creds})
    except Exception as e:
        st.error(f"❌ 인증 저장 실패: {e}")

def load_google_creds_from_firebase(safe_key):
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    if data and 'creds' in data:
        return pickle.loads(bytes.fromhex(data['creds']))
    return None

# --- 3. Google Calendar Service (PKCE 에러 완벽 방어 버전) ---

def get_google_calendar_service(safe_key):
    # 이미 활성화된 서비스가 있다면 즉시 반환
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 설정 확인
    if not GOOGLE_CALENDAR_CLIENT_SECRET or "redirect_uri" not in GOOGLE_CALENDAR_CLIENT_SECRET:
        st.error("🚨 Secrets 설정(redirect_uri 등)이 누락되었습니다.")
        return None

    redirect_uri = GOOGLE_CALENDAR_CLIENT_SECRET["redirect_uri"]
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

    # 1. 🔑 인증 코드 처리 (구글에서 돌아온 직후)
    auth_code = st.query_params.get("code")
    if auth_code:
        # 인증 시작 시 세션에 저장해뒀던 Flow 객체가 있는지 확인
        if 'auth_flow' in st.session_state:
            try:
                flow = st.session_state.auth_flow
                flow.fetch_token(code=auth_code) # 여기서 code_verifier가 자동으로 사용됨
                creds = flow.credentials
                
                save_google_creds_to_firebase(safe_key, creds)
                st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
                
                # 인증 완료 후 깨끗하게 정리
                st.query_params.clear()
                del st.session_state.auth_flow
                st.rerun()
            except Exception as e:
                st.error(f"⚠️ 토큰 교환 오류 (Verifier 유실 가능성): {e}")
                del st.session_state.auth_flow # 실패 시 초기화
                st.query_params.clear()
        else:
            # 코드는 있는데 Flow 객체가 없다면 (세션 만료 등) 다시 시작해야 함
            st.warning("인증 세션이 만료되었습니다. 다시 시도해 주세요.")
            st.query_params.clear()

    # 2. DB에서 기존 Creds 로드 시도
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

    # 3. 🔑 신규 인증 링크 생성 (인증이 없는 경우)
    # Flow 객체를 새로 만들고 세션에 저장
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, 
            scopes=SCOPES, 
            redirect_uri=redirect_uri
        )

    auth_url, _ = st.session_state.auth_flow.authorization_url(
        prompt='consent', 
        access_type='offline'
    )
    
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 구글 계정 인증 링크]({auth_url})**")
    return None
