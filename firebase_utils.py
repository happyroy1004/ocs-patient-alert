import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import pickle
import json

# local imports: config에서 상수를 가져옵니다.
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
    
except Exception as e:
    st.error(f"🚨 Secrets 로드 오류: {e}")
    FIREBASE_CREDENTIALS = None
    DB_URL = None
    GOOGLE_CALENDAR_CLIENT_SECRET = {}

# --- 1. DB 레퍼런스 및 초기화 ---

@st.cache_resource
def get_db_refs():
    """Firebase Admin SDK 초기화 및 레퍼런스 반환"""
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
            st.error(f"❌ Firebase 초기화 실패: {e}")
            return None, None, None 

    base_ref = db.reference()
    users_ref = base_ref.child('users')
    doctor_users_ref = base_ref.child('doctor_users')
    
    def db_ref_func(path):
        return base_ref.child(path)
        
    return users_ref, doctor_users_ref, db_ref_func

# --- 2. 유틸리티 함수 ---

def sanitize_path(email):
    """이메일 주소를 DB 키 형식으로 변환"""
    return email.replace('.', '_')

def recover_email(safe_key):
    """DB에서 실제 이메일 주소를 찾아 반환"""
    db_ref = db.reference()
    paths = [f'users/{safe_key}', f'doctor_users/{safe_key}', safe_key]
    for path in paths:
        data = db_ref.child(path).get()
        if data and isinstance(data, dict) and 'email' in data:
            return data['email']
    return None

# --- 3. Google Credentials 관리 ---

def save_google_creds_to_firebase(safe_key, creds):
    """인증 정보를 Pickle/Hex 형식으로 Firebase에 저장"""
    try:
        creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
        encoded_creds = pickle.dumps(creds).hex()
        creds_ref.set({'creds': encoded_creds})
    except Exception as e:
        st.error(f"❌ 인증 정보 저장 실패: {e}")

def load_google_creds_from_firebase(safe_key):
    """Firebase에서 인증 정보 로드 및 마이그레이션 지원"""
    # 1. 신규 형식 로드
    data_new = db.reference(f'google_calendar_creds/{safe_key}').get()
    if data_new and 'creds' in data_new:
        return pickle.loads(bytes.fromhex(data_new['creds']))

    # 2. 구 형식(Plaintext) 로드 및 자동 마이그레이션
    db_ref = db.reference()
    for path in [f'{safe_key}/google_creds', f'users/{safe_key}/google_creds', f'doctor_users/{safe_key}/google_creds']:
        data_old = db_ref.child(path).get()
        if data_old and data_old.get('refresh_token'):
            try:
                creds = Credentials(
                    token=data_old.get('token'),
                    refresh_token=data_old.get('refresh_token'),
                    token_uri=data_old.get('token_uri') or 'https://oauth2.googleapis.com/token',
                    client_id=data_old.get('client_id'),
                    client_secret=data_old.get('client_secret'),
                    scopes=data_old.get('scopes') or SCOPES
                )
                save_google_creds_to_firebase(safe_key, creds)
                return creds
            except: continue
    return None

# --- 4. Google Calendar Service 핵심 로직 (PKCE 수정 완료) ---

def get_google_calendar_service(safe_key):
    """구글 캘린더 서비스 빌드 및 OAuth2 인증 흐름 제어"""
    
    # 세션에 이미 서비스가 활성화되어 있다면 즉시 반환
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 1. 설정 로드
    if not GOOGLE_CALENDAR_CLIENT_SECRET or "redirect_uri" not in GOOGLE_CALENDAR_CLIENT_SECRET:
        st.error("🚨 Secrets.toml의 [google_calendar] 설정이나 redirect_uri가 누락되었습니다.")
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

    # 2. 🔑 중요: Flow 객체를 세션에 고정 (PKCE 'Missing code verifier' 방지 핵심)
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, 
            scopes=SCOPES, 
            redirect_uri=redirect_uri
        )

    # 3. URL 파라미터(코드) 처리: 인증 완료 후 돌아온 경우
    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            # 세션에 저장해뒀던 바로 그 flow 객체로 토큰 교환
            st.session_state.auth_flow.fetch_token(code=auth_code)
            creds = st.session_state.auth_flow.credentials
            
            save_google_creds_to_firebase(safe_key, creds)
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
            
            # 정리 및 리로드
            st.query_params.clear()
            if 'auth_flow' in st.session_state:
                del st.session_state.auth_flow
            st.rerun()
        except Exception as e:
            st.error(f"⚠️ 인증 실패: {e}")
            st.query_params.clear()
            if 'auth_flow' in st.session_state:
                del st.session_state.auth_flow
            return None

    # 4. DB에서 기존 Creds 로드 시도
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
            except Exception as e:
                st.warning("로그인이 만료되어 재인증이 필요합니다.")

    # 5. 인증 URL 생성 및 사용자 안내
    auth_url, _ = st.session_state.auth_flow.authorization_url(
        prompt='consent', 
        access_type='offline',
        include_granted_scopes='true'
    )
    
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 구글 계정 인증 링크]({auth_url})**")
    
    with st.expander("🔑 연동 방법 안내"):
        st.write("""
        1. 위 **인증 링크**를 클릭합니다.
        2. 구글 로그인 후 '권한 허용'을 모두 체크하고 확인을 누릅니다.
        3. 앱이 자동으로 재시작되며 연동이 완료됩니다.
        """)
    
    return None
