# firebase_utils.py

import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import pickle
import json

# config에서 순수한 상수(SCOPES)만 가져옵니다.
from config import SCOPES

# --- 0. 환경 설정 및 Secrets 로드 ---
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
    st.error(f"🚨 Secrets 로드 중 오류 발생: {e}")
    FIREBASE_CREDENTIALS = None
    DB_URL = None
    GOOGLE_CALENDAR_CLIENT_SECRET = {}

# --- 1. DB 레퍼런스 및 초기화 ---

@st.cache_resource
def get_db_refs():
    """Firebase Admin SDK를 초기화하고 DB 레퍼런스 객체를 반환합니다."""
    if not firebase_admin._apps:
        try:
            if FIREBASE_CREDENTIALS and DB_URL:
                cred = credentials.Certificate(FIREBASE_CREDENTIALS)
                firebase_admin.initialize_app(cred, {'databaseURL': DB_URL})
            else:
                return None, None, None
        except Exception as e:
            st.error(f"❌ Firebase 앱 초기화 실패: {e}")
            return None, None, None 

    base_ref = db.reference()
    users_ref = base_ref.child('users')
    doctor_users_ref = base_ref.child('doctor_users')
    
    def db_ref_func(path):
        return base_ref.child(path)
        
    return users_ref, doctor_users_ref, db_ref_func

# --- 2. Safe Key 및 인증 정보 관리 ---

def sanitize_path(email):
    """
    이메일 주소를 데이터베이스 전체에서 통일된 safe_key 형식으로 변환합니다.
    형식: user_at_domain_dot_com
    """
    if not email:
        return ""
    return email.lower().replace('@', '_at_').replace('.', '_dot_')

def save_google_creds_to_firebase(safe_key, creds):
    """
    Google 캘린더 OAuth2 Credentials 객체를 통합 노드에 저장합니다.
    경로: google_calendar_creds/{safe_key}
    """
    creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
    # 객체를 pickle로 직렬화하여 hex 문자열로 저장 (데이터 무결성 보장)
    pickled_creds = pickle.dumps(creds)
    encoded_creds = pickled_creds.hex()
    creds_ref.set({'creds': encoded_creds})

def load_google_creds_from_firebase(safe_key):
    """
    통합된 google_calendar_creds 노드에서 인증 정보를 로드합니다.
    """
    creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
    data = creds_ref.get()
    
    if data and 'creds' in data:
        try:
            encoded_creds = data['creds']
            pickled_creds = bytes.fromhex(encoded_creds)
            return pickle.loads(pickled_creds)
        except Exception as e:
            st.error(f"⚠️ 인증 데이터 복원 실패: {e}")
    return None

def check_google_connection_status(safe_key):
    """
    현재 사용자의 구글 연동 상태를 실시간으로 체크합니다.
    반환값: (bool: 연동여부, str: 상태메시지)
    """
    creds = load_google_creds_from_firebase(safe_key)
    
    if not creds:
        return False, "미연동 (인증 정보 없음)"
    
    try:
        if creds.valid:
            return True, "연동 정상"
        elif creds.expired and creds.refresh_token:
            # 만료된 경우 자동 갱신 시도
            creds.refresh(Request())
            save_google_creds_to_firebase(safe_key, creds)
            return True, "연동 정상 (자동 갱신됨)"
        else:
            return False, "인증 만료 (재인증 필요)"
    except Exception as e:
        return False, f"연동 오류 ({e})"

# --- 3. Google Calendar Service 구축 ---

def get_google_calendar_service(safe_key):
    """
    Google Calendar 서비스 객체를 생성하거나 인증 플로우를 시작합니다.
    """
    creds = load_google_creds_from_firebase(safe_key)

    # 1. 기존 인증 정보가 유효한 경우 바로 서비스 반환
    if creds and (creds.valid or (creds.expired and creds.refresh_token)):
        if creds.expired:
            creds.refresh(Request())
            save_google_creds_to_firebase(safe_key, creds)
        
        st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
        return

    # 2. 인증이 필요한 경우 OAuth Flow 시작
    if not GOOGLE_CALENDAR_CLIENT_SECRET:
        st.error("🚨 구글 캘린더 클라이언트 설정이 없습니다.")
        return

    redirect_uri = GOOGLE_CALENDAR_CLIENT_SECRET.get("redirect_uri")
    flow = InstalledAppFlow.from_client_config(
        {"installed": GOOGLE_CALENDAR_CLIENT_SECRET}, 
        SCOPES, 
        redirect_uri=redirect_uri 
    )
    
    auth_code = st.query_params.get("code")
    if auth_code:
        flow.fetch_token(code=auth_code)
        save_google_creds_to_firebase(safe_key, flow.credentials)
        st.success("✅ 구글 인증 성공!")
        st.query_params.clear() 
        st.rerun() 
    else:
        auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
        st.info("구글 계정 연동을 위해 아래 링크를 클릭해주세요.")
        st.markdown(f"**[🔗 구글 캘린더 연동하기]({auth_url})**")

def recover_email(safe_key):
    """통일된 safe_key를 사용하여 실제 이메일 주소를 복구합니다."""
    # 이 로직은 주로 알림 전송 시 수신자 이메일을 찾기 위해 사용됩니다.
    # safe_key 자체가 user_at_domain_dot_com 형태이므로 역변환하거나 DB를 조회합니다.
    for node in ['users', 'doctor_users']:
        data = db.reference(f'{node}/{safe_key}').get()
        if data and 'email' in data:
            return data['email']
    return None
