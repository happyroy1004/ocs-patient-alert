# firebase_utils.py

import streamlit as st # 💡 st.secrets 및 캐싱을 위해 필요
import firebase_admin
from firebase_admin import credentials, db, auth
from google_auth_oauthlib.flow import InstalledAppFlow, Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import os
import io
import pickle
import json

# local imports: config에서 순수한 상수(SCOPES)만 가져옵니다.
from config import SCOPES

# 💡 st.secrets를 사용하여 인증 정보를 로드하고 전역 변수로 설정
try:
    # 1. Firebase Admin SDK 인증 정보 로드
    FIREBASE_CREDENTIALS = dict(st.secrets["firebase"]) 
    
    # 2. DB URL 로드
    DB_URL = st.secrets["database_url"] 

    # 3. Google Calendar Client Secret 로드: 섹션이 없는 경우를 대비해 안전하게 처리
    google_calendar_secrets = st.secrets.get("google_calendar")
    if google_calendar_secrets:
        # 딕셔너리로 변환하여 사용
        GOOGLE_CALENDAR_CLIENT_SECRET = dict(google_calendar_secrets)
    else:
        st.error("🚨 Secrets.toml에 [google_calendar] 섹션이 누락되었습니다. Google Calendar 기능을 사용할 수 없습니다.")
        GOOGLE_CALENDAR_CLIENT_SECRET = {} # 빈 딕셔너리로 초기화
    
except KeyError as e:
    # 'firebase'나 'database_url' 같은 필수 키가 누락되었을 때 앱이 멈추지 않도록 처리
    st.error(f"🚨 중요: Secrets.toml 설정 오류. '{e.args[0]}' 키를 찾을 수 없습니다. secrets.toml 파일의 키 이름과 위치를 확인해 주세요.")
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
    """
    Firebase Admin SDK를 초기화하고 DB 레퍼런스 객체를 반환합니다.
    """
    users_ref = None
    doctor_users_ref = None
    
    if not firebase_admin._apps:
        try:
            if FIREBASE_CREDENTIALS is None or DB_URL is None:
                st.warning("DB 연결 정보가 불완전하여 초기화를 건너뜠습니다.")
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


# --- 2. Google Calendar 인증 및 Creds 관리 ---

def sanitize_path(email):
    """
    이메일 주소를 Firebase Realtime Database 키로 사용할 수 있도록 정리합니다.
    """
    safe_email = email.replace('.', '_')
    return safe_email


def save_google_creds_to_firebase(safe_key, creds):
    """Google 캘린더 OAuth2 Credentials 객체를 Firebase의 새 형식에 맞게 저장합니다 (pickle 직렬화)."""
    creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
    
    pickled_creds = pickle.dumps(creds)
    encoded_creds = pickled_creds.hex()
    
    creds_ref.set({'creds': encoded_creds})


def load_google_creds_from_firebase(safe_key):
    """Firebase에서 Google Calendar OAuth2 Credentials 객체를 로드합니다."""
    
    # 1. 새롭고 올바른 경로 (Pickle/Hex 형식)에서 로드 시도
    creds_ref_new = db.reference(f'google_calendar_creds/{safe_key}')
    data_new = creds_ref_new.get()
    
    if data_new and 'creds' in data_new:
        encoded_creds = data_new['creds']
        pickled_creds = bytes.fromhex(encoded_creds)
        creds = pickle.loads(pickled_creds)
        return creds

    # 2. 🚨 기존 경로 (Plaintext 형식)에서 로드 시도 (마이그레이션 레이어)
    
    def get_old_creds_data(safe_key):
        db_ref = db.reference()
        
        paths_to_check = [
            f'{safe_key}/google_creds', 
            f'google_creds/{safe_key}'
            f'users/{safe_key}/google_creds', 
            f'doctor_users/{safe_key}/google_creds'
        ]
        
        for path in paths_to_check:
            data = db_ref.child(path).get()
            if data: return data
        return None

    data_old = get_old_creds_data(safe_key)
    
    if data_old and data_old.get('refresh_token'):
        st.warning("🚨 기존 Google Credentials를 감지했습니다. 새 형식으로 마이그레이션을 시도합니다.")
        try:
            scopes_data = data_old.get('scopes')
            if isinstance(scopes_data, dict):
                 scopes_list = list(scopes_data.values())
            elif isinstance(scopes_data, list):
                 scopes_list = scopes_data
            else:
                 scopes_list = SCOPES

            creds = Credentials(
                token=data_old.get('token'),
                refresh_token=data_old.get('refresh_token'),
                token_uri=data_old.get('token_uri') or 'https://oauth2.googleapis.com/token',
                client_id=data_old.get('client_id'),
                client_secret=data_old.get('client_secret'),
                scopes=scopes_list
            )
            save_google_creds_to_firebase(safe_key, creds)
            st.success("✅ 기존 인증 정보를 성공적으로 로드하고 마이그레이션했습니다.")
            return creds

        except Exception as e:
            st.error(f"❌ 기존 Credentials 마이그레이션 실패: 다시 인증을 시도해 주세요. ({e})")
            return None 

    return None


# --- 3. Google Calendar Service 로드/인증 흐름 ---

def get_google_calendar_service(safe_key):
    """
    Google Calendar 서비스 객체를 로드하거나, 인증이 필요하면 리다이렉트 흐름을 시작합니다.
    """
    user_id_safe = safe_key
    st.session_state.google_calendar_service = None
    
    # 1. Credentials 로드 (새 형식 -> 구 형식 순으로 시도)
    creds = load_google_creds_from_firebase(user_id_safe)

    # 2. Secrets에서 client_config 준비 (OAuth 라이브러리 형식에 맞게)
    google_secrets_flat = GOOGLE_CALENDAR_CLIENT_SECRET 
    if not isinstance(google_secrets_flat, dict) or not google_secrets_flat:
        st.info("구글 캘린더 Secrets 설정이 불완전합니다. Secrets.toml을 확인해주세요.")
        return

    # OAuth 라이브러리가 기대하는 'installed' 구조로 감싸기
    client_config = {
    "web": {
        "client_id": google_secrets_flat.get("client_id"),
        "client_secret": google_secrets_flat.get("client_secret"),
        "auth_uri": google_secrets_flat.get("auth_uri"),
        "token_uri": google_secrets_flat.get("token_uri"),
        "redirect_uris": [google_secrets_flat.get("redirect_uri")] # 리스트 형태여야 할 수도 있음
    }
}

    # 3. Credentials 유효성 검사 및 갱신 시도
    if creds and creds.valid:
        st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
        return
        
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            save_google_creds_to_firebase(user_id_safe, creds)
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
            return
        except Exception as e:
            st.warning(f"Refresh Token 갱신 실패: {e}. 재인증이 필요합니다.")
            creds = None # 갱신 실패 시 폴백

    # 4. 🚨 인증 플로우 시작 (리다이렉트 로직)
    
    # redirect_uri 유효성 검사 및 추출
    redirect_uri = google_secrets_flat.get("redirect_uri")
    if not redirect_uri:
        st.error("🚨 Google Calendar Secrets에 'redirect_uri'가 정의되어 있지 않습니다. secrets.toml을 확인해주세요.")
        return

    # 인증 플로우 생성 (InstalledAppFlow 사용)
    flow = InstalledAppFlow.from_client_config(
        client_config, 
        SCOPES, 
        redirect_uri=redirect_uri 
    )
    
    if not creds:
        auth_code = st.query_params.get("code")
        
        if auth_code:
            # 인증 코드를 사용하여 토큰을 교환
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            
            save_google_creds_to_firebase(user_id_safe, creds)
            
            st.success("Google Calendar 인증이 완료되었습니다.")
            
            # 리디렉션으로 인한 쿼리 파라미터 정리 및 앱 리로드
            st.query_params.clear() 
            st.rerun() 
            
        else:
            # 인증 URL 생성 및 사용자에게 표시
            auth_url, _ = flow.authorization_url(prompt='consent')
            st.warning("구글 캘린더 연동을 위해 인증이 필요합니다. 아래 링크를 클릭하여 권한을 부여하세요.")
            st.markdown(f"**[Google Calendar 인증 링크]({auth_url})**")
            
            # 신규 사용자에게 연동 방법을 명확히 안내
            st.info("""
            ### 🔑 구글 캘린더 연동 방법
            1. [Google Calendar 인증 링크]를 클릭하여 Google 로그인 및 권한 부여 페이지로 이동합니다.
            2. Warning이 나타나도 무시하고, Advanced를 누른 후, 밑줄로 나타나는 사이트를 클릭합니다.
            3. 그 후, 권한을 승인하면, 다시 이 사이트로 자동으로 돌아옵니다.
            4. 다시 로그인했을 때, 권한이 허용되어 있는 것을 확인할 수 있습니다. 
            """)
            return None

    if creds:
         st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
         return
    
    return None


def recover_email(safe_key):
    """Firebase의 user 노드에서 safe_key에 해당하는 실제 이메일을 찾습니다."""
    db_ref = db.reference()
    
    paths_to_check = [f'users/{safe_key}', f'doctor_users/{safe_key}', safe_key]
    
    for path in paths_to_check:
        try:
            data = db_ref.child(path).get()
            if data and 'email' in data:
                return data['email']
        except Exception:
            continue
            
    return None
