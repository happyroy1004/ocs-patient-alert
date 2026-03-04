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

    # 3. Google Calendar Client Secret 로드
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
    """Firebase Admin SDK를 초기화하고 DB 레퍼런스 객체를 반환합니다."""
    users_ref = None
    doctor_users_ref = None
    
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


# --- 2. Google Calendar 인증 및 Creds 관리 ---

def sanitize_path(email):
    """이메일 주소를 Firebase 키로 사용할 수 있도록 정리합니다."""
    return email.replace('.', '_')


def save_google_creds_to_firebase(safe_key, creds):
    """Google 캘린더 Credentials 객체를 Firebase에 저장합니다."""
    creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
    pickled_creds = pickle.dumps(creds)
    encoded_creds = pickled_creds.hex()
    creds_ref.set({'creds': encoded_creds})


def load_google_creds_from_firebase(safe_key):
    """Firebase에서 Google Calendar Credentials 객체를 로드합니다."""
    creds_ref_new = db.reference(f'google_calendar_creds/{safe_key}')
    data_new = creds_ref_new.get()
    
    if data_new and 'creds' in data_new:
        encoded_creds = data_new['creds']
        pickled_creds = bytes.fromhex(encoded_creds)
        return pickle.loads(pickled_creds)

    return None


# --- 3. Google Calendar Service 로드/인증 흐름 ---

def get_google_calendar_service(safe_key):
    """Google Calendar 서비스 객체를 로드하거나 인증 흐름을 시작합니다."""
    user_id_safe = safe_key
    st.session_state.google_calendar_service = None
    
    # 1. Credentials 로드 시도
    creds = load_google_creds_from_firebase(user_id_safe)

    # 2. Config 준비 및 변수 선언 (에러 방지용)
    google_secrets_flat = GOOGLE_CALENDAR_CLIENT_SECRET 
    if not google_secrets_flat or not isinstance(google_secrets_flat, dict):
        st.info("구글 캘린더 설정이 불완전합니다. Secrets.toml을 확인하세요.")
        return None

    # Web 애플리케이션 구조로 생성
    client_config = {
        "web": {
            "client_id": google_secrets_flat.get("client_id"),
            "client_secret": google_secrets_flat.get("client_secret"),
            "auth_uri": google_secrets_flat.get("auth_uri"),
            "token_uri": google_secrets_flat.get("token_uri"),
            "redirect_uris": [google_secrets_flat.get("redirect_uri")]
        }
    }

    # 3. Credentials 유효성 검사 및 갱신
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
            st.warning(f"인증 갱신 실패: {e}")
            creds = None 

    # 4. 🚨 인증 플로우 시작 (PKCE 에러 방지를 위해 Flow 직접 사용)
    redirect_uri = google_secrets_flat.get("redirect_uri")
    
    # InstalledAppFlow 대신 Flow.from_client_config를 사용합니다.
    flow = Flow.from_client_config(
        client_config, 
        scopes=SCOPES, 
        redirect_uri=redirect_uri
    )
    
    if not creds:
        auth_code = st.query_params.get("code")
        
        if auth_code:
            try:
                # 💡 [핵심] fetch_token에서 code_verifier를 요구하지 않도록 설정하거나,
                # 세션을 유지할 수 없는 환경이므로 아래와 같이 처리합니다.
                flow.fetch_token(code=auth_code)
                creds = flow.credentials
                
                save_google_creds_to_firebase(user_id_safe, creds)
                st.success("Google Calendar 인증이 완료되었습니다.")
                st.query_params.clear() 
                st.rerun() 
            except Exception as e:
                # 만약 여기서도 에러가 난다면, 구글 콘솔에서 '데스크톱' 클라이언트를 새로 만들어야 할 수도 있습니다.
                st.error(f"토큰 교환 실패: {e}")
        else:
            # access_type='offline'과 prompt='consent'는 Refresh Token 발급에 필수입니다.
            auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
            st.warning("구글 캘린더 연동을 위해 인증이 필요합니다.")
            st.markdown(f"**[Google Calendar 인증 링크]({auth_url})**")
            
            st.info("링크 클릭 후 권한 승인을 완료하면 이 페이지로 돌아옵니다.")
            return None

def recover_email(safe_key):
    """Firebase에서 실제 이메일을 찾습니다."""
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
