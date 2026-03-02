import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
import json
import os

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
    return safe_key.replace('_', '.')

# --- 3. Google 인증 관리 (Missing code verifier 대응) ---
def save_google_creds_to_firebase(safe_key, creds):
    ref = db.reference(f'google_calendar_creds/{safe_key}')
    ref.set({'creds': pickle.dumps(creds).hex()})

def load_google_creds_from_firebase(safe_key):
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    if data and 'creds' in data:
        return pickle.loads(bytes.fromhex(data['creds']))
    return None

def get_google_calendar_service(safe_key):
    # 1. 이미 인증된 서비스가 있는 경우
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

    # 2. 기존 저장된 토큰 로드
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

    # 3. Flow 객체 세션 초기화 (없으면 생성)
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf.get("redirect_uri")
        )

    # 4. URL 파라미터 처리 (인증 코드 복구)
    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            # 세션에 flow가 있을 때만 진행
            if 'auth_flow' in st.session_state:
                st.session_state.auth_flow.fetch_token(code=auth_code)
                new_creds = st.session_state.auth_flow.credentials
                save_google_creds_to_firebase(safe_key, new_creds)
                st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
                st.query_params.clear()
                del st.session_state.auth_flow
                st.rerun()
        except Exception as e:
            # 에러 발생 시 flow를 삭제하여 새로 만들 수 있게 함
            st.warning("⚠️ 인증 세션이 만료되었습니다. 아래 링크를 다시 눌러주세요.")
            st.query_params.clear()
            if 'auth_flow' in st.session_state:
                del st.session_state.auth_flow
            # 여기서 멈추지 않고 아래 링크 생성 코드로 넘어가게 함

    # 5. 인증 링크 다시 생성 (Flow 객체 유무 확인 후 생성)
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf.get("redirect_uri")
        )

    auth_url, _ = st.session_state.auth_flow.authorization_url(
        prompt='consent', 
        access_type='offline',
        include_granted_scopes='true'
    )
    
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 구글 계정 연동하기]({auth_url})**")
    return None
