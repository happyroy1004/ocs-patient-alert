import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
import json
import os

# 권한 범위 설정 (config.py의 SCOPES와 일치 확인)
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

# --- 3. Google 인증 관리 (PKCE 오류 완전 방어) ---
def save_google_creds_to_firebase(safe_key, creds):
    ref = db.reference(f'google_calendar_creds/{safe_key}')
    ref.set({'creds': pickle.dumps(creds).hex()})

def load_google_creds_from_firebase(safe_key):
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    if data and 'creds' in data:
        try:
            return pickle.loads(bytes.fromhex(data['creds']))
        except: return None
    return None

def get_google_calendar_service(safe_key):
    # 1. 세션에 서비스가 이미 로드되어 있다면 즉시 반환
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 2. Firebase에서 기존 토큰 로드 시도
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
            except: pass # 갱신 실패 시 재인증 진행

    # 3. OAuth Flow 설정
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

    # 🔑 Flow 객체를 세션에 고정 (가장 중요)
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf.get("redirect_uri")
        )

    # 4. URL의 인증 코드(code) 처리
    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            # 세션에 저장된 flow를 사용하여 토큰 교환 (verifier 유지)
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            
            # Firebase에 영구 저장
            save_google_creds_to_firebase(safe_key, new_creds)
            
            # 서비스 빌드 및 세션 저장
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
            
            # 성공 후 정리 및 리런
            st.query_params.clear()
            if 'auth_flow' in st.session_state: del st.session_state.auth_flow
            st.rerun()
        except Exception as e:
            # 실패 시 flow 초기화하여 다시 시작할 수 있게 함
            if 'auth_flow' in st.session_state: del st.session_state.auth_flow
            st.query_params.clear()
            # 아래에서 새로운 auth_url을 생성하도록 흐름을 넘김

    # 5. 인증되지 않은 경우 링크 표시 (새로운 Flow 객체 필요 시 생성)
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf.get("redirect_uri")
        )

    auth_url, _ = st.session_state.auth_flow.authorization_url(
        prompt='consent', access_type='offline', include_granted_scopes='true'
    )
    
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 구글 계정 연동하기]({auth_url})**")
    return None
