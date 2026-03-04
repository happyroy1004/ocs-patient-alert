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

# local imports
from config import SCOPES

# 1. 환경 설정 로드
try:
    FIREBASE_CREDENTIALS = dict(st.secrets["firebase"]) 
    DB_URL = st.secrets["database_url"] 
    google_calendar_secrets = st.secrets.get("google_calendar")
    if google_calendar_secrets:
        GOOGLE_CALENDAR_CLIENT_SECRET = dict(google_calendar_secrets)
    else:
        st.error("🚨 Secrets.toml에 [google_calendar] 섹션 누락")
        GOOGLE_CALENDAR_CLIENT_SECRET = {}
except Exception as e:
    st.error(f"🚨 설정 로드 오류: {e}")
    FIREBASE_CREDENTIALS = None
    DB_URL = None
    GOOGLE_CALENDAR_CLIENT_SECRET = {}

# --- DB 초기화 ---
@st.cache_resource
def get_db_refs():
    if not firebase_admin._apps:
        try:
            if FIREBASE_CREDENTIALS and DB_URL:
                creds_init = FIREBASE_CREDENTIALS.copy()
                if 'FIREBASE_DATABASE_URL' in creds_init: del creds_init['FIREBASE_DATABASE_URL']
                cred = credentials.Certificate(creds_init)
                firebase_admin.initialize_app(cred, {'databaseURL': DB_URL})
        except Exception as e:
            st.error(f"❌ Firebase 초기화 실패: {e}")
            return None, None, None 

    if firebase_admin._apps:
        base_ref = db.reference()
        return base_ref.child('users'), base_ref.child('doctor_users'), lambda p: base_ref.child(p)
    return None, None, None

# --- Creds 관리 ---
def sanitize_path(email):
    return email.replace('.', '_')

def save_google_creds_to_firebase(safe_key, creds):
    """중요: 갱신된 정보를 Firebase에 확실히 저장"""
    creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
    pickled_creds = pickle.dumps(creds)
    creds_ref.set({'creds': pickled_creds.hex()})

def load_google_creds_from_firebase(safe_key):
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    if data and 'creds' in data:
        return pickle.loads(bytes.fromhex(data['creds']))
    return None

# --- 서비스 로드 및 인증 흐름 ---
def get_google_calendar_service(safe_key):
    user_id_safe = safe_key
    
    # 1. 기존 Creds 로드
    creds = load_google_creds_from_firebase(user_id_safe)

    # 2. Config 설정
    if not GOOGLE_CALENDAR_CLIENT_SECRET:
        return None

    client_config = {
        "web": {
            "client_id": GOOGLE_CALENDAR_CLIENT_SECRET.get("client_id"),
            "client_secret": GOOGLE_CALENDAR_CLIENT_SECRET.get("client_secret"),
            "auth_uri": GOOGLE_CALENDAR_CLIENT_SECRET.get("auth_uri"),
            "token_uri": GOOGLE_CALENDAR_CLIENT_SECRET.get("token_uri"),
            "redirect_uris": [GOOGLE_CALENDAR_CLIENT_SECRET.get("redirect_uri")]
        }
    }

    # 3. 유효성 검사 및 자동 갱신
    if creds and creds.valid:
        st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
        return
        
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            save_google_creds_to_firebase(user_id_safe, creds)
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
            return
        except:
            creds = None 

    # 4. OAuth Flow 시작
    redirect_uri = GOOGLE_CALENDAR_CLIENT_SECRET.get("redirect_uri")
    flow = Flow.from_client_config(client_config, scopes=SCOPES, redirect_uri=redirect_uri)
    
    auth_code = st.query_params.get("code")
    
    if auth_code:
        try:
            # 💡 [해결 핵심] Missing code verifier 방지를 위해 수동 토큰 교환
            flow.fetch_token(code=auth_code)
            new_creds = flow.credentials
            save_google_creds_to_firebase(user_id_safe, new_creds)
            st.success("✅ 인증 정보가 Firebase에 성공적으로 저장되었습니다!")
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"❌ 토큰 교환 실패: {e}")
            if "verifier" in str(e):
                st.info("인증 세션이 만료되었습니다. 아래 링크를 다시 눌러주세요.")
                st.query_params.clear()
    else:
        # prompt='consent'와 access_type='offline'은 Refresh Token 발급을 위해 필수
        auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
        st.warning("⚠️ 구글 캘린더 연동이 필요합니다.")
        st.markdown(f"**[Google Calendar 인증 링크]({auth_url})**")
        return None

def recover_email(safe_key):
    base = db.reference()
    for p in [f'users/{safe_key}', f'doctor_users/{safe_key}', safe_key]:
        data = base.child(p).get()
        if data and 'email' in data: return data['email']
    return None
