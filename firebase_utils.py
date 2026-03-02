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

# config에서 SCOPES 가져오기
try:
    from config import SCOPES
except ImportError:
    SCOPES = ['https://www.googleapis.com/auth/calendar']

# --- 0. Secrets 로드 및 초기 설정 ---
try:
    FIREBASE_CREDENTIALS = dict(st.secrets["firebase"]) 
    DB_URL = st.secrets["database_url"] 
    GOOGLE_CALENDAR_CLIENT_SECRET = dict(st.secrets["google_calendar"])
except Exception as e:
    st.error(f"🚨 Secrets.toml 로드 실패: {e}")
    FIREBASE_CREDENTIALS = None
    DB_URL = None
    GOOGLE_CALENDAR_CLIENT_SECRET = {}

# --- 1. DB 초기화 ---
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
        return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)
    return None, None, None

# --- 2. Credentials 관리 ---
def save_google_creds_to_firebase(safe_key, creds):
    creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
    encoded_creds = pickle.dumps(creds).hex()
    creds_ref.set({'creds': encoded_creds})

def load_google_creds_from_firebase(safe_key):
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    if data and 'creds' in data:
        return pickle.loads(bytes.fromhex(data['creds']))
    return None

# --- 3. Google Calendar Service (Bad Request 해결 버전) ---

def get_google_calendar_service(safe_key):
    """
    세션 상태를 엄격히 관리하여 invalid_grant(Bad Request/Missing verifier)를 방지합니다.
    """
    user_id_safe = safe_key
    
    # 1. 기존 세션 확인
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 2. Firebase 로드 및 갱신
    creds = load_google_creds_from_firebase(user_id_safe)
    if creds:
        if creds.valid:
            service = build('calendar', 'v3', credentials=creds)
            st.session_state.google_calendar_service = service
            return service
        elif creds.refresh_token:
            try:
                creds.refresh(Request())
                save_google_creds_to_firebase(user_id_safe, creds)
                service = build('calendar', 'v3', credentials=creds)
                st.session_state.google_calendar_service = service
                return service
            except:
                creds = None

    # 3. 신규 인증 설정
    redirect_uri = GOOGLE_CALENDAR_CLIENT_SECRET.get("redirect_uri")
    
    # [중요] 구글 서버가 요구하는 표준 규격으로 Config 재구성
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

    # Flow 객체를 세션에 고정 (PKCE verifier 유지 핵심)
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, 
            scopes=SCOPES, 
            redirect_uri=redirect_uri
        )

    # 4. 토큰 교환 처리
    auth_code = st.query_params.get("code")

    if auth_code:
        try:
            # 세션에 저장된 flow 객체로 토큰 교환
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            
            save_google_creds_to_firebase(user_id_safe, new_creds)
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
            
            # 정리 및 리런
            st.query_params.clear()
            if 'auth_flow' in st.session_state:
                del st.session_state.auth_flow
            
            st.success("✅ 구글 인증 성공!")
            st.rerun()
            
        except Exception as e:
            st.error(f"⚠️ 인증 처리 중 오류 발생: {e}")
            # 에러 발생 시 세션 초기화하여 재시도 가능하게 함
            st.query_params.clear()
            if 'auth_flow' in st.session_state:
                del st.session_state.auth_flow
            return None
    
    else:
        # 인증 링크 표시
        auth_url, _ = st.session_state.auth_flow.authorization_url(
            prompt='consent', 
            access_type='offline',
            include_granted_scopes='true'
        )
        st.info("구글 캘린더 연동이 필요합니다.")
        st.markdown(f"**[🔗 Google Calendar 인증 링크]({auth_url})**")
        return None

    return None
