# firebase_utils.py

import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os
import pickle

from config import SCOPES

# --- 1. 환경 설정 로드 ---
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


# --- 2. DB 초기화 (캐싱 적용됨) ---
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


# --- 3. Creds 관리 ---
def sanitize_path(email):
    if not email: return ""
    return email.replace('@', '_at_').replace('.', '_dot_')

def save_google_creds_to_firebase(safe_key, creds):
    creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
    pickled_creds = pickle.dumps(creds)
    creds_ref.set({'creds': pickled_creds.hex()})

def load_google_creds_from_firebase(safe_key):
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    if data and 'creds' in data:
        return pickle.loads(bytes.fromhex(data['creds']))
    return None


# --- 4. 서비스 로드 및 인증 흐름 ---
def get_google_calendar_service(safe_key):
    user_id_safe = safe_key
    st.session_state.google_calendar_service = None
    
    creds = load_google_creds_from_firebase(user_id_safe)

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

    redirect_uri = GOOGLE_CALENDAR_CLIENT_SECRET.get("redirect_uri")
    flow = Flow.from_client_config(client_config, scopes=SCOPES, redirect_uri=redirect_uri)
    
    auth_code = st.query_params.get("code")
    
    if auth_code:
        try:
            temp_ref = db.reference(f'temp_auth/{user_id_safe}')
            temp_data = temp_ref.get()
            
            if temp_data and 'code_verifier' in temp_data:
                flow.code_verifier = temp_data['code_verifier'] 
                
            flow.fetch_token(code=auth_code)
            new_creds = flow.credentials
            
            save_google_creds_to_firebase(user_id_safe, new_creds)
            temp_ref.delete() 
            
            st.success("✅ 구글 캘린더 연동 완료!")
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"❌ 토큰 교환 실패: {e}")
            st.query_params.clear()
    else:
        auth_url, state = flow.authorization_url(prompt='consent', access_type='offline')
        
        code_verifier = getattr(flow, 'code_verifier', None)
        if code_verifier:
            db.reference(f'temp_auth/{user_id_safe}').set({
                'code_verifier': code_verifier
            })
            
        st.warning("⚠️ 구글 캘린더 연동이 필요합니다.")
        st.markdown(f"**[Google Calendar 인증 링크]({auth_url})**")
        st.info("링크 클릭 후 권한 승인을 완료하면 이 페이지로 돌아옵니다.")
        return None


# --- 5. 최적화된 이메일 복구 (캐싱 적용 & 경로 단축) ---
@st.cache_data(ttl=3600) # 💡 추가됨: 1시간 동안 결과 기억 (매번 DB 조회 안 함)
def recover_email(safe_key):
    """Firebase에서 safe_key에 해당하는 이메일을 찾습니다 (최적화 완료)."""
    base = db.reference()
    
    # 💡 최적화: 루트 경로(safe_key) 제거, 정확한 2곳만 확인
    for p in [f'users/{safe_key}', f'doctor_users/{safe_key}']:
        data = base.child(p).get()
        if data and 'email' in data: return data['email']
    return None
