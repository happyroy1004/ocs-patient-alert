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

# --- 설정 로드 ---
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
    st.error(f"🚨 설정 로드 오류: {e}")
    FIREBASE_CREDENTIALS = None
    DB_URL = None
    GOOGLE_CALENDAR_CLIENT_SECRET = {}


# --- 1. DB 레퍼런스 및 초기화 ---
@st.cache_resource
def get_db_refs():
    if not firebase_admin._apps:
        try:
            if FIREBASE_CREDENTIALS and DB_URL:
                creds_init = FIREBASE_CREDENTIALS.copy()
                if 'FIREBASE_DATABASE_URL' in creds_init: 
                    del creds_init['FIREBASE_DATABASE_URL']
                cred = credentials.Certificate(creds_init)
                firebase_admin.initialize_app(cred, {'databaseURL': DB_URL})
        except Exception as e:
            st.error(f"❌ Firebase 초기화 실패: {e}")
            return None, None, None 

    if firebase_admin._apps:
        base_ref = db.reference()
        return base_ref.child('users'), base_ref.child('doctor_users'), lambda p: base_ref.child(p)
    return None, None, None


# --- 2. Google Calendar 인증 및 Creds 관리 ---
def sanitize_path(email):
    """
    이메일 주소를 Firebase Realtime DB에서 안전하게 사용할 수 있는 고유 키로 변환합니다.
    예: test@naver.com -> test_at_naver_dot_com
    """
    if not email:
        return ""
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


# --- 3. Google Calendar Service 로드/인증 흐름 ---
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

    # 4. OAuth 인증 플로우
    redirect_uri = GOOGLE_CALENDAR_CLIENT_SECRET.get("redirect_uri")
    flow = Flow.from_client_config(client_config, scopes=SCOPES, redirect_uri=redirect_uri)
    
    auth_code = st.query_params.get("code")
    
    if auth_code:
        try:
            # 💡 [핵심 해결책] Streamlit 세션 초기화로 인해 날아간 verifier를 Firebase에서 복구
            temp_ref = db.reference(f'temp_auth/{user_id_safe}')
            temp_data = temp_ref.get()
            
            if temp_data and 'code_verifier' in temp_data:
                flow.code_verifier = temp_data['code_verifier'] # 보관해둔 키 주입!
                
            flow.fetch_token(code=auth_code)
            new_creds = flow.credentials
            
            save_google_creds_to_firebase(user_id_safe, new_creds)
            temp_ref.delete() # 사용이 끝난 임시 키는 깔끔하게 삭제
            
            st.success("✅ 구글 캘린더 연동 완료!")
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"❌ 토큰 교환 실패: {e}")
            st.query_params.clear()
    else:
        auth_url, state = flow.authorization_url(prompt='consent', access_type='offline')
        
        # 💡 [핵심 해결책] 구글로 넘어가기 전, 생성된 verifier를 Firebase에 안전하게 보관
        code_verifier = getattr(flow, 'code_verifier', None)
        if code_verifier:
            db.reference(f'temp_auth/{user_id_safe}').set({
                'code_verifier': code_verifier
            })
            
        st.warning("⚠️ 구글 캘린더 연동이 필요합니다.")
        st.markdown(f"**[Google Calendar 인증 링크]({auth_url})**")
        st.info("링크 클릭 후 권한 승인을 완료하면 이 페이지로 돌아옵니다.")
        return None

def recover_email(safe_key):
    """
    Firebase에서 safe_key(통일된 형식)에 해당하는 실제 이메일을 단번에 찾습니다.
    """
    db_ref = db.reference()
    
    # 💡 최적화: 이제 safe_key 자체가 고유하고 정확하므로, 
    # users와 doctor_users 경로만 딱 두 번 조회하고 끝냅니다.
    # (과거처럼 safe_key 경로 전체를 불필요하게 뒤지지 않습니다)
    
    paths_to_check = [f'users/{safe_key}', f'doctor_users/{safe_key}']
    
    for path in paths_to_check:
        try:
            data = db_ref.child(path).get()
            if data and 'email' in data:
                return data['email']
        except Exception as e:
            # 에러 로그를 남기면 나중에 문제 찾기 편합니다
            print(f"이메일 복구 중 오류: {e}") 
            continue
            
    return None
