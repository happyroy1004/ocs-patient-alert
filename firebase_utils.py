import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
import time
import json

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
            if 'FIREBASE_DATABASE_URL' in creds_dict: 
                del creds_dict['FIREBASE_DATABASE_URL']
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

# --- 3. Google 인증 및 저장 (연속성 보장 버전) ---
def save_google_creds_to_firebase(safe_key, creds):
    try:
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        ref.set({'creds': pickle.dumps(creds).hex()})
        return True
    except Exception as e:
        st.error(f"❌ DB 저장 실패: {e}")
        return False

def load_google_creds_from_firebase(safe_key):
    try:
        data = db.reference(f'google_calendar_creds/{safe_key}').get()
        if data and 'creds' in data:
            return pickle.loads(bytes.fromhex(data['creds']))
    except:
        return None
    return None

def get_google_calendar_service(safe_key):
    # 1. 이미 인증된 서비스가 있는 경우 즉시 반환
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 2. DB에서 기존 토큰 로드 (있으면 바로 통과)
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

    # 3. OAuth 설정 구성
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

    # 4. 🔑 세션 유지를 위한 Flow 객체 생성 및 고정
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf.get("redirect_uri")
        )

    # 5. 인증 결과 처리 (Redirect 후 돌아온 시점)
    auth_code = st.query_params.get("code")
    if auth_code and 'auth_flow' in st.session_state:
        try:
            # 세션에 저장된 verifier를 사용하여 토큰 교환
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            
            if save_google_creds_to_firebase(safe_key, new_creds):
                st.success("✅ 구글 캘린더 연동 성공! 잠시 후 화면이 갱신됩니다.")
                st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
                
                # 주소창 정리 및 세션 정리
                st.query_params.clear()
                if 'auth_flow' in st.session_state: del st.session_state.auth_flow
                time.sleep(1)
                st.rerun()
        except Exception as e:
            st.warning("⚠️ 인증 세션이 만료되었습니다. 다시 시도해 주세요.")
            st.query_params.clear()
            if 'auth_flow' in st.session_state: del st.session_state.auth_flow

    # 6. 인증이 필요한 경우 링크 표시 (URL에 현재 사용자의 safe_key를 포함시켜 보냄)
    if 'auth_flow' not in st.session_state:
         st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf.get("redirect_uri")
        )

    # state 파라미터를 사용해 safe_key를 인코딩하여 전달 (선택사항, 현재는 기본 prompt 유지)
    auth_url, _ = st.session_state.auth_flow.authorization_url(
        prompt='consent', access_type='offline', include_granted_scopes='true'
    )
    
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 구글 계정 연동하기]({auth_url})**")
    return None
