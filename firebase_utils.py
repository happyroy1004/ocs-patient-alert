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
    """Firebase DB 참조 객체를 초기화하고 반환합니다."""
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

# --- 2. 유틸리티 함수 (ImportError 방지를 위해 ui_manager에서 호출하는 모든 함수 포함) ---
def sanitize_path(email):
    """이메일의 마침표를 언더바로 치환하여 Firebase 경로로 사용 가능하게 합니다."""
    return email.replace('.', '_')

def recover_email(safe_key):
    """치환된 키를 다시 이메일 형식으로 복구합니다."""
    return safe_key.replace('_', '.')

# --- 3. Google 자격 증명 저장/로드 ---
def save_google_creds_to_firebase(safe_key, creds):
    """인증 정보를 DB에 저장하고, 저장이 완료되었는지 확인합니다."""
    try:
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        ref.set({'creds': pickle.dumps(creds).hex()})
        # 데이터가 서버에 도달할 때까지 아주 짧은 대기 후 확인
        time.sleep(0.5)
        return ref.get() is not None
    except Exception as e:
        st.error(f"❌ 자격 증명 저장 중 오류 발생: {e}")
        return False

def load_google_creds_from_firebase(safe_key):
    """Firebase에서 기존 인증 정보를 불러옵니다."""
    try:
        data = db.reference(f'google_calendar_creds/{safe_key}').get()
        if data and 'creds' in data:
            return pickle.loads(bytes.fromhex(data['creds']))
    except:
        return None
    return None

# --- 4. Google 캘린더 서비스 로직 ---
def get_google_calendar_service(safe_key):
    """인증 흐름을 관리하고 구글 서비스 객체를 반환합니다."""
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # (A) DB에서 기존 자격 증명 로드 시도
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

    # (B) OAuth 설정 및 Flow 생성
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

    # Flow 객체를 세션에 고정하여 보안 키(Verifier) 유실 방지
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf.get("redirect_uri")
        )

    # (C) 인증 후 리디렉션된 경우 처리
    auth_code = st.query_params.get("code")
    if auth_code and 'auth_flow' in st.session_state:
        try:
            with st.spinner("🔄 정보를 안전하게 저장 중입니다..."):
                st.session_state.auth_flow.fetch_token(code=auth_code)
                new_creds = st.session_state.auth_flow.credentials
                
                # DB 저장 시도 및 확인
                if save_google_creds_to_firebase(safe_key, new_creds):
                    st.success("✅ 구글 캘린더 연동 성공!")
                    st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
                    st.query_params.clear()
                    if 'auth_flow' in st.session_state: del st.session_state.auth_flow
                    time.sleep(2) # DB 쓰기 안정화 대기
                    st.rerun()
                else:
                    st.error("❌ 저장에 실패했습니다. Firebase 권한을 확인하세요.")
        except Exception as e:
            st.warning("⚠️ 세션이 만료되었습니다. 다시 시도해 주세요.")
            st.query_params.clear()
            if 'auth_flow' in st.session_state: del st.session_state.auth_flow

    # (D) 인증이 필요한 경우 링크 표시
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
