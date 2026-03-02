import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
import pickle
import time
import os

# 구글 API 관련 임포트 (ImportError 방지를 위해 상단 배치)
try:
    from google_auth_oauthlib.flow import Flow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
except ImportError as e:
    st.error(f"❌ 필수 라이브러리 로드 실패: {e}. requirements.txt 설정을 확인하세요.")

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

# --- 2. Google 인증 관리 (저장 보장 및 성공 메시지) ---
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
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 기존 DB 로드 시도
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

    # OAuth 설정
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

    # Flow 객체 세션 고정 (Missing code verifier 방지 핵심)
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf.get("redirect_uri")
        )

    # 인증 코드 처리
    auth_code = st.query_params.get("code")
    if auth_code and 'auth_flow' in st.session_state:
        try:
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            
            if save_google_creds_to_firebase(safe_key, new_creds):
                # ✅ 저장 성공 시 명확한 알림
                st.success("🎉 구글 캘린더 권한이 승인되어 데이터베이스에 안전하게 저장되었습니다!")
                st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
                
                st.query_params.clear()
                if 'auth_flow' in st.session_state: del st.session_state.auth_flow
                time.sleep(2)
                st.rerun()
        except Exception as e:
            st.error(f"⚠️ 인증 처리 오류 (세션 만료 등): {e}")
            st.query_params.clear()
            if 'auth_flow' in st.session_state: del st.session_state.auth_flow

    # 새 인증 링크 생성
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
