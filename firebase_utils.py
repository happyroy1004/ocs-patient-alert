import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
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

# --- 3. Google 인증 관리 (개선 버전) ---

def save_google_creds_to_firebase(safe_key, creds):
    """Pickle 대신 안정적인 JSON 형태로 저장합니다."""
    ref = db.reference(f'google_calendar_creds/{safe_key}')
    # creds.to_json()은 모든 필요한 토큰 정보를 포함합니다.
    ref.set({'creds': creds.to_json()})

def load_google_creds_from_firebase(safe_key):
    """JSON 데이터를 다시 Credentials 객체로 복원합니다."""
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    if data and 'creds' in data:
        try:
            return Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
        except Exception as e:
            st.error(f"Creds 로드 실패: {e}")
    return None

def get_google_calendar_service(safe_key):
    # 0. 세션에 이미 서비스가 로드되어 있다면 즉시 반환
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

    # 1단계: 기존 저장된 자격 증명 확인 및 갱신
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
            except Exception:
                pass # 갱신 실패 시 다시 인증 유도

    # 2단계: 인증 코드 처리 (이 부분이 핵심 수정 사항)
    # Streamlit 리런 시 URL의 ?code= 파라미터를 감지합니다.
    auth_code = st.query_params.get("code")
    
    if auth_code:
        try:
            # 새로 Flow를 생성하여 fetch_token 수행 (state 검증을 생략하거나 자동 처리)
            # PKCE 'code_verifier' 문제를 피하기 위해 세션에 저장된 flow 대신 즉석 생성 사용
            flow = Flow.from_client_config(
                client_config, 
                scopes=SCOPES, 
                redirect_uri=conf.get("redirect_uri")
            )
            
            # fetch_token 실행 시 내부적으로 code_verifier 없이도 동작하도록 처리 시도
            # (최신 라이브러리 환경에서 가장 안정적인 방식)
            flow.fetch_token(code=auth_code)
            new_creds = flow.credentials
            
            save_google_creds_to_firebase(safe_key, new_creds)
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
            
            # 쿼리 스트링 정리 후 페이지 리런
            st.query_params.clear()
            st.rerun()
            
        except Exception as e:
            st.warning(f"⚠️ 인증 처리 중 오류가 발생했습니다: {e}")
            st.query_params.clear()
            # 오류 발생 시 다시 시작할 수 있도록 중단
            st.stop()

    # 3단계: 인증 URL 생성 및 사용자 클릭 유도
    # 이미 인증 코드가 URL에 없는 경우에만 노출
    flow = Flow.from_client_config(
        client_config, 
        scopes=SCOPES, 
        redirect_uri=conf.get("redirect_uri")
    )
    
    auth_url, _ = flow.authorization_url(
        prompt='consent', 
        access_type='offline',
        include_granted_scopes='true'
    )
    
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 구글 계정 연동하기]({auth_url})**")
    return None
