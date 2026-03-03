import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from google.oauth2 import id_token
from google.auth.transport import requests as google_requests
import json

# [1] Firebase 초기화 (TOML 구조에 맞게 수정)
if not firebase_admin._apps:
    try:
        fb_conf = dict(st.secrets["firebase"])
        # 보내주신 TOML에 database_url로 되어 있으므로 이를 명시적으로 사용합니다.
        db_url = fb_conf.get("database_url")
        
        cred = credentials.Certificate(fb_conf)
        firebase_admin.initialize_app(cred, {
            'databaseURL': db_url
        })
    except Exception as e:
        st.error(f"Firebase 초기화 실패: {e}")
        st.stop()

# [2] ui_manager.py 대응 (반환값 3개)
def get_db_refs():
    """
    users_ref, doctor_users_ref, db_ref_func = get_db_refs() 호출 대응
    """
    users_ref = db.reference('users')
    doctor_users_ref = db.reference('doctor_users')
    
    def db_ref_func(path):
        return db.reference(path)
        
    return users_ref, doctor_users_ref, db_ref_func

# [3] 유틸리티 함수
def sanitize_path(email):
    return email.replace('.', '_') if email else "unknown"

def recover_email(sanitized_email):
    return sanitized_email.replace('_', '.') if sanitized_email else ""

# [4] 데이터 저장 및 로드
SCOPES = [
    'https://www.googleapis.com/auth/calendar.events',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid'
]

def save_google_creds_to_firebase(clean_key, creds):
    try:
        db.reference(f'google_calendar_creds/{clean_key}').set({
            'creds': creds.to_json()
        })
        return True
    except:
        return False

def load_google_creds_from_firebase(safe_key):
    """ImportError 방지를 위한 명시적 함수"""
    if not safe_key:
        return None
    clean_key = sanitize_path(safe_key)
    try:
        data = db.reference(f'google_calendar_creds/{clean_key}').get()
        if data and 'creds' in data:
            return Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
    except:
        pass
    return None

# [5] 구글 캘린더 서비스 빌드
def get_google_calendar_service(safe_key=None):
    # 구글 인증 응답 처리
    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            
            # ID 토큰에서 이메일 추출
            id_info = id_token.verify_oauth2_token(creds.id_token, google_requests.Request(), conf["client_id"])
            google_email = id_info.get('email')
            
            # safe_key가 없으면 연동한 구글 계정 이메일을 키로 사용
            target_key = safe_key if safe_key else google_email
            if target_key:
                save_google_creds_to_firebase(sanitize_path(target_key), creds)
                st.success(f"✅ {google_email} 계정 연동 성공!")
                st.query_params.clear()
                st.rerun()
        except Exception as e:
            st.error(f"인증 처리 중 오류: {e}")

    # 서비스 로드 로직
    creds = load_google_creds_from_firebase(safe_key)
    if creds:
        try:
            if not creds.valid and creds.refresh_token:
                creds.refresh(Request())
                save_google_creds_to_firebase(sanitize_path(safe_key), creds)
            return build('calendar', 'v3', credentials=creds)
        except:
            pass

    # 연동 버튼 표시
    try:
        conf = dict(st.secrets["google_calendar"])
        flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
        auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
        st.info("📅 구글 캘린더 연동이 필요합니다.")
        st.markdown(
            f'<a href="{auth_url}" target="_self" style="text-decoration:none;">'
            f'<div style="background-color:#4285F4; color:white; padding:10px; border-radius:5px; text-align:center;">'
            f'구글 계정 연동하기</div></a>', 
            unsafe_allow_html=True
        )
    except:
        st.warning("secrets의 google_calendar 설정을 확인해주세요.")
    return None
