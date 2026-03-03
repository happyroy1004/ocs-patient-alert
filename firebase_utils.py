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

# [1] Firebase 초기화
if not firebase_admin._apps:
    try:
        fb_dict = dict(st.secrets["firebase"])
        cred = credentials.Certificate(fb_dict)
        firebase_admin.initialize_app(cred, {
            'databaseURL': fb_dict.get("databaseURL")
        })
    except Exception as e:
        print(f"Firebase 초기화 에러: {e}")

# [2] ui_manager.py의 32번 라인 에러 해결 (반환값 3개 보장)
def get_db_refs():
    """
    ui_manager.py: users_ref, doctor_users_ref, db_ref_func = get_db_refs() 대응
    """
    users_ref = db.reference('users')
    doctor_users_ref = db.reference('doctor_users')
    
    def db_ref_func(path):
        return db.reference(path)
        
    return users_ref, doctor_users_ref, db_ref_func

# [3] 유틸리티 함수들
def sanitize_path(email):
    if not email: return "unknown"
    return email.replace('.', '_')

def recover_email(sanitized_email):
    if not sanitized_email: return ""
    return sanitized_email.replace('_', '.')

# [4] 구글 인증 관련 (ui_manager가 import하는 모든 이름 포함)
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
    """ImportError를 막기 위해 반드시 필요한 함수명"""
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

# [5] 메인 서비스 빌더
def get_google_calendar_service(safe_key=None):
    # 1. 구글 인증 응답(code) 처리
    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            
            # ID 토큰에서 이메일 추출하여 DB 경로 확보
            id_info = id_token.verify_oauth2_token(creds.id_token, google_requests.Request(), conf["client_id"])
            google_email = id_info.get('email')
            
            # safe_key가 없으면 구글 이메일로 저장
            target_key = safe_key if safe_key else google_email
            if target_key:
                save_google_creds_to_firebase(sanitize_path(target_key), creds)
                st.success("✅ 연동 완료!")
                st.query_params.clear()
                st.rerun()
        except:
            pass

    # 2. 서비스 로드
    creds = load_google_creds_from_firebase(safe_key)
    if creds:
        try:
            if not creds.valid and creds.refresh_token:
                creds.refresh(Request())
                save_google_creds_to_firebase(sanitize_path(safe_key), creds)
            return build('calendar', 'v3', credentials=creds)
        except:
            pass

    # 3. 연동 버튼 표시
    try:
        conf = dict(st.secrets["google_calendar"])
        flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
        auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
        st.info("📅 구글 캘린더 연동이 필요합니다.")
        st.markdown(f'<a href="{auth_url}" target="_self" style="text-decoration:none;"><div style="background-color:#4285F4; color:white; padding:10px; border-radius:5px; text-align:center;">구글 계정 연동하기</div></a>', unsafe_allow_html=True)
    except:
        st.error("OAuth 설정(secrets)이 올바르지 않습니다.")
        
    return None
