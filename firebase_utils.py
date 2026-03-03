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

# [1] Firebase 초기화 (URL 체크 보강)
if not firebase_admin._apps:
    try:
        fb_conf = dict(st.secrets["firebase"])
        # databaseURL이 secrets에 있는지 확실히 확인
        db_url = fb_conf.get("databaseURL") or fb_conf.get("database_url")
        
        if not db_url:
            st.error("❌ Firebase databaseURL이 st.secrets에 설정되지 않았습니다.")
            st.stop()
            
        cred = credentials.Certificate(fb_conf)
        firebase_admin.initialize_app(cred, {
            'databaseURL': db_url
        })
    except Exception as e:
        st.error(f"Firebase 초기화 실패: {e}")
        st.stop()

# [2] ui_manager.py 32번 라인 대응 (반환값 개수 3개)
def get_db_refs():
    """
    ui_manager.py: users_ref, doctor_users_ref, db_ref_func = get_db_refs() 대응
    """
    # 초기화가 정상적이지 않을 때를 대비해 예외 처리
    try:
        users_ref = db.reference('users')
        doctor_users_ref = db.reference('doctor_users')
        
        def db_ref_func(path):
            return db.reference(path)
            
        return users_ref, doctor_users_ref, db_ref_func
    except Exception as e:
        st.error(f"DB 참조 획득 중 오류 발생: {e}")
        # 에러 발생 시 앱이 죽지 않도록 더미 데이터라도 반환
        return None, None, lambda x: None

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
    except: return False

def load_google_creds_from_firebase(safe_key):
    if not safe_key: return None
    clean_key = sanitize_path(safe_key)
    try:
        data = db.reference(f'google_calendar_creds/{clean_key}').get()
        if data and 'creds' in data:
            return Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
    except: pass
    return None

# [5] 구글 캘린더 서비스 빌드
def get_google_calendar_service(safe_key=None):
    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            id_info = id_token.verify_oauth2_token(creds.id_token, google_requests.Request(), conf["client_id"])
            google_email = id_info.get('email')
            target_key = safe_key if safe_key else google_email
            if target_key:
                save_google_creds_to_firebase(sanitize_path(target_key), creds)
                st.success(f"✅ {google_email} 연동 성공!")
                st.query_params.clear()
                st.rerun()
        except: pass

    creds = load_google_creds_from_firebase(safe_key)
    if creds:
        if not creds.valid and creds.refresh_token:
            creds.refresh(Request())
            save_google_creds_to_firebase(sanitize_path(safe_key), creds)
        return build('calendar', 'v3', credentials=creds)

    try:
        conf = dict(st.secrets["google_calendar"])
        flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
        auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
        st.info("📅 구글 캘린더 연동이 필요합니다.")
        st.markdown(f'<a href="{auth_url}" target="_self" style="text-decoration:none;"><div style="background-color:#4285F4; color:white; padding:10px; border-radius:5px; text-align:center;">구글 계정 연동하기</div></a>', unsafe_allow_html=True)
    except: pass
    return None
