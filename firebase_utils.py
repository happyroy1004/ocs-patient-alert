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

# [2] ui_manager.py 32번 라인 대응 (반환값 3개로 수정)
def get_db_refs():
    """
    ui_manager.py의 다음 코드를 지원합니다:
    users_ref, doctor_users_ref, db_ref_func = get_db_refs()
    """
    users_ref = db.reference('users')
    doctor_users_ref = db.reference('doctor_users') # 혹은 적절한 경로
    
    # 세 번째 인자인 db_ref_func는 필요할 때 경로를 인자로 받는 함수 형태여야 함
    def db_ref_func(path):
        return db.reference(path)
        
    return users_ref, doctor_users_ref, db_ref_func

# [3] 나머지 유틸리티 함수들
def sanitize_path(email):
    return email.replace('.', '_') if email else "unknown"

def recover_email(sanitized_email):
    return sanitized_email.replace('_', '.') if sanitized_email else ""

# [4] 데이터 저장 및 로드 (ui_manager가 호출하는 이름들)
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
    if not safe_key:
        return None
    clean_key = sanitize_path(safe_key)
    data = db.reference(f'google_calendar_creds/{clean_key}').get()
    if data and 'creds' in data:
        try:
            return Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
        except:
            return None
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
            
            if google_email:
                target_key = safe_key if safe_key else google_email
                clean_key = sanitize_path(target_key)
                if save_google_creds_to_firebase(clean_key, creds):
                    st.success(f"✅ {google_email} 연동 성공!")
                    st.query_params.clear()
                    st.rerun()
        except:
            pass

    creds = load_google_creds_from_firebase(safe_key)
    if creds:
        if not creds.valid and creds.refresh_token:
            creds.refresh(Request())
            save_google_creds_to_firebase(sanitize_path(safe_key), creds)
        return build('calendar', 'v3', credentials=creds)

    # 연동 UI 출력
    try:
        conf = dict(st.secrets["google_calendar"])
        flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
        auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
        st.info("📅 구글 캘린더 연동이 필요합니다.")
        st.markdown(f'<a href="{auth_url}" target="_self" style="text-decoration:none;"><div style="background-color:#4285F4; color:white; padding:10px; border-radius:5px; text-align:center;">구글 계정 연동하기</div></a>', unsafe_allow_html=True)
    except:
        pass
    return None
