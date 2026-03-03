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

# [1] 권한 범위 설정
SCOPES = [
    'https://www.googleapis.com/auth/calendar.events',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid'
]

# [2] Firebase 초기화
if not firebase_admin._apps:
    try:
        fb_dict = dict(st.secrets["firebase"])
        cred = credentials.Certificate(fb_dict)
        firebase_admin.initialize_app(cred, {
            'databaseURL': fb_dict.get("databaseURL")
        })
    except Exception as e:
        print(f"Firebase 초기화 에러: {e}")

# [3] UI_MANAGER가 요구하는 유틸리티 함수들
def get_db_refs():
    try:
        return {
            "users": db.reference('users'),
            "calendar": db.reference('google_calendar_creds'),
            "settings": db.reference('settings')
        }
    except:
        return {}

def sanitize_path(email):
    return email.replace('.', '_') if email else "unknown"

def recover_email(sanitized_email):
    return sanitized_email.replace('_', '.') if sanitized_email else ""

# [4] 데이터 저장 및 로드 함수 (이름 매칭 완료)
def save_google_creds_to_firebase(clean_key, creds):
    """구글 인증 정보를 DB에 저장"""
    try:
        db.reference(f'google_calendar_creds/{clean_key}').set({
            'creds': creds.to_json()
        })
        return True
    except Exception as e:
        st.error(f"저장 오류: {e}")
        return False

def load_google_creds_from_firebase(safe_key):
    """ui_manager에서 요구하는 이름 그대로 구현"""
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

# [5] 구글 캘린더 서비스 빌드 함수
def get_google_calendar_service(safe_key=None):
    auth_code = st.query_params.get("code")

    # 인증 콜백 처리
    if auth_code:
        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            
            # ID 토큰에서 이메일 추출
            id_info = id_token.verify_oauth2_token(creds.id_token, google_requests.Request(), conf["client_id"])
            google_email = id_info.get('email')
            
            if google_email:
                target_key = safe_key if safe_key else google_email
                clean_key = sanitize_path(target_key)
                if save_google_creds_to_firebase(clean_key, creds):
                    st.success(f"✅ {google_email} 연동 성공!")
                    st.query_params.clear()
                    st.rerun()
        except Exception as e:
            st.error(f"인증 에러: {e}")
            return None

    # 서비스 빌드
    creds = load_google_creds_from_firebase(safe_key)
    if creds:
        if not creds.valid and creds.refresh_token:
            creds.refresh(Request())
            save_google_creds_to_firebase(sanitize_path(safe_key), creds)
        return build('calendar', 'v3', credentials=creds)

    # 연동 버튼 출력
    try:
        conf = dict(st.secrets["google_calendar"])
        flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
        auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
        st.info("📅 구글 캘린더 연동이 필요합니다.")
        st.markdown(f'<a href="{auth_url}" target="_self" style="text-decoration:none;"><div style="background-color:#4285F4; color:white; padding:10px; border-radius:5px; text-align:center;">구글 계정 연동하기</div></a>', unsafe_allow_html=True)
    except:
        pass
    return None
