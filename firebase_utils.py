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

# [1] Firebase 초기화 - URL 매칭 로직 강화
if not firebase_admin._apps:
    try:
        # st.secrets["firebase"] 내용을 딕셔너리로 가져옴
        fb_conf = dict(st.secrets["firebase"])
        
        # TOML에 적힌 database_url 또는 databaseURL을 모두 찾음
        db_url = fb_conf.get("database_url") or fb_conf.get("databaseURL")
        
        if not db_url:
            st.error("❌ Secrets에 'database_url' 주소가 없습니다. 설정을 확인해주세요.")
            st.stop()
            
        cred = credentials.Certificate(fb_conf)
        firebase_admin.initialize_app(cred, {
            'databaseURL': db_url  # Firebase 내부에서는 이 키 이름을 사용함
        })
    except Exception as e:
        st.error(f"Firebase 초기화 에러: {e}")
        st.stop()

# [2] ui_manager.py 32번 라인 대응 (반환값 3개 보장)
def get_db_refs():
    """
    users_ref, doctor_users_ref, db_ref_func = get_db_refs() 호출을 완벽하게 지원
    """
    try:
        users_ref = db.reference('users')
        doctor_users_ref = db.reference('doctor_users')
        
        # 세 번째 인자: 경로를 받아서 reference를 반환하는 함수
        def db_ref_func(path):
            return db.reference(path)
            
        return users_ref, doctor_users_ref, db_ref_func
    except Exception as e:
        # 여기서 ValueError가 나지 않도록 방어 로직 추가
        st.error(f"DB 참조 실패 (URL 설정 확인 필요): {e}")
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
        try:
            if not creds.valid and creds.refresh_token:
                creds.refresh(Request())
                save_google_creds_to_firebase(sanitize_path(safe_key), creds)
            return build('calendar', 'v3', credentials=creds)
        except: pass

    # 연동 버튼 (OAuth 정보가 있을 때만 표시)
    try:
        conf = dict(st.secrets["google_calendar"])
        flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
        auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
        st.info("📅 구글 캘린더 연동이 필요합니다.")
        st.markdown(f'''
            <a href="{auth_url}" target="_self" style="text-decoration:none;">
                <div style="background-color:#4285F4; color:white; padding:10px; border-radius:5px; text-align:center;">
                    구글 계정 연동하기
                </div>
            </a>''', unsafe_allow_html=True)
    except: pass
    return None
    
def save_google_creds_to_firebase(safe_key, creds):
    """표준 경로: google_calendar_creds/{safe_key} 에 저장"""
    # 💡 유의: safe_key는 이미 sanitize_path()를 거친 상태여야 함
    creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
    # 가급적 pickle 보다는 호환성이 좋은 to_json() 권장
    creds_ref.set({'creds': creds.to_json()})

def load_google_creds_from_firebase(safe_key):
    """표준 경로에서만 데이터를 긁어옵니다."""
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    
    if data and 'creds' in data:
        creds_info = data['creds']
        # JSON 문자열인 경우 파싱
        if isinstance(creds_info, str):
            creds_info = json.loads(creds_info)
        return Credentials.from_authorized_user_info(creds_info, SCOPES)
    return None
