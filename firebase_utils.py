# firebase_utils.py
import streamlit as st
import firebase_admin
from firebase_admin import credentials, db, auth
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import pickle
import json
from config import SCOPES

try:
    FIREBASE_CREDENTIALS = dict(st.secrets["firebase"]) 
    DB_URL = st.secrets["database_url"] 
    google_calendar_secrets = st.secrets.get("google_calendar")
    GOOGLE_CALENDAR_CLIENT_SECRET = dict(google_calendar_secrets) if google_calendar_secrets else {}
except Exception as e:
    st.error(f"🚨 Secrets 로드 오류: {e}")
    FIREBASE_CREDENTIALS, DB_URL, GOOGLE_CALENDAR_CLIENT_SECRET = None, None, {}

@st.cache_resource
def get_db_refs():
    if not firebase_admin._apps:
        try:
            if FIREBASE_CREDENTIALS and DB_URL:
                cred = credentials.Certificate(FIREBASE_CREDENTIALS)
                firebase_admin.initialize_app(cred, {'databaseURL': DB_URL})
        except Exception as e:
            st.error(f"❌ Firebase 초기화 실패: {e}")
            return None, None, None 

    if firebase_admin._apps:
        base_ref = db.reference()
        return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)
    return None, None, None

def sanitize_path(email):
    return email.replace('.', '_')

def save_google_creds_to_firebase(safe_key, creds):
    db.reference(f'google_calendar_creds/{safe_key}').set({'creds': pickle.dumps(creds).hex()})

def load_google_creds_from_firebase(safe_key):
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    if data and 'creds' in data:
        return pickle.loads(bytes.fromhex(data['creds']))
    return None

def get_google_calendar_service(safe_key):
    creds = load_google_creds_from_firebase(safe_key)
    if creds and creds.valid:
        st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
        return
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        save_google_creds_to_firebase(safe_key, creds)
        st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
        return
    
    # Auth Flow
    if GOOGLE_CALENDAR_CLIENT_SECRET:
        flow = InstalledAppFlow.from_client_config({"installed": GOOGLE_CALENDAR_CLIENT_SECRET}, SCOPES, redirect_uri=GOOGLE_CALENDAR_CLIENT_SECRET.get("redirect_uri"))
        auth_code = st.query_params.get("code")
        if auth_code:
            flow.fetch_token(code=auth_code)
            save_google_creds_to_firebase(safe_key, flow.credentials)
            st.query_params.clear()
            st.rerun()
        else:
            auth_url, _ = flow.authorization_url(prompt='consent')
            st.markdown(f"**[Google Calendar 인증 링크]({auth_url})**")

def recover_email(safe_key):
    for path in [f'users/{safe_key}', f'doctor_users/{safe_key}', safe_key]:
        data = db.reference(path).get()
        if data and 'email' in data: return data['email']
    return None


# firebase_utils.py 에 추가 또는 수정

def check_google_connection_status(safe_key):
    """
    현재 사용자의 구글 연동 상태를 상세히 체크합니다.
    반환값: (bool: 연동여부, str: 상태메시지)
    """
    creds = load_google_creds_from_firebase(safe_key)
    
    if not creds:
        return False, "미연동 (데이터 없음)"
    
    try:
        if creds.valid:
            return True, "연동 완료 (정상)"
        elif creds.expired and creds.refresh_token:
            # 토큰 갱신 시도
            creds.refresh(Request())
            save_google_creds_to_firebase(safe_key, creds)
            return True, "연동 완료 (갱신됨)"
        else:
            return False, "연동 만료 (재인증 필요)"
    except Exception as e:
        return False, f"오류 발생 ({str(e)})"

# 기존 get_google_calendar_service 내부에 이 로직을 통합하여 
# 세션 상태에 상세 상태를 저장하도록 변경할 수 있습니다.
