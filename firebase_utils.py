# firebase_utils.py
import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import pickle
from config import SCOPES

try:
    FIREBASE_CREDENTIALS = dict(st.secrets["firebase"])
    DB_URL = st.secrets["database_url"]
    GOOGLE_CALENDAR_CLIENT_SECRET = dict(st.secrets.get("google_calendar", {}))
except Exception as e:
    st.error(f"Secrets 로드 실패: {e}")
    FIREBASE_CREDENTIALS, DB_URL, GOOGLE_CALENDAR_CLIENT_SECRET = None, None, {}

@st.cache_resource
def get_db_refs():
    if not firebase_admin._apps:
        try:
            cred = credentials.Certificate(FIREBASE_CREDENTIALS)
            firebase_admin.initialize_app(cred, {'databaseURL': DB_URL})
        except: return None, None, None
    base_ref = db.reference()
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

def sanitize_path(email): return email.replace('.', '_')

def save_google_creds_to_firebase(safe_key, creds):
    db.reference(f'google_calendar_creds/{safe_key}').set({'creds': pickle.dumps(creds).hex()})

def load_google_creds_from_firebase(safe_key):
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    return pickle.loads(bytes.fromhex(data['creds'])) if data and 'creds' in data else None

def check_google_connection_status(safe_key):
    """DB의 크리덴셜 상태를 확인하여 연동 여부를 반환합니다."""
    creds = load_google_creds_from_firebase(safe_key)
    if not creds: return False, "미연동"
    try:
        if creds.valid: return True, "정상 연동"
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
            save_google_creds_to_firebase(safe_key, creds)
            return True, "연동 갱신됨"
        return False, "인증 만료"
    except: return False, "연동 오류"

def get_google_calendar_service(safe_key):
    creds = load_google_creds_from_firebase(safe_key)
    if creds and (creds.valid or (creds.expired and creds.refresh_token)):
        if creds.expired: creds.refresh(Request()); save_google_creds_to_firebase(safe_key, creds)
        st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
        return
    
    if GOOGLE_CALENDAR_CLIENT_SECRET:
        flow = InstalledAppFlow.from_client_config({"installed": GOOGLE_CALENDAR_CLIENT_SECRET}, SCOPES, redirect_uri=GOOGLE_CALENDAR_CLIENT_SECRET.get("redirect_uri"))
        auth_code = st.query_params.get("code")
        if auth_code:
            flow.fetch_token(code=auth_code)
            save_google_creds_to_firebase(safe_key, flow.credentials)
            st.query_params.clear(); st.rerun()
        else:
            auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
            st.markdown(f"**[Google Calendar 인증 링크]({auth_url})**")

def recover_email(safe_key):
    data = db.reference(f'users/{safe_key}').get() or db.reference(f'doctor_users/{safe_key}').get()
    return data.get('email') if data else None
