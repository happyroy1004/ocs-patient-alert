import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json

from config import SCOPES

# --- 1. Firebase 및 Google OAuth 설정 로드 ---
try:
    FIREBASE_CREDENTIALS = dict(st.secrets["firebase"])
    DB_URL = st.secrets.get("database_url") or FIREBASE_CREDENTIALS.get("database_url")
    google_calendar_secrets = st.secrets.get("google_calendar")
    GOOGLE_CALENDAR_CLIENT_SECRET = dict(google_calendar_secrets) if google_calendar_secrets else {}
except Exception as e:
    st.error(f"🚨 Secrets 로드 오류: {e}")
    FIREBASE_CREDENTIALS = None
    DB_URL = None

# --- 2. DB 초기화 ---
if not firebase_admin._apps:
    try:
        if FIREBASE_CREDENTIALS and DB_URL:
            cred = credentials.Certificate(FIREBASE_CREDENTIALS)
            firebase_admin.initialize_app(cred, {'databaseURL': DB_URL})
    except Exception as e:
        st.error(f"❌ Firebase 앱 초기화 실패: {e}")

@st.cache_resource
def get_db_refs():
    try:
        base_ref = db.reference()
        users_ref = base_ref.child('users')
        doctor_users_ref = base_ref.child('doctor_users')
        def db_ref_func(path): return base_ref.child(path)
        return users_ref, doctor_users_ref, db_ref_func
    except Exception as e:
        return None, None, lambda x: None

def sanitize_path(email):
    return email.replace('.', '_') if email else "unknown"

def recover_email(safe_key):
    db_ref = db.reference()
    for path in [f'users/{safe_key}', f'doctor_users/{safe_key}']:
        try:
            data = db_ref.child(path).get()
            if data and isinstance(data, dict) and 'email' in data:
                return data['email']
        except:
            continue
    # 과거 _at_ 데이터 잔재 완벽 호환
    if "_at_" in safe_key:
        return safe_key.replace("_at_", "@").replace("_dot_", ".").replace("_", ".")
    return safe_key.replace('_', '.') if safe_key else ""

def save_google_creds_to_firebase(safe_key, creds):
    try:
        db.reference(f'google_calendar_creds/{safe_key}').set({'creds': creds.to_json()})
        return True
    except Exception as e:
        return False

def load_google_creds_from_firebase(safe_key):
    if not safe_key: return None
    try:
        # 정리해주신 google_calendar_creds 노드만 정확히 봅니다.
        data = db.reference(f'google_calendar_creds/{safe_key}').get()
        if data and 'creds' in data:
            creds_info = data['creds']
            if isinstance(creds_info, str):
                creds_info = json.loads(creds_info)
            return Credentials.from_authorized_user_info(creds_info, SCOPES)
    except Exception as e:
        pass
    return None

def get_google_calendar_service(safe_key):
    if not safe_key: return None
    creds = load_google_creds_from_firebase(safe_key)
    
    # 1. DB에 인증 정보가 있는 경우 검증
    if creds:
        if creds.valid:
            return build('calendar', 'v3', credentials=creds)
        # 만료되었지만 자동 갱신(refresh_token)이 가능한 경우
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                save_google_creds_to_firebase(safe_key, creds)
                return build('calendar', 'v3', credentials=creds)
            except Exception:
                creds = None
        else:
            # 인증 정보는 있으나 자동 갱신 토큰이 누락된 경우 (초기화)
            creds = None

    # 2. 토큰이 없거나 갱신 불가(누락)인 경우 무조건 연동 버튼 표시
    conf = GOOGLE_CALENDAR_CLIENT_SECRET
    if not conf: return None
    redirect_uri = conf.get("redirect_uri")
    flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=redirect_uri)
    auth_code = st.query_params.get("code")
    
    if auth_code:
        try:
            flow.fetch_token(code=auth_code)
            save_google_creds_to_firebase(safe_key, flow.credentials)
            st.success("✅ 구글 캘린더 연동 완료!")
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"인증 처리 실패: {e}")
    else:
        auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
        st.warning("📅 DB에 연동 정보가 없거나, 구글 오프라인 갱신 권한(refresh_token)이 만료/누락되었습니다.")
        st.markdown(f'<a href="{auth_url}" target="_self" style="text-decoration:none;"><div style="background-color:#4285F4; color:white; padding:10px; border-radius:5px; text-align:center;">구글 계정 (재)연동하기</div></a>', unsafe_allow_html=True)
        st.info("💡 위 버튼을 눌러 딱 1번만 재연동 하시면 이후부터는 영구적으로 자동 갱신됩니다.")
    return None
