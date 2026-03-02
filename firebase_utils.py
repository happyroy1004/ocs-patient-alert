import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
import time

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

# --- 2. Google 자격 증명 저장/로드 ---
def save_google_creds_to_firebase(safe_key, creds):
    try:
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        ref.set({'creds': pickle.dumps(creds).hex()})
        return True
    except Exception as e:
        st.error(f"❌ DB 저장 실패: {e}")
        return False

def load_google_creds_from_firebase(safe_key):
    try:
        data = db.reference(f'google_calendar_creds/{safe_key}').get()
        if data and 'creds' in data:
            return pickle.loads(bytes.fromhex(data['creds']))
    except:
        return None
    return None

# --- 3. 핵심: get_google_calendar_service ---
def get_google_calendar_service(safe_key):
    # (1) 이미 인증된 서비스가 있다면 즉시 반환
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

    # (2) Firebase에서 기존 토큰 확인 (성공 시 새로고침 불필요)
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
            except: pass

    # (3) 인증 진행 중인지 확인 (URL에 code가 있는 경우)
    auth_code = st.query_params.get("code")
    
    # 💡 팁: 인증 페이지로 나갈 때 현재 safe_key를 state에 담아 보냅니다.
    # 돌아왔을 때 safe_key가 유실되지 않도록 하기 위함입니다.
    if auth_code:
        if 'auth_flow' in st.session_state:
            try:
                st.session_state.auth_flow.fetch_token(code=auth_code)
                new_creds = st.session_state.auth_flow.credentials
                if save_google_creds_to_firebase(safe_key, new_creds):
                    st.success("✅ 구글 캘린더 권한 승인 완료!")
                    st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
                    st.query_params.clear()
                    del st.session_state.auth_flow
                    time.sleep(1)
                    st.rerun()
            except Exception as e:
                st.error(f"⚠️ 인증 처리 실패: {e}")
                st.query_params.clear()
                if 'auth_flow' in st.session_state: del st.session_state.auth_flow
        else:
            # 🚨 흐름이 끊겼을 때: 다시 flow를 만들고 code를 사용해 시도 (PKCE 우회 시도)
            st.warning("⚠️ 인증 세션이 만료되었습니다. 다시 시도해 주세요.")
            st.query_params.clear()

    # (4) Flow 객체 생성 및 유지
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf.get("redirect_uri")
        )

    auth_url, _ = st.session_state.auth_flow.authorization_url(
        prompt='consent', access_type='offline', include_granted_scopes='true'
    )
    
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 구글 계정 연동하기]({auth_url})**")
    return None
