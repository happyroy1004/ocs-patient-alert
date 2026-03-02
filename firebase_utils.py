import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
import time
import json

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
            if 'FIREBASE_DATABASE_URL' in creds_dict: 
                del creds_dict['FIREBASE_DATABASE_URL']
            cred = credentials.Certificate(creds_dict)
            firebase_admin.initialize_app(cred, {'databaseURL': db_url})
        except Exception as e:
            st.error(f"❌ Firebase 초기화 실패: {e}")
            return None, None, None
    base_ref = db.reference()
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

# --- 2. 유틸리티 함수 ---
def sanitize_path(email):
    return email.replace('.', '_')

# --- 3. 핵심: 자격 증명 저장 및 로드 (안정성 강화) ---
def save_google_creds_to_firebase(safe_key, creds):
    """DB 저장이 완료될 때까지 확실히 확인"""
    try:
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        ref.set({'creds': pickle.dumps(creds).hex()})
        # 데이터가 잘 들어갔는지 다시 한번 확인 (Verification)
        check = ref.get()
        return check is not None
    except Exception as e:
        st.error(f"❌ DB 저장 오류: {e}")
        return False

def load_google_creds_from_firebase(safe_key):
    try:
        data = db.reference(f'google_calendar_creds/{safe_key}').get()
        if data and 'creds' in data:
            return pickle.loads(bytes.fromhex(data['creds']))
    except:
        return None
    return None

# --- 4. 메인 서비스 로직 ---
def get_google_calendar_service(safe_key):
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 기존 토큰 로드 시도
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

    # OAuth 설정 및 Flow 생성
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

    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf.get("redirect_uri")
        )

    # 💡 해결책: 인증 코드 처리 시 명시적 대기 및 확인 과정 추가
    auth_code = st.query_params.get("code")
    if auth_code and 'auth_flow' in st.session_state:
        try:
            with st.spinner("🔄 구글 인증 정보를 저장 중입니다..."):
                st.session_state.auth_flow.fetch_token(code=auth_code)
                new_creds = st.session_state.auth_flow.credentials
                
                # DB 저장을 시도하고 성공할 때까지 대기
                success = save_google_creds_to_firebase(safe_key, new_creds)
                
                if success:
                    st.success("✅ 권한 승인 완료! 정보를 안전하게 저장했습니다.")
                    st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
                    st.query_params.clear()
                    if 'auth_flow' in st.session_state: del st.session_state.auth_flow
                    time.sleep(2) # ⏱️ 리디렉션 전 DB 쓰기 완료를 위한 시간 벌기
                    st.rerun()
                else:
                    st.error("❌ 정보를 저장하지 못했습니다. 다시 시도해 주세요.")
        except Exception as e:
            st.error(f"⚠️ 인증 처리 중 오류: {e}")
            st.query_params.clear()
            if 'auth_flow' in st.session_state: del st.session_state.auth_flow

    # 새 인증 링크 생성
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
