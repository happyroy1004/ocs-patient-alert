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

# --- 2. Google 인증 관리 (저장 로직 강화) ---
def save_google_creds_to_firebase(safe_key, creds):
    """자격 증명을 Firebase에 확실히 저장하고 성공 여부 반환"""
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

def get_google_calendar_service(safe_key):
    # 1. 이미 세션에 로드된 경우 즉시 반환
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 2. DB에서 기존 토큰 로드 시도
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

    # 3. OAuth 설정 로드
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

    # 🔑 Flow 객체 세션 고정
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf.get("redirect_uri")
        )

    # 4. 인증 완료 후 돌아왔을 때 처리
    auth_code = st.query_params.get("code")
    if auth_code and 'auth_flow' in st.session_state:
        try:
            # 토큰 교환
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            
            # ✅ DB 저장 및 성공 메시지 출력
            if save_google_creds_to_firebase(safe_key, new_creds):
                st.success("✅ 구글 캘린더 권한이 승인되었습니다! 잠시 후 화면이 갱신됩니다.")
                st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
                
                # 시각적 확인을 위해 짧게 대기 후 리프레시
                st.query_params.clear()
                if 'auth_flow' in st.session_state: del st.session_state.auth_flow
                time.sleep(1.5)
                st.rerun()
        except Exception as e:
            st.error(f"⚠️ 인증 처리 중 오류 발생: {e}")
            st.query_params.clear()
            if 'auth_flow' in st.session_state: del st.session_state.auth_flow

    # 5. 인증 링크 표시 (새 Flow 객체 필요 시)
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
