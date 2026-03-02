import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
import json

# --- [설정] 권한 범위 고정 (ImportError 방지) ---
SCOPES = [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid'
]

# --- 1. Firebase 초기화 ---
@st.cache_resource
def get_db_refs():
    """Firebase Admin SDK 초기화 및 참조 반환"""
    if not firebase_admin._apps:
        try:
            # st.secrets 구조가 정확한지 확인하며 로드
            if "firebase" not in st.secrets or "database_url" not in st.secrets:
                st.error("🚨 Secrets.toml에 'firebase' 섹션이나 'database_url'이 누락되었습니다.")
                return None, None, None
                
            creds_dict = dict(st.secrets["firebase"])
            db_url = st.secrets["database_url"]
            
            # 인증 정보에서 불필요한 키 제거 (SDK 충돌 방지)
            if 'FIREBASE_DATABASE_URL' in creds_dict: 
                del creds_dict['FIREBASE_DATABASE_URL']
            
            cred = credentials.Certificate(creds_dict)
            firebase_admin.initialize_app(cred, {'databaseURL': db_url})
        except Exception as e:
            st.error(f"❌ Firebase 초기화 실패: {e}")
            return None, None, None

    base_ref = db.reference()
    # 기존 코드와의 호환성을 위해 3개의 객체 반환
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

# --- 2. Google Credentials 저장 및 로드 ---
def save_google_creds_to_firebase(safe_key, creds):
    """구글 인증 정보를 Firebase에 저장"""
    try:
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        encoded_creds = pickle.dumps(creds).hex()
        ref.set({'creds': encoded_creds})
        st.success(f"✅ 인증 정보가 DB에 저장되었습니다! (ID: {safe_key})")
    except Exception as e:
        st.error(f"❌ DB 저장 실패: {e}")

def load_google_creds_from_firebase(safe_key):
    """Firebase에서 구글 인증 정보 로드"""
    try:
        data = db.reference(f'google_calendar_creds/{safe_key}').get()
        if data and 'creds' in data:
            return pickle.loads(bytes.fromhex(data['creds']))
    except:
        return None
    return None

# --- 3. Google Calendar Service 핵심 로직 ---

def get_google_calendar_service(safe_key):
    """구글 캘린더 서비스 객체 생성 및 인증 흐름 관리"""
    
    # 세션에 서비스가 있다면 즉시 반환
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # Secrets에서 구글 설정 로드
    if "google_calendar" not in st.secrets:
        st.error("🚨 Secrets.toml에 [google_calendar] 섹션이 없습니다.")
        return None
        
    conf = dict(st.secrets["google_calendar"])
    redirect_uri = conf.get("redirect_uri")
    
    client_config = {
        "web": {
            "client_id": conf.get("client_id"),
            "project_id": conf.get("project_id"),
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "client_secret": conf.get("client_secret"),
            "redirect_uris": [redirect_uri]
        }
    }

    # Flow 객체 생성 및 세션 보관 (PKCE 오류 방지)
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, 
            scopes=SCOPES, 
            redirect_uri=redirect_uri
        )

    # URL 파라미터(코드) 처리
    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            
            # DB 저장
            save_google_creds_to_firebase(safe_key, new_creds)
            
            # 서비스 빌드 및 세션 저장
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
            
            # 정리 및 리셋
            st.query_params.clear()
            if 'auth_flow' in st.session_state:
                del st.session_state.auth_flow
            
            st.rerun()
            
        except Exception as e:
            st.error(f"⚠️ 인증 처리 중 오류 발생: {str(e)}")
            st.query_params.clear()
            return None

    # 기존 데이터 로드 및 갱신
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
            except:
                creds = None

    # 인증되지 않은 경우 링크 표시
    auth_url, _ = st.session_state.auth_flow.authorization_url(
        prompt='consent', 
        access_type='offline'
    )
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 구글 계정 인증 링크]({auth_url})**")
    
    return None

def sanitize_path(email):
    return email.replace('.', '_')
