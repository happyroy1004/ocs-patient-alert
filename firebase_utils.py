import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle

# 필수 권한 설정
SCOPES = ['https://www.googleapis.com/auth/calendar']

# --- 0. Secrets 및 환경 설정 ---
try:
    FIREBASE_CREDENTIALS = dict(st.secrets["firebase"]) 
    DB_URL = st.secrets["database_url"] 
    GOOGLE_CALENDAR_CLIENT_SECRET = dict(st.secrets["google_calendar"])
except Exception as e:
    st.error(f"🚨 Secrets.toml 설정 오류: {e}")
    GOOGLE_CALENDAR_CLIENT_SECRET = {}

# --- 1. Firebase 초기화 ---
@st.cache_resource
def get_db_refs():
    if not firebase_admin._apps:
        try:
            creds_init = FIREBASE_CREDENTIALS.copy()
            if 'FIREBASE_DATABASE_URL' in creds_init: del creds_init['FIREBASE_DATABASE_URL']
            cred = credentials.Certificate(creds_init)
            firebase_admin.initialize_app(cred, {'databaseURL': DB_URL})
        except Exception as e:
            st.error(f"❌ Firebase 초기화 실패: {e}")
            return None, None, None
    
    base_ref = db.reference()
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

# --- 2. Credentials 관리 ---
def save_google_creds_to_firebase(safe_key, creds):
    """DB에 직접 저장 시도하고 결과를 화면에 출력 (디버깅용)"""
    try:
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        encoded_creds = pickle.dumps(creds).hex()
        ref.set({'creds': encoded_creds})
        st.toast(f"✅ DB 저장 성공: {safe_key}")
    except Exception as e:
        st.error(f"❌ DB 저장 실패: {e}")

def load_google_creds_from_firebase(safe_key):
    try:
        data = db.reference(f'google_calendar_creds/{safe_key}').get()
        if data and 'creds' in data:
            return pickle.loads(bytes.fromhex(data['creds']))
    except:
        return None
    return None

# --- 3. Google Calendar Service (인증 로직 집중 수정) ---

def get_google_calendar_service(safe_key):
    # 1. 이미 세션에 서비스가 있다면 즉시 반환
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 2. Redirect URI 및 Config 설정 (표준화)
    redirect_uri = GOOGLE_CALENDAR_CLIENT_SECRET.get("redirect_uri")
    client_config = {
        "web": {
            "client_id": GOOGLE_CALENDAR_CLIENT_SECRET.get("client_id"),
            "project_id": GOOGLE_CALENDAR_CLIENT_SECRET.get("project_id"),
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "client_secret": GOOGLE_CALENDAR_CLIENT_SECRET.get("client_secret"),
            "redirect_uris": [redirect_uri]
        }
    }

    # Flow 객체를 세션에 고정 (PKCE Verifier 유지용)
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=redirect_uri
        )

    # 3. URL 파라미터(코드) 확인 및 처리
    auth_code = st.query_params.get("code")
    
    if auth_code:
        try:
            # [가장 중요한 단계] 토큰 교환
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            
            # DB 저장 강제 실행
            save_google_creds_to_firebase(safe_key, new_creds)
            
            # 서비스 빌드
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
            
            # 청소 및 리셋
            st.query_params.clear()
            if 'auth_flow' in st.session_state: del st.session_state.auth_flow
            
            st.success("🎉 구글 인증 및 계정 연동에 성공했습니다!")
            st.rerun()
            
        except Exception as e:
            # 에러 발생 시 상세 정보 출력
            st.error(f"⚠️ 인증 처리 오류: {str(e)}")
            st.query_params.clear()
            if 'auth_flow' in st.session_state: del st.session_state.auth_flow
            return None

    # 4. 기존 DB 데이터 로드 시도
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
                pass

    # 5. 인증되지 않은 경우 링크 표시
    auth_url, _ = st.session_state.auth_flow.authorization_url(
        prompt='consent', access_type='offline', include_granted_scopes='true'
    )
    st.info("구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 Google Calendar 인증 링크]({auth_url})**")
    return None

def sanitize_path(email):
    return email.replace('.', '_')
