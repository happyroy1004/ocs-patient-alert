import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
import json

# --- [설정] 권한 범위 (이메일 확인 권한 포함) ---
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
            creds_dict = dict(st.secrets["firebase"])
            db_url = st.secrets["database_url"]
            # 인증 정보에서 불필요한 키 제거
            if 'FIREBASE_DATABASE_URL' in creds_dict: 
                del creds_dict['FIREBASE_DATABASE_URL']
            
            cred = credentials.Certificate(creds_dict)
            firebase_admin.initialize_app(cred, {'databaseURL': db_url})
        except Exception as e:
            st.error(f"❌ Firebase 초기화 실패: {e}")
            return None, None, None

    base_ref = db.reference()
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

# --- 2. Google Credentials 저장 및 로드 ---
def save_google_creds_to_firebase(safe_key, creds):
    """구글 인증 정보를 Firebase에 저장 (성공/실패 화면 표시)"""
    try:
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        # Credentials 객체를 pickle로 직렬화 후 hex 문자열로 변환
        encoded_creds = pickle.dumps(creds).hex()
        ref.set({'creds': encoded_creds})
        st.success(f"✅ DB 저장 완료! 경로: google_calendar_creds/{safe_key}")
    except Exception as e:
        st.error(f"❌ DB 저장 실패: {e}. Firebase Rules를 확인하세요.")

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
    
    # A. 이미 세션에 서비스가 있다면 즉시 반환
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # B. Secrets 로드 및 Config 설정
    try:
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
    except Exception as e:
        st.error(f"🚨 Secrets 로드 실패: {e}")
        return None

    # C. Flow 객체 생성 및 세션 보관 (PKCE 오류 방지 핵심)
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, 
            scopes=SCOPES, 
            redirect_uri=redirect_uri
        )

    # D. 리다이렉트 후 인증 코드 처리 (코드가 URL에 있을 때)
    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            # 1. 토큰 교환
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            
            # 2. DB 저장 시도 (현재 로그인된 OCS 계정 키 사용)
            save_google_creds_to_firebase(safe_key, new_creds)
            
            # 3. 서비스 빌드 및 세션 저장
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
            
            # 4. 정리 및 리셋
            st.query_params.clear()
            if 'auth_flow' in st.session_state:
                del st.session_state.auth_flow
            
            st.rerun() # 주소창 청소 및 화면 갱신
            
        except Exception as e:
            st.error(f"⚠️ 인증 처리 중 오류 발생: {str(e)}")
            # 상세 에러 분석 가이드
            if "invalid_grant" in str(e):
                st.warning("인증 코드가 만료되었거나 테스트 사용자 등록이 안 된 계정입니다.")
            st.query_params.clear()
            return None

    # E. 기존 DB 데이터 로드 및 유효성 확인
    creds = load_google_creds_from_firebase(safe_key)
    if creds:
        if creds.valid:
            service = build('calendar', 'v3', credentials=creds)
            st.session_state.google_calendar_service = service
            return service
        elif creds.refresh_token:
            try:
                # 만료된 경우 자동 갱신
                creds.refresh(Request())
                save_google_creds_to_firebase(safe_key, creds)
                service = build('calendar', 'v3', credentials=creds)
                st.session_state.google_calendar_service = service
                return service
            except Exception as e:
                st.warning(f"토큰 갱신 실패: {e}")
                creds = None

    # F. 인증되지 않은 경우: 인증 링크 표시
    auth_url, _ = st.session_state.auth_flow.authorization_url(
        prompt='consent', 
        access_type='offline',
        include_granted_scopes='true'
    )
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 구글 계정 권한 허용 링크]({auth_url})**")
    st.caption("주의: 반드시 구글 콘솔에 등록된 '테스트 사용자' 계정으로 로그인해야 합니다.")
    
    return None

def sanitize_path(email):
    """이메일을 Firebase 키 형식으로 변환"""
    return email.replace('.', '_')
