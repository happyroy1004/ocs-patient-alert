import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
import pickle
import time
import os

# [중요] ImportError 방지를 위한 라이브러리 강제 로드 및 예외 처리
try:
    from google_auth_oauthlib.flow import Flow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
except ImportError as e:
    st.error(f"❌ 필수 라이브러리 로드 실패: {e}. requirements.txt에 'google-auth-oauthlib'와 'google-api-python-client'가 있는지 확인하세요.")
    st.stop()

# 권한 범위 설정 (동의 화면 범위와 일치해야 함)
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
            # Secrets에서 설정을 가져올 때 dict로 명시적 변환
            creds_dict = dict(st.secrets["firebase"])
            db_url = st.secrets["database_url"]
            
            # 불필요한 키 제거 (KeyError 방지)
            if 'FIREBASE_DATABASE_URL' in creds_dict: 
                del creds_dict['FIREBASE_DATABASE_URL']
                
            cred = credentials.Certificate(creds_dict)
            firebase_admin.initialize_app(cred, {'databaseURL': db_url})
        except Exception as e:
            st.error(f"❌ Firebase 초기화 실패 (Secrets 확인 필요): {e}")
            return None, None, None
            
    base_ref = db.reference()
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

# --- 2. 유틸리티 함수 ---
def sanitize_path(email):
    """Firebase 경로에 사용할 수 없는 문자 치환"""
    return email.replace('.', '_')

def recover_email(safe_key):
    """치환된 키를 다시 이메일로 복구"""
    return safe_key.replace('_', '.')

# --- 3. Google 인증 관리 (PKCE 'Missing code verifier' 및 저장 문제 해결) ---

def save_google_creds_to_firebase(safe_key, creds):
    """인증된 자격 증명을 Firebase에 헥사 문자열로 직렬화하여 저장"""
    try:
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        ref.set({'creds': pickle.dumps(creds).hex()})
        return True
    except Exception as e:
        st.error(f"❌ DB 저장 실패: {e}")
        return False

def load_google_creds_from_firebase(safe_key):
    """Firebase에서 직렬화된 자격 증명을 불러와 역직렬화"""
    try:
        data = db.reference(f'google_calendar_creds/{safe_key}').get()
        if data and 'creds' in data:
            return pickle.loads(bytes.fromhex(data['creds']))
    except:
        return None
    return None

def get_google_calendar_service(safe_key):
    """Google Calendar 서비스 빌드 및 OAuth2 인증 흐름 제어"""
    
    # 1. 세션에 이미 서비스 객체가 있다면 즉시 반환
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 2. DB에서 기존 자격 증명 로드 시도
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
                pass # 갱신 실패 시 재인증 진행

    # 3. OAuth 클라이언트 설정 로드
    if "google_calendar" not in st.secrets:
        st.error("🚨 Secrets에 [google_calendar] 설정이 누락되었습니다.")
        return None

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

    # 4. [핵심] Flow 객체를 세션에 고정하여 PKCE 보안 키 유실 방지
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf.get("redirect_uri")
        )

    # 5. URL에서 인증 코드(code) 처리
    auth_code = st.query_params.get("code")
    if auth_code and 'auth_flow' in st.session_state:
        try:
            # 세션에 보관된 Flow 객체로 토큰 교환
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            
            # DB 저장 및 성공 알림
            if save_google_creds_to_firebase(safe_key, new_creds):
                st.success("✅ 구글 캘린더 권한 승인 및 저장이 완료되었습니다!")
                st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
                
                # 파라미터 정리 및 세션 초기화 후 재실행
                st.query_params.clear()
                if 'auth_flow' in st.session_state:
                    del st.session_state.auth_flow
                time.sleep(2) # 성공 메시지 확인용
                st.rerun()
        except Exception as e:
            st.error(f"⚠️ 인증 처리 중 오류: {e}")
            st.query_params.clear()
            if 'auth_flow' in st.session_state:
                del st.session_state.auth_flow

    # 6. 인증이 필요한 경우 링크 표시 (새 Flow 객체 생성)
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
