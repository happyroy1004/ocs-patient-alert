import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle

# 필수 스코프 (이메일 확인 포함)
SCOPES = [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid'
]

# --- 1. DB 초기화 ---
@st.cache_resource
def get_db_refs():
    if not firebase_admin._apps:
        creds_dict = dict(st.secrets["firebase"])
        db_url = st.secrets["database_url"]
        if 'FIREBASE_DATABASE_URL' in creds_dict: del creds_dict['FIREBASE_DATABASE_URL']
        cred = credentials.Certificate(creds_dict)
        firebase_admin.initialize_app(cred, {'databaseURL': db_url})
    
    base_ref = db.reference()
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

# --- 2. Credentials 저장 함수 (디버깅 메시지 추가) ---
def save_google_creds_to_firebase(safe_key, creds):
    try:
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        encoded = pickle.dumps(creds).hex()
        ref.set({'creds': encoded})
        st.write(f"DEBUG: DB 저장 성공! (Key: {safe_key})") # 저장 확인용
    except Exception as e:
        st.error(f"DB 저장 중 오류 발생: {e}")

def load_google_creds_from_firebase(safe_key):
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    if data and 'creds' in data:
        return pickle.loads(bytes.fromhex(data['creds']))
    return None

# --- 3. Google Calendar Service (최종 끝판왕 버전) ---

def get_google_calendar_service(safe_key):
    # 세션에 서비스가 있으면 즉시 반환
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

    # 설정 로드
    conf = dict(st.secrets["google_calendar"])
    redirect_uri = conf.get("redirect_uri")
    
    # [핵심] Client Config 구조를 구글이 요구하는 'web' 형식을 명시적으로 생성
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

    # Flow 객체 생성 및 세션 보관 (PKCE 유지를 위해 필수)
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=redirect_uri
        )

    # 1. URL에 코드가 들어왔을 때 처리 (가장 높은 우선순위)
    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            # 토큰 교환 실행
            st.session_state.auth_flow.fetch_token(code=auth_code)
            creds = st.session_state.auth_flow.credentials
            
            # DB 저장 (OCS 계정 키로 저장)
            save_google_creds_to_firebase(safe_key, creds)
            
            # 서비스 빌드 및 세션 저장
            service = build('calendar', 'v3', credentials=creds)
            st.session_state.google_calendar_service = service
            
            # 청소 및 리런
            st.query_params.clear()
            if 'auth_flow' in st.session_state: del st.session_state.auth_flow
            st.success("인증 및 DB 저장 완료!")
            st.rerun()
            
        except Exception as e:
            # 여기서 상세 에러를 찍어서 'Bad Request'의 진짜 이유를 확인
            st.error(f"⚠️ 토큰 교환 실패 상세: {str(e)}")
            st.query_params.clear()
            return None

    # 2. 기존 DB 로드 시도
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

    # 3. 인증 링크 표시 (인증 안 된 경우)
    auth_url, _ = st.session_state.auth_flow.authorization_url(
        prompt='consent', access_type='offline'
    )
    st.info("구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 Google Calendar 인증 링크]({auth_url})**")
    return None
