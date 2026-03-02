import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
import json
import pickle
import time
import os

# Google 인증 관련 라이브러리 (정확한 import 확인)
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# 사용할 권한 범위 (과거 코드 기준)
SCOPES = ["https://www.googleapis.com/auth/calendar.events"] [cite: 21]

# --- 1. Firebase 초기화 ---
@st.cache_resource
def get_db_refs():
    if not firebase_admin._apps:
        try:
            creds_dict = dict(st.secrets["firebase"]) [cite: 1, 2]
            db_url = st.secrets["database_url"] [cite: 1]
            cred = credentials.Certificate(creds_dict) [cite: 1]
            firebase_admin.initialize_app(cred, {'databaseURL': db_url}) [cite: 1]
        except Exception as e:
            st.error(f"❌ Firebase 초기화 실패: {e}")
            return None, None, None
    base_ref = db.reference()
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

# --- 2. 경로 변환 함수 (과거 방식 복원) ---
def sanitize_path(email):
    # 과거 코드의 변환 방식 적용 [cite: 3]
    return email.replace(".", "_dot_").replace("@", "_at_")

def recover_email(safe_id):
    # 과거 코드의 복구 방식 적용 [cite: 3]
    return safe_id.replace("_at_", "@").replace("_dot_", ".")

# --- 3. 자격 증명 저장/로드 (과거 로직 복원) ---
def save_google_creds_to_firebase(user_id_safe, creds):
    try:
        # 과거 코드처럼 users 노드 아래 google_creds에 저장 
        ref = db.reference(f"users/{user_id_safe}/google_creds")
        # 객체 전체를 hex로 직렬화하여 저장
        ref.set({'creds_hex': pickle.dumps(creds).hex()})
        return True
    except Exception as e:
        st.error(f"❌ Credential 저장 실패: {e}") [cite: 4]
        return False

def load_google_creds_from_firebase(user_id_safe):
    try:
        ref = db.reference(f"users/{user_id_safe}/google_creds") [cite: 5]
        data = ref.get() [cite: 5]
        if data and 'creds_hex' in data:
            return pickle.loads(bytes.fromhex(data['creds_hex'])) [cite: 7]
        return None
    except:
        return None

# --- 4. 핵심 서비스 로직 (과거 세션 관리 방식 + NameError 방지) ---
def get_google_calendar_service(user_id_safe):
    # (A) 세션 상태에서 사용자별 인증 정보 확인 
    session_key = f"google_creds_{user_id_safe}"
    creds = st.session_state.get(session_key)
    
    # (B) 세션에 없으면 Firebase에서 로드 
    if not creds:
        creds = load_google_creds_from_firebase(user_id_safe)
        if creds:
            st.session_state[session_key] = creds

    # Google 클라이언트 설정 (secrets.toml 기반) [cite: 22]
    client_config = {
        "web": {
            "client_id": st.secrets["google_calendar"]["client_id"], [cite: 22]
            "client_secret": st.secrets["google_calendar"]["client_secret"], [cite: 22]
            "redirect_uris": [st.secrets["google_calendar"]["redirect_uri"]], [cite: 22]
            "auth_uri": "https://accounts.google.com/o/oauth2/auth", [cite: 22]
            "token_uri": "https://oauth2.googleapis.com/token" [cite: 22]
        }
    }

    # (C) 이미 유효한 인증 정보가 있는 경우 처리 [cite: 26]
    if creds:
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                save_google_creds_to_firebase(user_id_safe, creds)
                st.session_state[session_key] = creds
            except:
                pass
        
        if creds.valid:
            return build('calendar', 'v3', credentials=creds) [cite: 26]

    # (D) 인증 코드가 URL에 포함되어 돌아온 경우 (Redirect 처리) [cite: 23, 24]
    auth_code = st.query_params.get("code")
    if auth_code and 'auth_flow' in st.session_state:
        try:
            # 세션에 고정된 flow 객체로 토큰 교환 
            st.session_state.auth_flow.fetch_token(code=auth_code)
            creds = st.session_state.auth_flow.credentials
            
            # 성공 시 세션 및 Firebase 저장 
            st.session_state[session_key] = creds
            if save_google_creds_to_firebase(user_id_safe, creds):
                st.success("✅ Google Calendar 인증이 완료되었습니다.") [cite: 24]
                st.query_params.clear()
                if 'auth_flow' in st.session_state: 
                    del st.session_state.auth_flow
                time.sleep(1)
                st.rerun()
        except Exception as e:
            st.error(f"⚠️ 인증 처리 오류: {e}")
            st.query_params.clear()

    # (E) 인증이 필요한 경우 (Flow 생성 및 링크 표시) [cite: 25]
    if 'auth_flow' not in st.session_state:
        # 여기서 Flow 클래스를 명확히 사용 (NameError 방지)
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, 
            scopes=SCOPES, 
            redirect_uri=st.secrets["google_calendar"]["redirect_uri"]
        )

    auth_url, _ = st.session_state.auth_flow.authorization_url(
        prompt='consent', 
        access_type='offline',
        include_granted_scopes='true'
    )
    
    st.warning("📅 Google Calendar 연동을 위해 인증이 필요합니다.") [cite: 25]
    st.markdown(f"**[🔗 Google Calendar 인증 링크]({auth_url})**") [cite: 25]
    return None
