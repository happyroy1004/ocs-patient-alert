import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
import pickle
import time
import json
import os

# Google 인증 관련 라이브러리 (과거 성공 방식 )
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# 사용할 권한 범위 (과거 코드 기준 )
SCOPES = ["https://www.googleapis.com/auth/calendar.events"]

# --- 1. Firebase 초기화 (ui_manager가 기대하는 구조 ) ---
@st.cache_resource
def get_db_refs():
    if not firebase_admin._apps:
        try:
            # secrets.toml의 [firebase] 섹션 데이터 로드
            firebase_conf = st.secrets["firebase"]
            
            # 서비스 계정 JSON 문자열인 경우와 딕셔너리인 경우 모두 대응
            if "FIREBASE_SERVICE_ACCOUNT_JSON" in firebase_conf:
                creds_dict = json.loads(firebase_conf["FIREBASE_SERVICE_ACCOUNT_JSON"])
            else:
                creds_dict = dict(firebase_conf)
                
            db_url = firebase_conf["database_url"]
            cred = credentials.Certificate(creds_dict)
            firebase_admin.initialize_app(cred, {'databaseURL': db_url})
        except Exception as e:
            st.error(f"❌ Firebase 초기화 실패: {e}")
            return None, None, None
            
    base_ref = db.reference()
    # ui_manager.py가 언패킹(unpacking)하여 사용하는 3개 값 반환 
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

# --- 2. 경로 변환 함수 (과거 방식 복원 ) ---
def sanitize_path(email):
    """이메일을 Firebase 키로 변환 (. -> _dot_, @ -> _at_)"""
    return email.replace(".", "_dot_").replace("@", "_at_")

def recover_email(safe_id):
    """Firebase 키를 다시 이메일로 복구"""
    return safe_id.replace("_at_", "@").replace("_dot_", ".")

# --- 3. 자격 증명 저장/로드 (과거 성공 경로: users/{id}/google_creds ) ---
def save_google_creds_to_firebase(user_id_safe, creds):
    try:
        ref = db.reference(f"users/{user_id_safe}/google_creds")
        # 과거 코드 스타일로 개별 필드 저장 대신 안정적인 hex 직렬화 사용
        ref.set({'creds_hex': pickle.dumps(creds).hex()})
        return True
    except Exception as e:
        st.error(f"❌ DB 저장 실패 (규칙 확인 필요): {e}")
        return False

def load_google_creds_from_firebase(user_id_safe):
    try:
        ref = db.reference(f"users/{user_id_safe}/google_creds")
        data = ref.get()
        if data and 'creds_hex' in data:
            return pickle.loads(bytes.fromhex(data['creds_hex']))
        return None
    except:
        return None

# --- 4. Google 캘린더 서비스 로직 (과거 세션 관리 방식 ) ---
def get_google_calendar_service(user_id_safe):
    # 사용자별 고유 세션 키 사용
    session_key = f"google_creds_{user_id_safe}"
    creds = st.session_state.get(session_key)
    
    if not creds:
        creds = load_google_creds_from_firebase(user_id_safe)
        if creds:
            st.session_state[session_key] = creds

    # secrets.toml의 [google_calendar] 섹션 로드
    conf = st.secrets["google_calendar"]
    client_config = {
        "web": {
            "client_id": conf["client_id"],
            "client_secret": conf["client_secret"],
            "redirect_uris": [conf["redirect_uri"]],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token"
        }
    }

    if creds:
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                save_google_creds_to_firebase(user_id_safe, creds)
                st.session_state[session_key] = creds
            except:
                pass
        if creds.valid:
            return build('calendar', 'v3', credentials=creds)

    # 리디렉션 코드(Code) 처리 
    auth_code = st.query_params.get("code")
    if auth_code and 'auth_flow' in st.session_state:
        try:
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            st.session_state[session_key] = new_creds
            if save_google_creds_to_firebase(user_id_safe, new_creds):
                st.success("✅ 인증 성공! 정보가 저장되었습니다.")
                st.query_params.clear()
                if 'auth_flow' in st.session_state:
                    del st.session_state.auth_flow
                time.sleep(1)
                st.rerun()
        except Exception:
            st.query_params.clear()

    # 인증 흐름 생성 (PKCE Verifier 유지용 세션 고정)
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf["redirect_uri"]
        )

    auth_url, _ = st.session_state.auth_flow.authorization_url(
        prompt='consent', access_type='offline', include_granted_scopes='true'
    )
    
    st.warning("📅 Google Calendar 연동을 위해 인증이 필요합니다.")
    st.markdown(f"**[🔗 Google Calendar 인증 링크]({auth_url})**")
    return None
