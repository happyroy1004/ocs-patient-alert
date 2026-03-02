import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
import pickle
import time
import json
import os

# Google 인증 관련 라이브러리 (과거 성공 방식)
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# 사용할 권한 범위 (과거 코드 기준)
SCOPES = ["https://www.googleapis.com/auth/calendar.events"]

# --- 1. Firebase 초기화 (ui_manager가 호출함) ---
@st.cache_resource
def get_db_refs():
    """Firebase DB 참조 객체를 초기화하고 반환합니다."""
    if not firebase_admin._apps:
        try:
            # [span_8](start_span)secrets.toml의 [firebase] 섹션 구조 사용[span_8](end_span)
            creds_dict = json.loads(st.secrets["firebase"]["FIREBASE_SERVICE_ACCOUNT_JSON"])
            db_url = st.secrets["firebase"]["database_url"]
            cred = credentials.Certificate(creds_dict)
            firebase_admin.initialize_app(cred, {'databaseURL': db_url})
        except Exception as e:
            st.error(f"❌ Firebase 초기화 실패: {e}")
            return None, None, None
    base_ref = db.reference()
    # [span_9](start_span)ui_manager가 기대하는 3개의 반환값[span_9](end_span)
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

# --- 2. 경로 변환 함수 (ui_manager가 호출함) ---
def sanitize_path(email):
    [span_10](start_span)"""이메일을 Firebase 키로 변환[span_10](end_span)"""
    return email.replace(".", "_dot_").replace("@", "_at_")

def recover_email(safe_id):
    [span_11](start_span)"""Firebase 키를 다시 이메일로 복구[span_11](end_span)"""
    return safe_id.replace("_at_", "@").replace("_dot_", ".")

# --- 3. 자격 증명 저장/로드 (과거 성공 경로: users/...) ---
def save_google_creds_to_firebase(user_id_safe, creds):
    [span_12](start_span)"""사용자 노드 내부에 인증 정보를 저장[span_12](end_span)"""
    try:
        ref = db.reference(f"users/{user_id_safe}/google_creds")
        ref.set({'creds_hex': pickle.dumps(creds).hex()})
        return True
    except Exception as e:
        st.error(f"❌ 저장 실패: {e}")
        return False

def load_google_creds_from_firebase(user_id_safe):
    [span_13](start_span)"""Firebase에서 인증 정보를 로드[span_13](end_span)"""
    try:
        ref = db.reference(f"users/{user_id_safe}/google_creds")
        data = ref.get()
        if data and 'creds_hex' in data:
            return pickle.loads(bytes.fromhex(data['creds_hex']))
        return None
    except:
        return None

# --- 4. Google 캘린더 서비스 로직 (핵심) ---
def get_google_calendar_service(user_id_safe):
    [span_14](start_span)"""인증 흐름을 관리하고 서비스 객체를 반환[span_14](end_span)"""
    session_key = f"google_creds_{user_id_safe}"
    creds = st.session_state.get(session_key)
    
    if not creds:
        creds = load_google_creds_from_firebase(user_id_safe)
        if creds:
            st.session_state[session_key] = creds

    # [span_15](start_span)secrets.toml 데이터 로드[span_15](end_span)
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

    # [span_16](start_span)리디렉션 처리 로직[span_16](end_span)
    auth_code = st.query_params.get("code")
    if auth_code and 'auth_flow' in st.session_state:
        try:
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            st.session_state[session_key] = new_creds
            if save_google_creds_to_firebase(user_id_safe, new_creds):
                st.success("✅ 인증 성공!")
                st.query_params.clear()
                if 'auth_flow' in st.session_state:
                    del st.session_state.auth_flow
                time.sleep(1)
                st.rerun()
        except:
            st.query_params.clear()

    # [span_17](start_span)인증 흐름 생성 및 링크 표시[span_17](end_span)
    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf["redirect_uri"]
        )

    auth_url, _ = st.session_state.auth_flow.authorization_url(prompt='consent', access_type='offline')
    
    st.warning("📅 Google Calendar 연동을 위해 인증이 필요합니다.")
    st.markdown(f"**[🔗 Google Calendar 인증 링크]({auth_url})**")
    return None
