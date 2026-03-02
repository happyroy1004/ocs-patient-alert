import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
import json
import pickle
import time
import os

# Google 인증 라이브러리 (과거 성공 코드 기반)
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# 사용할 권한 범위
SCOPES = ["https://www.googleapis.com/auth/calendar.events"]

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

# --- 2. 경로 변환 (과거 방식 복원) ---
def sanitize_path(email):
    # 과거 코드 방식
    return email.replace(".", "_dot_").replace("@", "_at_")

def recover_email(safe_id):
    # 과거 코드 방식
    return safe_id.replace("_at_", "@").replace("_dot_", ".")

# --- 3. 자격 증명 저장/로드 (과거 방식: 사용자 노드 내부 저장) ---
def save_google_creds_to_firebase(user_id_safe, creds):
    try:
        # 과거 코드 방식: users/{safe_key}/google_creds 경로 사용
        ref = db.reference(f"users/{user_id_safe}/google_creds")
        ref.set({'creds_hex': pickle.dumps(creds).hex()})
        return True
    except Exception as e:
        st.error(f"❌ 저장 실패: {e}")
        return False

def load_google_creds_from_firebase(user_id_safe):
    try:
        # 과거 코드 방식
        ref = db.reference(f"users/{user_id_safe}/google_creds")
        data = ref.get()
        if data and 'creds_hex' in data:
            return pickle.loads(bytes.fromhex(data['creds_hex']))
        return None
    except:
        return None

# --- 4. Google 캘린더 서비스 관리 ---
def get_google_calendar_service(user_id_safe):
    # (A) 세션 상태 확인 (사용자별 키 분리)
    session_key = f"google_creds_{user_id_safe}"
    creds = st.session_state.get(session_key)
    
    # (B) 세션에 없으면 Firebase에서 로드
    if not creds:
        creds = load_google_creds_from_firebase(user_id_safe)
        if creds:
            st.session_state[session_key] = creds

    client_config = {
        "web": {
            "client_id": st.secrets["google_calendar"]["client_id"],
            "client_secret": st.secrets["google_calendar"]["client_secret"],
            "redirect_uris": [st.secrets["google_calendar"]["redirect_uri"]],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token"
        }
    }

    # (C) 인증 정보가 유효한 경우 서비스 반환
    if creds:
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                save_google_creds_to_firebase(user_id_safe, creds)
                st.session_state[session_key] = creds
            except:
                pass
        if creds.valid:
            # 과거 코드 방식
            return build('calendar', 'v3', credentials=creds)

    # (D) 인증 후 돌아온 코드 처리 (과거 코드 로직)
    auth_code = st.query_params.get("code")
    if auth_code and 'auth_flow' in st.session_state:
        try:
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            st.session_state[session_key] = new_creds
            
            if save_google_creds_to_firebase(user_id_safe, new_creds):
                st.success("✅ Google Calendar 인증이 완료되었습니다.")
                st.query_params.clear()
                if 'auth_flow' in st.session_state:
                    del st.session_state.auth_flow
                time.sleep(1) # 저장 안정화를 위한 대기
                st.rerun()
        except Exception as e:
            st.error(f"⚠️ 인증 실패: {e}")
            st.query_params.clear()

    # (E) 인증이 필요한 경우 링크 표시 (과거 코드 방식)
    if 'auth_flow' not in st.session_state:
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
    
    st.warning("📅 Google Calendar 연동을 위해 인증이 필요합니다.")
    st.markdown(f"**[🔗 Google Calendar 인증 링크]({auth_url})**")
    return None
