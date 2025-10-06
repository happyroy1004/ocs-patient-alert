# firebase_utils.py

import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
import json
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from config import SCOPES

# --- Firebase 초기화 및 전역 레퍼런스 ---
users_ref = None
doctor_users_ref = None

def init_firebase():
    """Firebase SDK를 초기화하고 전역 DB 레퍼런스를 설정합니다."""
    global users_ref, doctor_users_ref
    
    if not firebase_admin._apps:
        try:
            firebase_credentials_json_str = st.secrets["firebase"]["FIREBASE_SERVICE_ACCOUNT_JSON"]
            firebase_credentials_dict = json.loads(firebase_credentials_json_str)

            cred = credentials.Certificate(firebase_credentials_dict)
            firebase_admin.initialize_app(cred, {
                'databaseURL': st.secrets["firebase"]["database_url"]
            })
            
            # 전역 레퍼런스 설정
            users_ref = db.reference("users")
            doctor_users_ref = db.reference("doctor_users")

        except Exception as e:
            st.error(f"Firebase 초기화 오류: {e}")
            st.info("secrets.toml 파일의 Firebase 설정을 확인해주세요.")
            st.stop()
            
def get_db_refs():
    """초기화된 DB 레퍼런스를 반환합니다."""
    if users_ref is None or doctor_users_ref is None:
        init_firebase()
    return users_ref, doctor_users_ref, db.reference

# --- 경로 변환 및 복원 ---

def sanitize_path(email):
    """이메일을 Firebase 키로 사용하기 위해 안전한 경로로 변환합니다."""
    return email.replace(".", "_dot_").replace("@", "_at_").replace("-", "_dash_")

def recover_email(safe_id: str) -> str:
    """Firebase 안전 키에서 원래 이메일로 복원합니다."""
    return safe_id.replace("_at_", "@").replace("_dot_", ".").replace("_dash_", "-")

# --- Google Calendar Creds 관리 ---

def save_google_creds_to_firebase(user_id_safe, creds):
    """Google Calendar 인증 정보를 Firebase에 저장합니다."""
    try:
        ref = db.reference(f"users/{user_id_safe}/google_creds") if 'doctor' not in user_id_safe else db.reference(f"doctor_users/{user_id_safe}/google_creds")
        ref.set({
            'token': creds.token, 'refresh_token': creds.refresh_token, 'token_uri': creds.token_uri,
            'client_id': creds.client_id, 'client_secret': creds.client_secret, 'scopes': creds.scopes
        })
        return True
    except Exception as e:
        st.error(f"Failed to save Google credentials: {e}")
        return False

def load_google_creds_from_firebase(user_id_safe):
    """Firebase에서 Google Calendar 인증 정보를 로드합니다."""
    try:
        ref = db.reference(f"users/{user_id_safe}/google_creds") if 'doctor' not in user_id_safe else db.reference(f"doctor_users/{user_id_safe}/google_creds")
        creds_data = ref.get()
        if creds_data and 'token' in creds_data:
            creds = Credentials(
                token=creds_data.get('token'), refresh_token=creds_data.get('refresh_token'),
                token_uri=creds_data.get('token_uri'), client_id=creds_data.get('client_id'),
                client_secret=creds_data.get('client_secret'), scopes=creds_data.get('scopes')
            )
            return creds
        return None
    except Exception as e:
        st.error(f"Failed to load Google credentials: {e}")
        return None

def get_google_calendar_service(user_id_safe):
    """
    사용자별로 Google Calendar 서비스 객체를 반환합니다.
    인증 정보가 없거나 유효하지 않으면 None을 반환합니다.
    """
    creds = st.session_state.get(f"google_creds_{user_id_safe}")
    
    if not creds:
        creds = load_google_creds_from_firebase(user_id_safe)
        if creds:
            st.session_state[f"google_creds_{user_id_safe}"] = creds

    if not creds:
        return None

    if creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            st.session_state[f"google_creds_{user_id_safe}"] = creds
            save_google_creds_to_firebase(user_id_safe, creds)
        except Exception as e:
            st.warning(f"Google Creds 갱신 실패: {e}")
            return None

    try:
        service = build('calendar', 'v3', credentials=creds)
        return service
    except HttpError as error:
        st.error(f'Google Calendar 서비스 생성 실패: {error}')
        return None
    except Exception as e:
        st.error(f"알 수 없는 오류: {e}")
        return None
