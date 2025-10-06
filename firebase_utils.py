# firebase_utils.py

import streamlit as st # 💡 캐싱을 위해 Streamlit 임포트
import firebase_admin
from firebase_admin import credentials, db, auth
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import os
import io
import pickle
import json

# local imports: 상대 경로 임포트(.)를 절대 경로 임포트로 수정
from config import (
    SCOPES, FIREBASE_CREDENTIALS, GOOGLE_CALENDAR_CLIENT_SECRET, 
    GOOGLE_CALENDAR_CREDENTIAL_FILE, DB_URL
)

# --- 1. DB 레퍼런스 및 초기화 ---

@st.cache_resource
def get_db_refs():
    """
    Firebase Admin SDK를 초기화하고 DB 레퍼런스 객체를 반환합니다.
    @st.cache_resource로 앱 수명 주기 동안 단 한 번만 실행되도록 보장합니다.
    """
    users_ref = None
    doctor_users_ref = None
    
    # Firebase Admin SDK 초기화 확인 및 실행
    if not firebase_admin._apps:
        try:
            # FIREBASE_CREDENTIALS는 secrets.toml에서 로드된 딕셔너리여야 합니다.
            if isinstance(FIREBASE_CREDENTIALS, dict):
                cred = credentials.Certificate(FIREBASE_CREDENTIALS)
            else:
                # 딕셔너리가 아닌 경우 (예: 로드 실패 또는 잘못된 형식), 초기화 실패를 명확히 함
                st.error("🚨 Firebase 인증 정보를 딕셔너리 형태로 로드하지 못했습니다. Secrets 설정을 확인하세요.")
                return None, None, None

            firebase_admin.initialize_app(cred, {'databaseURL': DB_URL})
            
        except Exception as e:
            st.error(f"❌ Firebase 앱 초기화 실패: {e}")
            return None, None, None # 초기화 실패 시 None 반환

    # 초기화 성공 시에만 레퍼런스 반환
    if firebase_admin._apps:
        base_ref = db.reference()
        users_ref = base_ref.child('users')
        doctor_users_ref = base_ref.child('doctor_users')
        
        # 동적으로 경로를 참조하기 위한 함수
        def db_ref_func(path):
            return base_ref.child(path)
            
        return users_ref, doctor_users_ref, db_ref_func
        
    return None, None, None


# --- 2. Google Calendar 인증 및 Creds 관리 ---

def sanitize_path(email):
    """
    이메일 주소를 Firebase Realtime Database 키로 사용할 수 있도록 정리합니다.
    (., $, #, [, ], /, \ 등 특수 문자 제거)
    """
    # 2024년 4월 기준, RTDB 키로 사용할 수 없는 문자들을 대체합니다.
    # '.'을 '_'로 대체하는 것은 일반적인 관례입니다.
    safe_email = email.replace('.', '_')
    return safe_email


def save_google_creds_to_firebase(safe_key, creds):
    """Google 캘린더 OAuth2 Credentials 객체를 Firebase에 저장합니다 (pickle 직렬화)."""
    # Google Calendar 인증 정보 저장을 위한 Firebase 레퍼런스
    creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
    
    # Credentials 객체를 pickle로 직렬화
    pickled_creds = pickle.dumps(creds)
    # 바이너리 데이터를 Base64로 인코딩하여 문자열로 저장
    encoded_creds = pickled_creds.hex()
    
    creds_ref.set({'creds': encoded_creds})


def load_google_creds_from_firebase(safe_key):
    """Firebase에서 Google Calendar OAuth2 Credentials 객체를 로드합니다."""
    creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
    data = creds_ref.get()
    
    if data and 'creds' in data:
        encoded_creds = data['creds']
        # Base64 문자열을 디코딩
        pickled_creds = bytes.fromhex(encoded_creds)
        # pickle 역직렬화
        creds = pickle.loads(pickled_creds)
        return creds
    return None


# --- 3. Google Calendar Service 로드/인증 흐름 ---

def get_google_calendar_service(safe_key):
    """
    Google Calendar 서비스 객체를 로드하거나, 인증이 필요하면 인증 흐름을 시작합니다.
    결과는 st.session_state에 저장됩니다.
    """
    st.session_state.google_calendar_service = None
    creds = load_google_creds_from_firebase(safe_key)

    if creds and creds.expired and creds.refresh_token:
        # 토큰 갱신이 필요하면 갱신
        creds.refresh(Request())
        save_google_creds_to_firebase(safe_key, creds)
    
    elif not creds or not creds.valid:
        # 인증 또는 재인증이 필요한 경우
        
        # client_secret.json 파일 내용 로드
        if isinstance(GOOGLE_CALENDAR_CLIENT_SECRET, dict):
            client_config = GOOGLE_CALENDAR_CLIENT_SECRET
        else:
            st.warning("Google Client Secret 정보를 로드하지 못했습니다. Secrets 설정을 확인하세요.")
            return

        flow = Flow.from_client_config(
            client_config, 
            scopes=SCOPES, 
            redirect_uri='urn:ietf:wg:oauth:2.0:oob' # Streamlit OOB (Out-of-Band) URI
        )

        auth_url, _ = flow.authorization_url(prompt='consent')

        st.session_state.google_calendar_auth_needed = True
        st.info("Google Calendar 연동을 위해 인증이 필요합니다.")
        st.markdown(f"[**Google 인증 링크 열기**]({auth_url})")

        verification_code = st.text_input("위 링크에서 받은 인증 코드(Verification Code)를 입력하세요", key="google_auth_code_input")
        
        if verification_code:
            try:
                flow.fetch_token(code=verification_code)
                creds = flow.credentials
                
                # Firebase에 Credentials 객체 저장
                save_google_creds_to_firebase(safe_key, creds)

                st.session_state.google_calendar_auth_needed = False
                st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
                st.success("🎉 Google Calendar 연동이 성공적으로 완료되었습니다!")
                st.rerun()
            except Exception as e:
                st.error(f"인증 코드 오류: 코드를 다시 확인하거나 [Google 인증 링크]({auth_url})를 다시 열어 시도하세요. ({e})")
                return

    if creds and creds.valid:
        # 인증된 서비스 객체 생성
        st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)


def recover_email(safe_key):
    """Firebase의 user 노드에서 safe_key에 해당하는 실제 이메일을 찾습니다."""
    try:
        data = db.reference('users').child(safe_key).get()
        if data and 'email' in data:
            return data['email']
    except Exception:
        pass
        
    try:
        data = db.reference('doctor_users').child(safe_key).get()
        if data and 'email' in data:
            return data['email']
    except Exception:
        pass
        
    return None
