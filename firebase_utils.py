# firebase_utils.py

import streamlit as st # 💡 st.secrets 및 캐싱을 위해 필요
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

# local imports: config에서 순수한 상수(SCOPES)만 가져옵니다.
from config import SCOPES

# 💡 st.secrets를 사용하여 인증 정보를 로드하고 전역 변수로 설정
try:
    # 1. Firebase Admin SDK 인증 정보 로드: [firebase] 섹션 전체를 딕셔너리로 변환하여 로드
    # 🚨 수정: dict() 생성자를 사용해 st.secrets 객체를 안전하게 복사/변환합니다.
    FIREBASE_CREDENTIALS = dict(st.secrets["firebase"]) 
    
    # 2. DB URL 로드: 최상위 database_url 키를 참조하도록 통일
    DB_URL = st.secrets["database_url"] 

    # 3. Google Calendar Client Secret 로드
    GOOGLE_CALENDAR_CLIENT_SECRET = dict(st.secrets["google_calendar"])
    
except KeyError as e:
    st.error(f"🚨 중요: Secrets.toml 설정 오류. '{e.args[0]}' 키를 찾을 수 없습니다. secrets.toml 파일의 키 이름과 위치를 확인해 주세요.")
    FIREBASE_CREDENTIALS = None
    DB_URL = None
    GOOGLE_CALENDAR_CLIENT_SECRET = None
except Exception as e:
    st.error(f"🚨 Secrets 로드 중 예상치 못한 오류 발생: {e}")
    FIREBASE_CREDENTIALS = None
    DB_URL = None
    GOOGLE_CALENDAR_CLIENT_SECRET = None


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
            # Secrets 로드 실패 시 초기화 시도 자체를 건너뜀
            if FIREBASE_CREDENTIALS is None or DB_URL is None:
                st.warning("DB 연결 정보가 불완전하여 초기화를 건너뜁니다.")
                return None, None, None

            # Admin SDK에 전달하기 전에 DB URL 관련 키(Admin SDK는 필요 없음)는 제거합니다.
            # FIREBASE_CREDENTIALS는 이미 dict 객체이므로 안전하게 copy() 호출 가능
            creds_for_init = FIREBASE_CREDENTIALS.copy()
            if 'FIREBASE_DATABASE_URL' in creds_for_init: 
                 del creds_for_init['FIREBASE_DATABASE_URL']
            
            # Firebase Admin SDK가 기대하는 딕셔너리(서비스 계정)를 전달합니다.
            cred = credentials.Certificate(creds_for_init)
            
            # DB URL을 사용하여 앱을 초기화합니다.
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
    """
    safe_email = email.replace('.', '_')
    return safe_email


def save_google_creds_to_firebase(safe_key, creds):
    """Google 캘린더 OAuth2 Credentials 객체를 Firebase에 저장합니다 (pickle 직렬화)."""
    creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
    
    pickled_creds = pickle.dumps(creds)
    encoded_creds = pickled_creds.hex()
    
    creds_ref.set({'creds': encoded_creds})


def load_google_creds_from_firebase(safe_key):
    """Firebase에서 Google Calendar OAuth2 Credentials 객체를 로드합니다."""
    creds_ref = db.reference(f'google_calendar_creds/{safe_key}')
    data = creds_ref.get()
    
    if data and 'creds' in data:
        encoded_creds = data['creds']
        pickled_creds = bytes.fromhex(encoded_creds)
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
        creds.refresh(Request())
        save_google_creds_to_firebase(safe_key, creds)
    
    elif not creds or not creds.valid:
        
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
                
                save_google_creds_to_firebase(safe_key, creds)

                st.session_state.google_calendar_auth_needed = False
                st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
                st.success("🎉 Google Calendar 연동이 성공적으로 완료되었습니다!")
                st.rerun()
            except Exception as e:
                st.error(f"인증 코드 오류: 코드를 다시 확인하거나 [Google 인증 링크]({auth_url})를 다시 열어 시도하세요. ({e})")
                return

    if creds and creds.valid:
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
