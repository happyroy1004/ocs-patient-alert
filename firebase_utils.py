# firebase_utils.py

import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.oauth2 import id_token
from google.auth.transport import requests as google_requests
import json

# config에서 권한 범위(SCOPES) 가져오기
from config import SCOPES

# --- 1. Firebase 및 Google OAuth 설정 로드 ---

try:
    # Secrets에서 Firebase 설정 로드
    FIREBASE_CREDENTIALS = dict(st.secrets["firebase"])
    # 보내주신 TOML 구조에 맞춰 database_url 키를 정확히 매칭합니다.
    DB_URL = st.secrets.get("database_url") or FIREBASE_CREDENTIALS.get("database_url")

    # Google Calendar 설정 로드
    google_calendar_secrets = st.secrets.get("google_calendar")
    GOOGLE_CALENDAR_CLIENT_SECRET = dict(google_calendar_secrets) if google_calendar_secrets else {}
    
except Exception as e:
    st.error(f"🚨 Secrets 로드 오류: {e}")
    FIREBASE_CREDENTIALS = None
    DB_URL = None

# --- 2. DB 초기화 및 레퍼런스 함수 ---

if not firebase_admin._apps:
    try:
        if FIREBASE_CREDENTIALS and DB_URL:
            cred = credentials.Certificate(FIREBASE_CREDENTIALS)
            firebase_admin.initialize_app(cred, {'databaseURL': DB_URL})
        else:
            st.error("❌ Firebase 설정(URL)이 누락되었습니다. Secrets를 확인하세요.")
    except Exception as e:
        st.error(f"❌ Firebase 앱 초기화 실패: {e}")

@st.cache_resource

def get_db_refs():
    # ui_manager.py의 (users_ref, doctor_users_ref, db_ref_func) 대응
    users_ref = db.reference('users')
    doctor_users_ref = db.reference('doctor_users')
    def db_ref_func(path):
        return db.reference(path)
    return users_ref, doctor_users_ref, db_ref_func

# --- 3. 유틸리티 및 데이터 관리 함수 (단일 노드 최적화) ---

def sanitize_path(email):
    """이메일 주소의 마침표(.)를 언더바(_)로 변환하여 DB 키로 사용합니다."""
    return email.replace('.', '_') if email else "unknown"

def recover_email(safe_key):
    """사용자 노드에서 실제 이메일을 찾습니다."""
    db_ref = db.reference()
    # 통합된 구조라도 이메일 복구는 users 혹은 doctor_users 노드에서 수행합니다.
    for path in [f'users/{safe_key}', f'doctor_users/{safe_key}']:
        try:
            data = db_ref.child(path).get()
            if data and isinstance(data, dict) and 'email' in data:
                return data['email']
        except:
            continue
    return safe_key.replace('_', '.') if safe_key else ""

def save_google_creds_to_firebase(safe_key, creds):
    """통합된 경로(google_calendar_creds)에 인증 정보를 저장합니다."""
    try:
        # 단일 표준 경로 사용
        db.reference(f'google_calendar_creds/{safe_key}').set({
            'creds': creds.to_json()
        })
        return True
    except Exception as e:
        st.error(f"인증 정보 저장 실패: {e}")
        return False

def load_google_creds(safe_key):
    """통합된 단일 경로(google_calendar_creds)에서만 인증 정보를 로드합니다."""
    if not safe_key:
        return None

    try:
        # 여러 노드를 뒤질 필요 없이 지정된 경로만 확인합니다.
        data = db.reference(f'google_calendar_creds/{safe_key}').get()
        if data and 'creds' in data:
            creds_info = data['creds']
            # JSON 문자열인 경우 딕셔너리로 변환
            if isinstance(creds_info, str):
                creds_info = json.loads(creds_info)
            return Credentials.from_authorized_user_info(creds_info, SCOPES)
    except Exception as e:
        # 로그 기록 (필요 시)
        pass
    return None

# --- 4. Google Calendar 서비스 구축 및 인증 흐름 ---

def get_google_calendar_service(safe_key):
    """서비스 객체를 반환하거나 인증 리다이렉트 흐름을 처리합니다."""
    if not safe_key:
        return None

    # 1. 통합 경로에서 인증 정보 로드
    creds = load_google_creds(safe_key)

    # 2. 토큰 갱신 로직
    if creds:
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                save_google_creds_to_firebase(safe_key, creds)
            except:
                creds = None
        
        if creds and creds.valid:
            return build('calendar', 'v3', credentials=creds)

    # 3. 인증 플로우 (인증 정보가 없거나 만료된 경우)
    conf = GOOGLE_CALENDAR_CLIENT_SECRET
    if not conf:
        return None

    redirect_uri = conf.get("redirect_uri")
    flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=redirect_uri)
    
    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            flow.fetch_token(code=auth_code)
            new_creds = flow.credentials
            save_google_creds_to_firebase(safe_key, new_creds)
            st.success("✅ 구글 캘린더 연동 완료!")
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"인증 처리 실패: {e}")
    else:
        # 연동 버튼 출력
        auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
        st.info("📅 구글 캘린더 연동이 필요합니다.")
        st.markdown(
            f'<a href="{auth_url}" target="_self" style="text-decoration:none;">'
            f'<div style="background-color:#4285F4; color:white; padding:10px; border-radius:5px; text-align:center;">'
            f'구글 계정 연동하기</div></a>', 
            unsafe_allow_html=True
        )
    
    return None

def get_db_refs():
    """ui_manager.py 32번 라인 대응: 3개 반환"""
    users_ref = db.reference('users')
    doctor_users_ref = db.reference('doctor_users')
    def db_ref_func(path):
        return db.reference(path)
    return users_ref, doctor_users_ref, db_ref_func

def sanitize_path(email):
    return email.replace('.', '_') if email else "unknown"

def recover_email(safe_key):
    return safe_key.replace('_', '.') if safe_key else ""

def save_google_creds_to_firebase(safe_key, creds):
    db.reference(f'google_calendar_creds/{safe_key}').set({'creds': creds.to_json()})

def load_google_creds(safe_key):
    """이름을 ui_manager의 import문과 일치시킴"""
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    if data and 'creds' in data:
        return Credentials.from_authorized_user_info(json.loads(data['creds']))
    return None
