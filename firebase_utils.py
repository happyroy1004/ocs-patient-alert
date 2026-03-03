import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from google.oauth2 import id_token
from google.auth.transport import requests as google_requests
import json

# 1. Firebase 초기화 (중복 방지)
if not firebase_admin._apps:
    try:
        fb_conf = dict(st.secrets["firebase"])
        cred = credentials.Certificate(fb_conf)
        firebase_admin.initialize_app(cred, {
            'databaseURL': st.secrets["firebase"]["databaseURL"]
        })
    except Exception as e:
        st.error(f"Firebase 초기화 실패: {e}")

# --- 기본 유틸리티 함수 ---

def get_db_refs():
    """UI 매니저 등에서 사용하는 DB 참조 포인트들을 반환"""
    return {
        "users": db.reference('users'),
        "calendar": db.reference('google_calendar_creds'),
        "settings": db.reference('settings')
    }

def sanitize_path(email):
    """이메일의 마침표(.)를 Firebase에서 허용하는 언더바(_)로 변환"""
    if not email:
        return "unknown"
    return email.replace('.', '_')

def recover_email(sanitized_email):
    """변환된 경로를 다시 이메일로 복구"""
    if not sanitized_email:
        return ""
    return sanitized_email.replace('_', '.')

# --- 구글 캘린더 핵심 로직 ---

SCOPES = [
    'https://www.googleapis.com/auth/calendar.events',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid'
]

def save_google_creds_to_firebase(clean_key, creds):
    """구글 인증 정보를 Firebase에 저장"""
    try:
        ref = db.reference(f'google_calendar_creds/{clean_key}')
        ref.set({
            'creds': creds.to_json()
        })
        return True
    except Exception as e:
        st.error(f"DB 저장 오류: {e}")
        return False

def get_google_calendar_service(safe_key=None):
    """
    구글 캘린더 서비스를 호출하거나, 인증 과정(Callback)을 처리합니다.
    """
    # [A] 리다이렉트 후 URL에 포함된 'code' 확인 (인증 처리 단계)
    auth_code = st.query_params.get("code")

    if auth_code:
        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config(
                {"web": conf}, 
                scopes=SCOPES, 
                redirect_uri=conf["redirect_uri"]
            )
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            
            # [핵심] ID 토큰에서 구글 이메일을 추출하여 저장 경로로 사용
            id_info = id_token.verify_oauth2_token(
                creds.id_token, 
                google_requests.Request(), 
                conf["client_id"]
            )
            google_email = id_info.get('email')
            
            if google_email:
                # 전달받은 safe_key가 없더라도 구글 이메일로 강제 저장
                target_key = safe_key if safe_key else google_email
                clean_key = sanitize_path(target_key)
                
                if save_google_creds_to_firebase(clean_key, creds):
                    st.success(f"✅ [{google_email}] 계정 연동 및 DB 저장 완료!")
                    st.query_params.clear()
                    st.rerun()
        except Exception as e:
            st.error(f"❌ 인증 처리 중 오류 발생: {e}")
            return None

    # [B] 기존에 저장된 데이터가 있는지 확인 (로드 단계)
    if safe_key:
        clean_key = sanitize_path(safe_key)
        data = db.reference(f'google_calendar_creds/{clean_key}').get()
        
        if data and 'creds' in data:
            try:
                loaded_creds = Credentials.from_authorized_user_info(
                    json.loads(data['creds']), SCOPES
                )
                # 토큰 만료 시 갱신
                if not loaded_creds.valid:
                    if loaded_creds.refresh_token:
                        loaded_creds.refresh(Request())
                        save_google_creds_to_firebase(clean_key, loaded_creds)
                
                return build('calendar', 'v3', credentials=loaded_creds)
            except Exception:
                pass

    # [C] 인증 데이터가 없으면 '연동 버튼' 노출
    conf = dict(st.secrets["google_calendar"])
    flow = Flow.from_client_config(
        {"web": conf}, 
        scopes=SCOPES, 
        redirect_uri=conf["redirect_uri"]
    )
    auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
    
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(
        f'<a href="{auth_url}" target="_self" style="text-decoration:none;">'
        f'<div style="background-color:#4285F4; color:white; padding:10px; border-radius:5px; text-align:center;">'
        f'구글 계정 연동하기</div></a>', 
        unsafe_allow_html=True
    )
    return None
