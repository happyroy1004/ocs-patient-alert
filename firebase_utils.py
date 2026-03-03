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

# [1] Firebase 초기화 로직 (가장 안전한 방식)
if not firebase_admin._apps:
    try:
        # secrets에서 정보를 가져와 딕셔너리로 변환
        fb_dict = dict(st.secrets["firebase"])
        # 필드가 누락되었을 경우를 대비해 databaseURL 확인
        db_url = fb_dict.get("databaseURL")
        
        cred = credentials.Certificate(fb_dict)
        firebase_admin.initialize_app(cred, {
            'databaseURL': db_url
        })
    except Exception as e:
        # 초기화에 실패하더라도 Streamlit이 멈추지 않게 로그만 남김
        print(f"Firebase 초기화 에러: {e}")

# [2] UI_MANAGER가 찾는 필수 함수들 (에러 방지를 위해 상단 배치)

def get_db_refs():
    """DB 참조 포인트를 안전하게 반환"""
    try:
        return {
            "users": db.reference('users'),
            "calendar": db.reference('google_calendar_creds'),
            "settings": db.reference('settings')
        }
    except Exception as e:
        st.error(f"DB 참조 획득 실패: {e}")
        return {}

def sanitize_path(email):
    """이메일을 Firebase 경로용으로 변환"""
    return email.replace('.', '_') if email else "unknown"

def recover_email(sanitized_email):
    """변환된 경로를 이메일로 복구"""
    return sanitized_email.replace('_', '.') if sanitized_email else ""

# [3] 구글 캘린더 관련 함수

SCOPES = [
    'https://www.googleapis.com/auth/calendar.events',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid'
]

def save_google_creds_to_firebase(clean_key, creds):
    try:
        db.reference(f'google_calendar_creds/{clean_key}').set({
            'creds': creds.to_json()
        })
        return True
    except Exception as e:
        st.error(f"저장 오류: {e}")
        return False

def get_google_calendar_service(safe_key=None):
    # URL 파라미터 확인 (Callback 처리)
    auth_code = st.query_params.get("code")

    if auth_code:
        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config(
                {"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"]
            )
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            
            # ID 토큰에서 이메일 추출
            id_info = id_token.verify_oauth2_token(
                creds.id_token, google_requests.Request(), conf["client_id"]
            )
            google_email = id_info.get('email')
            
            if google_email:
                target_key = safe_key if safe_key else google_email
                clean_key = sanitize_path(target_key)
                if save_google_creds_to_firebase(clean_key, creds):
                    st.success(f"✅ {google_email} 연동 성공!")
                    st.query_params.clear()
                    st.rerun()
        except Exception as e:
            st.error(f"인증 처리 에러: {e}")
            return None

    # 기존 데이터 로드
    if safe_key:
        clean_key = sanitize_path(safe_key)
        data = db.reference(f'google_calendar_creds/{clean_key}').get()
        if data and 'creds' in data:
            try:
                l_creds = Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
                if not l_creds.valid and l_creds.refresh_token:
                    l_creds.refresh(Request())
                    save_google_creds_to_firebase(clean_key, l_creds)
                return build('calendar', 'v3', credentials=l_creds)
            except: pass

    # 연동 버튼 출력
    try:
        conf = dict(st.secrets["google_calendar"])
        flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
        auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
        st.info("📅 구글 캘린더 연동이 필요합니다.")
        st.markdown(f'<a href="{auth_url}" target="_self" style="text-decoration:none;"><div style="background-color:#4285F4; color:white; padding:10px; border-radius:5px; text-align:center;">구글 계정 연동하기</div></a>', unsafe_allow_html=True)
    except:
        st.warning("구글 설정(secrets)을 확인해주세요.")
    return None
