import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
import json

# 라이브러리가 설치되지 않았을 때를 대비한 예외 처리
try:
    from google_auth_oauthlib.flow import Flow
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
except ImportError:
    st.error("❌ 필수 라이브러리가 누락되었습니다. requirements.txt를 확인해주세요.")
    st.stop()

# 스코프 설정
SCOPES = [
    'https://www.googleapis.com/auth/calendar.events',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid'
]

def get_google_calendar_service(safe_key=None):
    # 1. 인증 후 돌아온 코드 처리 (URL 파라미터 확인)
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
            
            # 이메일 식별값 결정
            target_email = safe_key if safe_key else "skyeloveillustration@gmail.com"
            clean_key = target_email.replace('.', '_')
            
            # DB 저장
            ref = db.reference(f'google_calendar_creds/{clean_key}')
            ref.set({'creds': flow.credentials.to_json()})
            
            st.success(f"✅ 인증 성공! 계정: {target_email}")
            st.query_params.clear()
            st.rerun()
            return None
        except Exception as e:
            st.error(f"❌ 토큰 처리 실패: {e}")
            st.stop()

    # 2. 기존 데이터 로드
    target_key = safe_key if safe_key else "skyeloveillustration@gmail.com"
    clean_key = target_key.replace('.', '_')
    data = db.reference(f'google_calendar_creds/{clean_key}').get()
    
    if data and 'creds' in data:
        try:
            creds = Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
            if not creds.valid and creds.refresh_token:
                creds.refresh(Request())
                db.reference(f'google_calendar_creds/{clean_key}').update({'creds': creds.to_json()})
            return build('calendar', 'v3', credentials=creds)
        except:
            pass

    # 3. 인증 버튼 노출
    try:
        conf = dict(st.secrets["google_calendar"])
        flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
        auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
        st.markdown(f"**[📅 구글 계정 연동하기]({auth_url})**")
    except:
        st.error("Google API 설정(st.secrets)을 확인해주세요.")
        
    return None
