import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from googleapiclient.discovery import build
import json

# 스코프 설정 (GCP 콘솔과 똑같아야 함)
SCOPES = [
    'https://www.googleapis.com/auth/calendar.events',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid'
]

def get_google_calendar_service(safe_key=None):
    # [1] URL에 'code'가 보이면 무조건 가로채서 저장 실행
    auth_code = st.query_params.get("code")

    if auth_code:
        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
            flow.fetch_token(code=auth_code)
            
            # [테스트용 강제 설정] 이메일 유실 방지를 위해 본인 계정으로 직접 경로 지정
            # 이메일 점(.)은 언더바(_)로 바꿔야 에러가 안 납니다.
            test_email_key = "skyeloveillustration@gmail_com" 
            
            ref = db.reference(f'google_calendar_creds/{test_email_key}')
            ref.set({'creds': flow.credentials.to_json()})
            
            st.success(f"✅ DB에 데이터가 강제 저장되었습니다! (경로: {test_email_key})")
            st.query_params.clear()
            st.rerun()
            return None
        except Exception as e:
            st.error(f"❌ 토큰 교환 중 오류: {e}")
            st.stop()

    # [2] 평상시 데이터 불러오기
    target_key = safe_key.replace('.', '_') if safe_key else "skyeloveillustration@gmail_com"
    data = db.reference(f'google_calendar_creds/{target_key}').get()
    
    if data and 'creds' in data:
        from google.oauth2.credentials import Credentials
        return build('calendar', 'v3', credentials=Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES))

    # [3] 인증 버튼
    conf = dict(st.secrets["google_calendar"])
    flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
    auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
    st.markdown(f"**[📅 구글 계정 연동하기]({auth_url})**")
    return None
