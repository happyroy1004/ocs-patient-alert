import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json

# [중요] 스코프를 최소한으로 설정 (GCP 콘솔 설정과 반드시 일치해야 함)
SCOPES = [
    'https://www.googleapis.com/auth/calendar.events',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid'
]

def get_google_calendar_service(safe_key=None):
    # 1. URL에서 code 확인 (구글 리다이렉트 직후)
    auth_code = st.query_params.get("code")

    if auth_code:
        st.write("🔄 인증 코드를 확인했습니다. 토큰을 가져오는 중...")
        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config(
                {"web": conf}, 
                scopes=SCOPES, 
                redirect_uri=conf["redirect_uri"]
            )
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            
            # 2. 이메일 추출 (이게 안 되면 저장이 안 됨)
            # 만약 id_token 파싱이 어렵다면, 임시로 고정 이메일을 사용해 저장 테스트
            target_email = "skyeloveillustration@gmail.com" # 우선 본인 계정으로 고정 테스트
            clean_key = target_email.replace('.', '_')
            
            # 3. Firebase 저장 시도
            ref = db.reference(f'google_calendar_creds/{clean_key}')
            ref.set({'creds': creds.to_json()})
            
            st.success(f"✅ DB 저장 성공! (계정: {target_email})")
            st.query_params.clear()
            st.rerun()
            return None
            
        except Exception as e:
            st.error(f"❌ 인증 에러 발생: {e}")
            st.info("💡 팁: GCP 콘솔의 'OAuth 동의 화면'에서 스코프가 정확히 추가되었는지 확인하세요.")
            st.stop()

    # 기존 데이터 로드 로직 (평상시)
    if safe_key:
        clean_key = safe_key.replace('.', '_')
        data = db.reference(f'google_calendar_creds/{clean_key}').get()
        if data and 'creds' in data:
            try:
                c = Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
                return build('calendar', 'v3', credentials=c)
            except: pass

    # 버튼 노출
    conf = dict(st.secrets["google_calendar"])
    flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
    auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
    
    st.markdown(f"**[📅 구글 계정 연동하기]({auth_url})**")
    return None
