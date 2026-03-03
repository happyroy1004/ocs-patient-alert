import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json

# 권한 범위 (ID 토큰 파싱을 위해 userinfo.email 추가)
SCOPES = [
    'https://www.googleapis.com/auth/calendar.events',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid'
]

def get_google_calendar_service(safe_key=None):
    # [1] 긴급 가로채기: URL에 code가 있는지 확인
    # Streamlit은 페이지가 리런될 때 이 파라미터를 읽을 수 있습니다.
    auth_code = st.query_params.get("code")

    if auth_code:
        st.info("🔄 구글 인증 응답을 처리 중입니다. 잠시만 기다려주세요...")
        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config(
                {"web": conf}, 
                scopes=SCOPES, 
                redirect_uri=conf["redirect_uri"]
            )
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            
            # [핵심] safe_key가 유실되었을 확률 99%. 
            # 구글이 돌려준 ID 토큰에서 로그인한 이메일을 강제로 꺼냅니다.
            from google.oauth2 import id_token
            from google.auth.transport import requests
            id_info = id_token.verify_oauth2_token(creds.id_token, requests.Request(), conf["client_id"])
            google_email = id_info.get('email')
            
            if google_email:
                clean_key = google_email.replace('.', '_')
                # DB 저장 강제 실행
                ref = db.reference(f'google_calendar_creds/{clean_key}')
                ref.set({'creds': creds.to_json()})
                
                st.success(f"✅ [{google_email}] 계정 연동 및 DB 저장 완료!")
                
                # 쿼리 정리 후 깨끗한 상태로 리런
                st.query_params.clear()
                st.rerun()
                return None
        except Exception as e:
            st.error(f"❌ 인증 저장 실패: {e}")
            st.stop()

    # [2] 기존 저장된 데이터 로드 (평상시 동작)
    if safe_key:
        clean_key = safe_key.replace('.', '_')
        data = db.reference(f'google_calendar_creds/{clean_key}').get()
        if data and 'creds' in data:
            try:
                loaded_creds = Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
                if loaded_creds.valid:
                    return build('calendar', 'v3', credentials=loaded_creds)
                elif loaded_creds.refresh_token:
                    loaded_creds.refresh(Request())
                    db.reference(f'google_calendar_creds/{clean_key}').update({'creds': loaded_creds.to_json()})
                    return build('calendar', 'v3', credentials=loaded_creds)
            except:
                pass

    # [3] 연동 버튼 노출
    conf = dict(st.secrets["google_calendar"])
    flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
    auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
    
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f'<a href="{auth_url}" target="_self" style="text-decoration:none;"><div style="background-color:#4285F4; color:white; padding:10px; border-radius:5px; text-align:center;">구글 계정 연동하기</div></a>', unsafe_allow_html=True)
    return None
