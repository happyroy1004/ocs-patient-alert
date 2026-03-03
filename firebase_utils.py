import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json

# 권한 범위 (가장 안정적인 범위)
SCOPES = ['https://www.googleapis.com/auth/calendar.events', 'https://www.googleapis.com/auth/userinfo.email', 'openid']

def get_google_calendar_service(safe_key=None):
    # [1] URL 파라미터 확인 (구글에서 돌아온 직후)
    auth_code = st.query_params.get("code")
    
    if auth_code:
        st.info("🔍 구글 응답 수신 중... 데이터베이스 저장을 시도합니다.")
        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config(
                {"web": conf}, 
                scopes=SCOPES, 
                redirect_uri=conf["redirect_uri"]
            )
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            
            # [핵심] safe_key가 유실되었다면, 구글 인증 정보에서 직접 이메일 추출
            # 이메일이 일치할 때 가장 빛을 발하는 로직입니다.
            if not safe_key:
                from google.oauth2 import id_token
                from google.auth.transport import requests
                # ID 토큰에서 이메일 추출 시도
                id_info = id_token.verify_oauth2_token(creds.id_token, requests.Request(), conf["client_id"])
                safe_key = id_info.get('email')
                st.write(f"📧 인증 계정 확인됨: {safe_key}")

            if safe_key:
                clean_key = safe_key.replace('.', '_')
                ref = db.reference(f'google_calendar_creds/{clean_key}')
                ref.set({'creds': creds.to_json()})
                
                st.success(f"🎉 [{clean_key}] 계정 연동 성공! DB 저장 완료.")
                st.query_params.clear()
                st.rerun()
            else:
                st.error("❗ 저장할 대상 이메일을 식별할 수 없습니다.")
        except Exception as e:
            st.error(f"❌ 인증 처리 중 오류: {e}")
            st.stop()

    # [2] 기존 저장 데이터 확인 로직 (이미 되어있다면 생략 가능)
    if safe_key:
        clean_key = safe_key.replace('.', '_')
        data = db.reference(f'google_calendar_creds/{clean_key}').get()
        if data and 'creds' in data:
            return build('calendar', 'v3', credentials=Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES))

    # [3] 인증 버튼 노출
    conf = dict(st.secrets["google_calendar"])
    flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
    auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
    
    st.markdown(f"**[📅 구글 캘린더 연동하기]({auth_url})**")
    return None
