import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.oauth2 import id_token  # 이메일 추출을 위해 필요
from google.auth.transport import requests as google_requests
from googleapiclient.discovery import build
import json

# 권한 범위 (이메일 정보를 가져오기 위해 openid, userinfo.email 추가 필수)
SCOPES = [
    'https://www.googleapis.com/auth/calendar.events',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid'
]

def get_google_calendar_service(safe_key=None):
    # 1. URL 파라미터에서 인증 코드 확인
    auth_code = st.query_params.get("code")

    # [중요] 인증 코드가 있다면 즉시 가로채서 처리
    if auth_code:
        st.info("🔄 구글 인증 응답을 감지했습니다. 계정 정보를 확인하여 DB에 저장합니다...")
        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config(
                {"web": conf}, 
                scopes=SCOPES, 
                redirect_uri=conf["redirect_uri"]
            )
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            
            # [핵심] 유실된 이메일을 구글 토큰에서 직접 추출 (계정 유추 방법)
            id_info = id_token.verify_oauth2_token(
                creds.id_token, 
                google_requests.Request(), 
                conf["client_id"]
            )
            google_email = id_info.get('email')
            
            if google_email:
                # Firebase 경로는 점(.)을 허용하지 않으므로 변환
                db_path_key = google_email.replace('.', '_')
                
                # 직접 Firebase에 쓰기 실행
                ref = db.reference(f'google_calendar_creds/{db_path_key}')
                ref.set({'creds': creds.to_json()})
                
                st.success(f"✅ [{google_email}] 계정 연동 성공! 데이터베이스에 저장되었습니다.")
                
                # URL 정리 후 리프레시
                st.query_params.clear()
                st.rerun()
            else:
                st.error("❗ 구글 토큰에서 이메일 정보를 읽어올 수 없습니다.")
        except Exception as e:
            st.error(f"❌ 인증 처리 중 오류 발생: {e}")
            st.stop()

    # 2. 평상시: DB에서 기존 데이터 로드
    if safe_key:
        db_path_key = safe_key.replace('.', '_')
        data = db.reference(f'google_calendar_creds/{db_path_key}').get()
        if data and 'creds' in data:
            try:
                loaded_creds = Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
                # 유효성 검사 및 갱신 로직 생략(간결화)
                return build('calendar', 'v3', credentials=loaded_creds)
            except:
                pass

    # 3. 인증이 안 된 경우 버튼 노출
    conf = dict(st.secrets["google_calendar"])
    flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
    auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
    
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 구글 계정 연동하기]({auth_url})**")
    return None
