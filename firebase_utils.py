import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json

# 권한 범위를 URL에 표시된 것과 일치시키거나, 더 포괄적으로 설정
SCOPES = ['https://www.googleapis.com/auth/calendar']

def save_google_creds_to_firebase(safe_key, creds):
    try:
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        ref.set({'creds': creds.to_json()})
        st.success(f"✅ 인증 정보가 {safe_key} 계정에 성공적으로 연결되었습니다!")
    except Exception as e:
        st.error(f"❌ Firebase 저장 중 오류 발생: {e}")

def get_google_calendar_service(safe_key):
    # 0. 입력받은 safe_key가 없으면 세션에서라도 찾아봄
    if not safe_key:
        safe_key = st.session_state.get('user_email_safe') # 로그인 시 저장해둔 키가 있다면
    
    if not safe_key:
        st.error("❗ 사용자 정보를 확인할 수 없습니다. 로그인을 먼저 진행해주세요.")
        return None

    conf = dict(st.secrets["google_calendar"])
    client_config = {
        "web": {
            "client_id": conf.get("client_id"),
            "project_id": conf.get("project_id"),
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "client_secret": conf.get("client_secret"),
            "redirect_uris": [conf.get("redirect_uri")]
        }
    }

    # 1단계: 기존 DB 데이터 로드 시도
    data = db.reference(f'google_calendar_creds/{safe_key}').get()
    if data and 'creds' in data:
        creds = Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
        if creds.valid:
            return build('calendar', 'v3', credentials=creds)
        elif creds.refresh_token:
            try:
                creds.refresh(Request())
                save_google_creds_to_firebase(safe_key, creds)
                return build('calendar', 'v3', credentials=creds)
            except: pass

    # 2단계: 구글 인증 후 돌아온 코드 처리 (?code= 가 있는 경우)
    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            # flow 생성 시 state 검증을 피하기 위해 임의의 state 주입 또는 생략
            flow = Flow.from_client_config(
                client_config, 
                scopes=SCOPES, 
                redirect_uri=conf.get("redirect_uri")
            )
            flow.fetch_token(code=auth_code)
            
            # [핵심] 여기서 safe_key가 확실히 전달되어야 함
            save_google_creds_to_firebase(safe_key, flow.credentials)
            
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"❌ 인증 토큰 처리 중 오류: {e}")
            st.stop()

    # 3단계: 인증 버튼 표시
    flow = Flow.from_client_config(
        client_config, 
        scopes=SCOPES, 
        redirect_uri=conf.get("redirect_uri")
    )
    auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
    
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f'<a href="{auth_url}" target="_self" style="text-decoration:none;"><div style="background-color:#4285F4; color:white; padding:10px; border-radius:5px; text-align:center;">구글 계정 연동하기</div></a>', unsafe_allow_html=True)
    return None
