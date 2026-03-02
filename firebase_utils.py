
import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json

# 권한 범위 (URL에 찍힌 scope와 정확히 일치시킴)
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

def save_google_creds_to_firebase(safe_key, creds):
    """Firebase 실시간 데이터베이스에 JSON 저장"""
    try:
        # 이메일 점(.)을 언더바(_)로 변환하는 sanitize_path 사용 권장
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        ref.set({'creds': creds.to_json()})
        st.success(f"✅ [{safe_key}] 계정 연동 성공! 데이터가 저장되었습니다.")
    except Exception as e:
        st.error(f"❌ Firebase 저장 실패: {e}")

def get_google_calendar_service(safe_key=None):
    # 1. 호출 시 전달된 이메일이 있다면 세션에 고정
    if safe_key:
        st.session_state['fixed_email_key'] = safe_key
    
    # 2. 세션에서 이메일 복구 (구글 인증 후 돌아왔을 때를 위함)
    active_key = st.session_state.get('fixed_email_key')
    
    # 디버깅용 (문제가 해결되면 삭제하세요)
    # st.write(f"현재 추적 중인 계정 키: {active_key}")

    if not active_key:
        st.warning("⚠️ 인증을 진행할 계정 정보가 없습니다. 먼저 로그인해 주세요.")
        return None

    conf = dict(st.secrets["google_calendar"])
    client_config = {
        "web": {
            "client_id": conf["client_id"],
            "project_id": conf["project_id"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "client_secret": conf["client_secret"],
            "redirect_uris": [conf["redirect_uri"]]
        }
    }

    # 3. 기존 데이터 로드 및 갱신 로직
    data = db.reference(f'google_calendar_creds/{active_key}').get()
    if data and 'creds' in data:
        try:
            creds = Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
            if creds.valid:
                return build('calendar', 'v3', credentials=creds)
            elif creds.refresh_token:
                creds.refresh(Request())
                save_google_creds_to_firebase(active_key, creds)
                return build('calendar', 'v3', credentials=creds)
        except:
            pass

    # 4. 구글 인증 코드 처리 (URL에 code가 있는 경우)
    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            flow = Flow.from_client_config(client_config, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
            flow.fetch_token(code=auth_code)
            
            # [핵심] 세션에 저장해둔 active_key(이메일)를 사용하여 저장
            save_google_creds_to_firebase(active_key, flow.credentials)
            
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"❌ 토큰 교환 실패: {e}")
            st.stop()

    # 5. 인증 버튼 생성
    flow = Flow.from_client_config(client_config, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
    auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
    
    st.info(f"📅 [{active_key}] 계정의 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 구글 계정 연동하기]({auth_url})**")
    return None
