import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json

# 권한 범위 설정 (URL에 찍힌 범위와 일치시킴)
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

def save_google_creds_to_firebase(safe_key, creds):
    """Firebase에 확실하게 저장하고 성공 메시지 출력"""
    try:
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        ref.set({'creds': creds.to_json()})
        st.success(f"✅ {safe_key} 계정의 구글 연동 데이터가 저장되었습니다!")
    except Exception as e:
        st.error(f"❌ Firebase 저장 실패: {e}")

def get_google_calendar_service(safe_key):
    # 1. safe_key가 들어오면 세션에 백업 (구글 갔다 올 때 유실 방지)
    if safe_key:
        st.session_state['pending_safe_key'] = safe_key
    else:
        # 인자값이 없으면 세션에 저장해둔 키를 꺼내옴
        safe_key = st.session_state.get('pending_safe_key')

    # 2. 기존 저장된 데이터가 있는지 확인
    if safe_key:
        data = db.reference(f'google_calendar_creds/{safe_key}').get()
        if data and 'creds' in data:
            try:
                creds = Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
                if creds.valid:
                    return build('calendar', 'v3', credentials=creds)
                elif creds.refresh_token:
                    creds.refresh(Request())
                    save_google_creds_to_firebase(safe_key, creds)
                    return build('calendar', 'v3', credentials=creds)
            except:
                pass

    # 3. 구글에서 인증 코드(code)를 가지고 돌아온 경우 처리
    auth_code = st.query_params.get("code")
    if auth_code:
        # 저장할 키가 없으면 진행 불가
        if not safe_key:
            st.error("❗ 인증을 처리할 사용자 정보를 잃어버렸습니다. 다시 시도해주세요.")
            st.stop()
            
        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config(
                {"web": {
                    "client_id": conf["client_id"],
                    "project_id": conf["project_id"],
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "client_secret": conf["client_secret"],
                    "redirect_uris": [conf["redirect_uri"]] # 슬래시 없는 URI 그대로 사용
                }}, 
                scopes=SCOPES, 
                redirect_uri=conf["redirect_uri"]
            )
            flow.fetch_token(code=auth_code)
            
            # 드디어 저장!
            save_google_creds_to_firebase(safe_key, flow.credentials)
            
            # 성공 후 URL 정리 및 리런
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"❌ 토큰 교환 에러: {e}")
            st.stop()

    # 4. 연동 버튼 표시 (인증이 안 되어 있는 경우)
    conf = dict(st.secrets["google_calendar"])
    flow = Flow.from_client_config(
        {"web": {
            "client_id": conf["client_id"],
            "project_id": conf["project_id"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "client_secret": conf["client_secret"],
            "redirect_uris": [conf["redirect_uri"]]
        }}, 
        scopes=SCOPES, 
        redirect_uri=conf["redirect_uri"]
    )
    
    auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
    
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 구글 계정 연동하기]({auth_url})**")
    return None
