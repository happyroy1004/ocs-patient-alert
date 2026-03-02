import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json

# --- 수정된 Google 인증 관리 부분 ---

def save_google_creds_to_firebase(safe_key, creds):
    """안전하게 JSON 문자열로 변환하여 저장"""
    try:
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        creds_json = creds.to_json()
        ref.set({'creds': creds_json})
        st.success(f"✅ Firebase에 인증 정보 저장 완료! (Key: {safe_key})")
    except Exception as e:
        st.error(f"❌ Firebase 저장 실패: {e}")

def load_google_creds_from_firebase(safe_key):
    """Firebase에서 데이터를 가져와 Credentials 객체로 복원"""
    try:
        data = db.reference(f'google_calendar_creds/{safe_key}').get()
        if data and 'creds' in data:
            # SCOPES를 반드시 포함하여 로드
            return Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
    except Exception as e:
        st.warning(f"⚠️ 데이터 로드 중 오류: {e}")
    return None

def get_google_calendar_service(safe_key):
    # 0. 세션 캐시 확인
    if st.session_state.get('google_calendar_service'):
        return st.session_state.google_calendar_service

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

    # 1단계: 기존 저장된 데이터 불러오기
    creds = load_google_creds_from_firebase(safe_key)
    
    if creds:
        # 토큰이 만료되었다면 갱신 시도
        if not creds.valid and creds.refresh_token:
            try:
                creds.refresh(Request())
                save_google_creds_to_firebase(safe_key, creds)
            except Exception as e:
                st.error(f"🔄 토큰 갱신 실패: {e}")
                creds = None # 실패 시 다시 인증하도록 초기화

        if creds and creds.valid:
            try:
                service = build('calendar', 'v3', credentials=creds)
                st.session_state.google_calendar_service = service
                return service
            except Exception as e:
                st.error(f"🛠️ 서비스 빌드 실패: {e}")

    # 2단계: 인증 코드 처리 (URL에 ?code= 가 있을 때)
    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            flow = Flow.from_client_config(
                client_config, 
                scopes=SCOPES, 
                redirect_uri=conf.get("redirect_uri")
            )
            # state 검증 에러 방지를 위해 fetch_token 직접 호출
            flow.fetch_token(code=auth_code)
            new_creds = flow.credentials
            
            # 여기서 저장이 제대로 되는지 확인이 필요함
            save_google_creds_to_firebase(safe_key, new_creds)
            
            st.session_state.google_calendar_service = build('calendar', 'v3', credentials=new_creds)
            st.success("🎉 인증에 성공했습니다! 페이지를 새로고침합니다.")
            
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"❌ 인증 토큰 획득 실패: {e}")
            st.stop()

    # 3단계: 인증 링크 표시 (데이터가 없거나 유효하지 않을 때만)
    flow = Flow.from_client_config(
        client_config, 
        scopes=SCOPES, 
        redirect_uri=conf.get("redirect_uri")
    )
    auth_url, _ = flow.authorization_url(
        prompt='consent', 
        access_type='offline',
        include_granted_scopes='true'
    )
    
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 여기를 클릭하여 구글 계정 연동하기]({auth_url})**")
    return None
