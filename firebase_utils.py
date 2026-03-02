import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json

# --- 1. Firebase 저장/로드 (디버깅 로그 포함) ---

def save_google_creds_to_firebase(safe_key, creds):
    if not safe_key:
        st.error("❌ 저장 실패: safe_key가 없습니다!")
        return
    try:
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        ref.set({'creds': creds.to_json()})
        st.success(f"✅ Firebase 저장 성공: {safe_key}")
    except Exception as e:
        st.error(f"❌ Firebase 쓰기 에러: {e}")

def load_google_creds_from_firebase(safe_key):
    if not safe_key: return None
    try:
        data = db.reference(f'google_calendar_creds/{safe_key}').get()
        if data and 'creds' in data:
            return Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
    except:
        pass
    return None

# --- 2. 메인 서비스 함수 ---

def get_google_calendar_service(safe_key):
    # [중요] URL에서 state(safe_key)를 복구 시도 (리런 시 유실 방지)
    returned_state = st.query_params.get("state")
    current_key = returned_state if returned_state else safe_key

    if not current_key:
        st.warning("⚠️ 사용자 식별 키(safe_key)를 찾을 수 없습니다. 로그인을 먼저 해주세요.")
        return None

    # 1. 기존 데이터 확인
    creds = load_google_creds_from_firebase(current_key)
    if creds and creds.valid:
        return build('calendar', 'v3', credentials=creds)

    # 2. 인증 코드 처리 (구글에서 돌아온 경우)
    auth_code = st.query_params.get("code")
    if auth_code:
        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config(
                {"web": {
                    "client_id": conf.get("client_id"),
                    "project_id": conf.get("project_id"),
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "client_secret": conf.get("client_secret"),
                    "redirect_uris": [conf.get("redirect_uri")]
                }}, 
                scopes=SCOPES, 
                redirect_uri=conf.get("redirect_uri")
            )
            flow.fetch_token(code=auth_code)
            
            # 여기서 current_key(safe_key)를 사용하여 저장
            save_google_creds_to_firebase(current_key, flow.credentials)
            
            st.success("인증 완료! 데이터를 저장했습니다.")
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"토큰 획득 중 오류: {e}")
            st.stop()

    # 3. 인증 링크 생성 (state에 safe_key를 담아 보냄)
    flow = Flow.from_client_config(
        {"web": {
            "client_id": conf.get("client_id"),
            "project_id": conf.get("project_id"),
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "client_secret": conf.get("client_secret"),
            "redirect_uris": [conf.get("redirect_uri")]
        }}, 
        scopes=SCOPES, 
        redirect_uri=conf.get("redirect_uri")
    )
    
    # state=current_key 를 추가하여 구글로 보냈다가 다시 돌려받음
    auth_url, _ = flow.authorization_url(
        prompt='consent', 
        access_type='offline',
        state=current_key 
    )
    
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 구글 계정 연동하기]({auth_url})**")
    return None
