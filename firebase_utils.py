import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json

# 권한 범위 (가장 확실한 범위로 설정)
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

def save_google_creds_to_firebase(safe_key, creds):
    """이 함수가 실행되는지 화면에서 직접 확인하기 위해 st.write를 추가했습니다."""
    try:
        # 이메일의 점(.)은 Firebase 경로에서 오류를 일으키므로 반드시 변환
        clean_key = safe_key.replace('.', '_')
        ref = db.reference(f'google_calendar_creds/{clean_key}')
        
        st.write(f"📡 Firebase 저장 시도 중... 경로: google_calendar_creds/{clean_key}")
        
        ref.set({'creds': creds.to_json()})
        
        st.success(f"✅ [{clean_key}] 저장 완료 메시지가 떴습니다!")
        return True
    except Exception as e:
        st.error(f"❌ Firebase 저장 중 진짜 에러 발생: {e}")
        return False

def get_google_calendar_service(safe_key=None):
    # 1. 구글 다녀오기 전/후 이메일 유실 방지 (세션 고정)
    if safe_key:
        st.session_state['auth_email_key'] = safe_key
    
    active_key = st.session_state.get('auth_email_key')
    
    # 구글에서 돌아온 코드 확인
    auth_code = st.query_params.get("code")

    # 2. 기존 데이터 로드 (생략 가능하면 일단 스킵하고 인증부터 확인)
    if active_key and not auth_code:
        data = db.reference(f'google_calendar_creds/{active_key.replace(".", "_")}').get()
        if data and 'creds' in data:
            creds = Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
            if creds.valid:
                return build('calendar', 'v3', credentials=creds)

    # 3. 구글 인증 처리 (돌아온 직후)
    if auth_code:
        if not active_key:
            # [긴급 조치] 만약 이메일을 잃어버렸다면 테스트용으로 강제 지정
            active_key = "skyeloveillustration@gmail_com"
            st.warning(f"⚠️ 이메일 유실로 인해 임시 키({active_key})를 사용합니다.")

        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config(
                {"web": conf}, 
                scopes=SCOPES, 
                redirect_uri=conf["redirect_uri"]
            )
            flow.fetch_token(code=auth_code)
            
            # 여기서 저장이 실행되는지 눈으로 확인해야 합니다!
            success = save_google_creds_to_firebase(active_key, flow.credentials)
            
            if success:
                st.write("🔄 페이지를 새로고침합니다...")
                st.query_params.clear()
                st.rerun()
        except Exception as e:
            st.error(f"❌ 토큰 처리 실패: {e}")
            st.stop()

    # 4. 인증 버튼 (연동이 안 된 경우)
    conf = dict(st.secrets["google_calendar"])
    flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
    auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
    
    st.info(f"📅 [{active_key}] 계정 연동 버튼을 생성합니다.")
    st.markdown(f"**[🔗 구글 계정 연동하기]({auth_url})**")
    return None
