import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json

# 권한 범위 (정확히 일치)
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

def get_google_calendar_service(safe_key=None):
    # [1] 최우선 순위: 구글 인증 후 돌아왔는지 확인 (URL 파라미터 체크)
    # st.query_params는 딕셔너리처럼 작동하므로 .get() 사용
    auth_code = st.query_params.get("code")
    
    if auth_code:
        st.write("🔍 구글 인증 코드를 감지했습니다. 토큰 교환을 시작합니다...")
        
        # 돌아왔을 때 safe_key가 없으면 세션이나 URL(state)에서 복구 시도
        if not safe_key:
            safe_key = st.session_state.get('auth_email_key')
        
        # 만약 여전히 없다면 강제 테스트용 (본인 이메일로 수정 가능)
        if not safe_key:
            safe_key = "skyeloveillustration@gmail_com"

        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config(
                {"web": conf}, 
                scopes=SCOPES, 
                redirect_uri=conf["redirect_uri"]
            )
            flow.fetch_token(code=auth_code)
            
            # 저장 로직 실행
            clean_key = safe_key.replace('.', '_')
            ref = db.reference(f'google_calendar_creds/{clean_key}')
            ref.set({'creds': flow.credentials.to_json()})
            
            st.success(f"✅ [{clean_key}] 저장 완료! 리프레시 중...")
            
            # 쿼리 파라미터 삭제 후 리런 (무한 루프 방지)
            st.query_params.clear()
            st.rerun()
            return None
        except Exception as e:
            st.error(f"❌ 토큰 교환 중 에러 발생: {e}")
            st.stop()

    # [2] 이미 연동된 정보가 있는지 확인
    if safe_key:
        st.session_state['auth_email_key'] = safe_key # 유실 대비 세션 저장
        clean_key = safe_key.replace('.', '_')
        data = db.reference(f'google_calendar_creds/{clean_key}').get()
        
        if data and 'creds' in data:
            try:
                creds = Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
                if creds.valid:
                    return build('calendar', 'v3', credentials=creds)
                elif creds.refresh_token:
                    creds.refresh(Request())
                    # 갱신된 토큰 다시 저장
                    db.reference(f'google_calendar_creds/{clean_key}').update({'creds': creds.to_json()})
                    return build('calendar', 'v3', credentials=creds)
            except:
                pass

    # [3] 아무것도 해당 안 되면 인증 버튼 노출
    conf = dict(st.secrets["google_calendar"])
    flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
    auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
    
    st.info("📅 구글 캘린더 연동이 필요합니다.")
    st.markdown(f"**[🔗 구글 계정 연동하기]({auth_url})**")
    return None
