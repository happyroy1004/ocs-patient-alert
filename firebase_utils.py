import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json

# 권한 범위 (URL에 찍힌 scope와 정확히 일치)
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

def save_google_creds_to_firebase(safe_key, creds):
    """Firebase에 JSON 문자열로 저장"""
    try:
        ref = db.reference(f'google_calendar_creds/{safe_key}')
        ref.set({'creds': creds.to_json()})
        st.success(f"✅ [{safe_key}] 계정 연동 성공! 데이터가 저장되었습니다.")
    except Exception as e:
        st.error(f"❌ Firebase 저장 실패 (권한 문제일 수 있음): {e}")

def get_google_calendar_service(safe_key=None):
    # 1. URL에서 직접 이메일(state)과 인증코드(code)를 추출
    # 구글에서 돌아올 때 state에 이메일을 담아 보낼 예정입니다.
    returned_state = st.query_params.get("state")
    auth_code = st.query_params.get("code")
    
    # 현재 작업 대상 이메일 결정 (인자값 우선 -> URL 리턴값 순)
    active_key = safe_key if safe_key else returned_state

    # 2. 이미 데이터가 있는지 확인 (로그인 유지용)
    if active_key:
        try:
            data = db.reference(f'google_calendar_creds/{active_key}').get()
            if data and 'creds' in data:
                creds = Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
                if creds.valid:
                    return build('calendar', 'v3', credentials=creds)
                elif creds.refresh_token:
                    creds.refresh(Request())
                    save_google_creds_to_firebase(active_key, creds)
                    return build('calendar', 'v3', credentials=creds)
        except:
            pass

    # 3. 구글에서 인증 코드를 가지고 돌아온 경우 처리
    if auth_code:
        try:
            conf = dict(st.secrets["google_calendar"])
            flow = Flow.from_client_config(
                {"web": conf}, 
                scopes=SCOPES, 
                redirect_uri=conf["redirect_uri"]
            )
            flow.fetch_token(code=auth_code)
            
            # [중요] 구글이 돌려준 state(이메일)를 사용하여 저장
            # 만약 returned_state가 난수라면, skyeloveillustration@gmail.com을 직접 입력해서 테스트해보세요.
            target_key = active_key if "@" in str(active_key) else "skyeloveillustration@gmail_com"
            save_google_creds_to_firebase(target_key, flow.credentials)
            
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"❌ 토큰 교환 실패: {e}")
            st.stop()

    # 4. 연동 버튼 생성 (이메일을 state에 담아서 보냄)
    if active_key:
        conf = dict(st.secrets["google_calendar"])
        flow = Flow.from_client_config({"web": conf}, scopes=SCOPES, redirect_uri=conf["redirect_uri"])
        
        # state에 이메일을 넣어서 보내면, 구글이 인증 후 이 값을 그대로 돌려줍니다.
        auth_url, _ = flow.authorization_url(
            prompt='consent', 
            access_type='offline',
            state=active_key  
        )
        
        st.info(f"📅 [{active_key}] 계정의 구글 캘린더 연동이 필요합니다.")
        st.markdown(f"**[🔗 구글 계정 연동하기]({auth_url})**")
    
    return None
