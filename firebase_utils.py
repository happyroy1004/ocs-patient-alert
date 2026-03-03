import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google_auth_oauthlib.flow import Flow 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials  # 이 부분이 누락되었을 확률이 높습니다
from googleapiclient.discovery import build
import json

# 스코프 설정: GCP 콘솔 설정과 반드시 일치해야 합니다.
# 만약 'userinfo.email' 추가가 어려우시면 일단 calendar.events만 남겨두고 테스트하세요.
SCOPES = [
    'https://www.googleapis.com/auth/calendar.events',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid'
]

def get_google_calendar_service(safe_key=None):
    # [1] URL에 'code'가 보이면 무조건 가로채서 저장 실행
    auth_code = st.query_params.get("code")

    if auth_code:
        try:
            conf = dict(st.secrets["google_calendar"])
            # Flow 생성
            flow = Flow.from_client_config(
                {"web": conf}, 
                scopes=SCOPES, 
                redirect_uri=conf["redirect_uri"]
            )
            flow.fetch_token(code=auth_code)
            
            # [테스트용 강제 설정] 이메일 유실 방지를 위해 본인 계정으로 직접 경로 지정
            # safe_key가 없으면 기본값으로 skyelove...를 사용합니다.
            target_key = safe_key if safe_key else "skyeloveillustration@gmail.com"
            clean_key = target_key.replace('.', '_')
            
            ref = db.reference(f'google_calendar_creds/{clean_key}')
            ref.set({'creds': flow.credentials.to_json()})
            
            st.success(f"✅ DB에 데이터가 저장되었습니다! (계정: {target_key})")
            
            # URL 정리 및 리런 (무한 루프 방지)
            st.query_params.clear()
            st.rerun()
            return None
        except Exception as e:
            st.error(f"❌ 토큰 교환 중 오류 발생: {e}")
            st.info("💡 팁: 구글 클라우드 콘솔의 'OAuth 동의 화면'에서 스코프가 정확히 추가되었는지 확인하세요.")
            st.stop()

    # [2] 평상시 데이터 불러오기 로직
    # 저장된 키가 있는지 확인
    target_key = safe_key if safe_key else "skyeloveillustration@gmail.com"
    clean_key = target_key.replace('.', '_')
    
    data = db.reference(f'google_calendar_creds/{clean_key}').get()
    
    if data and 'creds' in data:
        try:
            # 여기서 Credentials를 사용하여 객체 복원
            creds = Credentials.from_authorized_user_info(json.loads(data['creds']), SCOPES)
            
            # 토큰 만료 시 갱신 시도
            if not creds.valid and creds.refresh_token:
                creds.refresh(Request())
                db.reference(f'google_calendar_creds/{clean_key}').update({'creds': creds.to_json()})
                
            return build('calendar', 'v3', credentials=creds)
        except Exception as e:
            st.warning(f"⚠️ 기존 인증 정보 로드 실패: {e}")

    # [3] 인증 버튼 표시 (인증 정보가 없을 때만)
    try:
        conf = dict(st.secrets["google_calendar"])
        flow = Flow.from_client_config(
            {"web": conf}, 
            scopes=SCOPES, 
            redirect_uri=conf["redirect_uri"]
        )
        auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
        
        st.info("📅 구글 캘린더 연동이 필요합니다.")
        st.markdown(f'<a href="{auth_url}" target="_self" style="text-decoration:none;"><div style="background-color:#4285F4; color:white; padding:10px; border-radius:5px; text-align:center;">구글 계정 연동하기</div></a>', unsafe_allow_html=True)
    except Exception as e:
        st.error(f"⚠️ 설정 오류: {e}")
        
    return None
