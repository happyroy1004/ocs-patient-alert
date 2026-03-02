import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
import pickle
import time
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# [span_8](start_span)범위 설정[span_8](end_span)
SCOPES = ["https://www.googleapis.com/auth/calendar.events"]

@st.cache_resource
def get_db_refs():
    if not firebase_admin._apps:
        try:
            creds_dict = dict(st.secrets["firebase"])
            db_url = st.secrets["database_url"]
            cred = credentials.Certificate(creds_dict)
            firebase_admin.initialize_app(cred, {'databaseURL': db_url})
        except Exception:
            return None, None, None
    base_ref = db.reference()
    return base_ref.child('users'), base_ref.child('doctor_users'), lambda path: base_ref.child(path)

def sanitize_path(email):
    return email.replace(".", "_dot_").replace("@", "_at_")

def save_google_creds_to_firebase(user_id_safe, creds):
    try:
        # [span_9](start_span)과거 성공 경로로 고정[span_9](end_span)
        ref = db.reference(f"users/{user_id_safe}/google_creds")
        ref.set({'creds_hex': pickle.dumps(creds).hex()})
        return True
    except:
        return False

def load_google_creds_from_firebase(user_id_safe):
    try:
        ref = db.reference(f"users/{user_id_safe}/google_creds")
        data = ref.get()
        if data and 'creds_hex' in data:
            return pickle.loads(bytes.fromhex(data['creds_hex']))
    except:
        return None
    return None

def get_google_calendar_service(user_id_safe):
    session_key = f"google_creds_{user_id_safe}"
    creds = st.session_state.get(session_key)
    
    if not creds:
        creds = load_google_creds_from_firebase(user_id_safe)
        if creds:
            st.session_state[session_key] = creds

    conf = st.secrets["google_calendar"]
    client_config = {
        "web": {
            "client_id": conf["client_id"],
            "client_secret": conf["client_secret"],
            "redirect_uris": [conf["redirect_uri"]],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token"
        }
    }

    if creds:
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                save_google_creds_to_firebase(user_id_safe, creds)
                st.session_state[session_key] = creds
            except: pass
        if creds.valid:
            return build('calendar', 'v3', credentials=creds)

    # [span_10](start_span)리디렉션 처리[span_10](end_span)
    auth_code = st.query_params.get("code")
    if auth_code and 'auth_flow' in st.session_state:
        try:
            st.session_state.auth_flow.fetch_token(code=auth_code)
            new_creds = st.session_state.auth_flow.credentials
            if save_google_creds_to_firebase(user_id_safe, new_creds):
                st.session_state[session_key] = new_creds
                st.success("✅ 인증 완료!")
                st.query_params.clear()
                del st.session_state.auth_flow
                time.sleep(1)
                st.rerun()
        except:
            st.query_params.clear()

    if 'auth_flow' not in st.session_state:
        st.session_state.auth_flow = Flow.from_client_config(
            client_config, scopes=SCOPES, redirect_uri=conf["redirect_uri"]
        )

    auth_url, _ = st.session_state.auth_flow.authorization_url(prompt='consent', access_type='offline')
    
    st.warning("📅 인증이 필요합니다.")
    st.markdown(f"**[🔗 구글 인증 링크]({auth_url})**")
    return None
