# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, db
import io
import msoffcrypto
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from openpyxl import load_workbook
from openpyxl.styles import Font
import re
import json
import os
import time
import datetime
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# --- 이메일 유효성 검사 함수 ---
def is_valid_email(email):
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(email_regex, email) is not None

# Firebase 초기화
if not firebase_admin._apps:
    try:
        # Streamlit secrets에서 Firebase 설정 가져오기
        firebase_credentials_json_str = st.secrets["firebase"]["FIREBASE_SERVICE_ACCOUNT_JSON"]
        firebase_credentials_dict = json.loads(firebase_credentials_json_str)

        cred = credentials.Certificate(firebase_credentials_dict)
        firebase_admin.initialize_app(cred, {
            'databaseURL': st.secrets["firebase"]["database_url"]
        })
    except Exception as e:
        st.error(f"Firebase 초기화 오류: {e}")
        st.info("secrets.toml 파일의 Firebase 설정(FIREBASE_SERVICE_ACCOUNT_JSON 또는 database_url)을 [firebase] 섹션 아래에 올바르게 작성했는지 확인해주세요.")
        st.stop()

# Firebase-safe 경로 변환 (이메일을 Firebase 키로 사용하기 위해)
def sanitize_path(email):
    return email.replace(".", "_dot_").replace("@", "_at_")

# 이메일 주소 복원 (Firebase 안전 키에서 원래 이메일로)
def recover_email(safe_id: str) -> str:
    email = safe_id.replace("_at_", "@").replace("_dot_", ".").replace("_com", ".com")
    return email

# --- Excel 파일 처리 관련 함수 ---
def decrypt_and_read_excel(file, password):
    """
    암호화된 엑셀 파일을 복호화하여 pandas DataFrame으로 읽어옵니다.
    """
    try:
        decrypted_file = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(file)
        office_file.decrypt(decrypted_file, password=password)
        df = pd.read_excel(decrypted_file, engine='openpyxl')
        return df
    except msoffcrypto.exceptions.InvalidKeyError:
        st.error("잘못된 비밀번호입니다.")
        return None
    except Exception as e:
        st.error(f"파일 처리 중 오류 발생: {e}")
        return None

def find_pid_in_dataframe(df, pid_list):
    """
    DataFrame에서 등록된 환자(pid) 정보를 찾아 반환합니다.
    """
    pid_column = '진료번호' # 엑셀 파일의 '진료번호' 컬럼
    # 엑셀 파일에 '진료번호' 컬럼이 있는지 확인
    if pid_column not in df.columns:
        st.error("엑셀 파일에 '진료번호' 컬럼이 없습니다.")
        return pd.DataFrame()
    
    # 엑셀 파일의 '진료번호' 컬럼을 문자열로 변환하여 비교
    df[pid_column] = df[pid_column].astype(str).str.strip()
    pid_in_df = df[df[pid_column].isin(pid_list)]
    
    return pid_in_df

# --- 이메일 전송 함수 ---
def send_email(to_email, subject, body):
    """
    Gmail SMTP 서버를 사용하여 이메일을 전송합니다.
    """
    try:
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        sender_email = st.secrets["gmail"]["email"]
        sender_password = st.secrets["gmail"]["password"]

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        return True
    except Exception as e:
        st.error(f"이메일 전송 실패: {e}")
        return False

# --- Google Calendar API 관련 설정 및 함수 ---
# secrets.toml 파일에서 Google Calendar API 설정 가져오기
try:
    client_id = st.secrets["googlecalendar"]["client_id"]
    client_secret = st.secrets["googlecalendar"]["client_secret"]
    redirect_uri = st.secrets["googlecalendar"]["redirect_uri"]
except KeyError:
    st.error("`secrets.toml` 파일에 Google Calendar 설정이 누락되었습니다. 파일을 확인해 주세요.")
    st.stop()

SCOPES = ['https://www.googleapis.com/auth/calendar.events']

def get_google_calendar_service(refresh_token=None):
    """
    Google Calendar API 서비스 객체를 반환합니다.
    사용자의 refresh token이 있으면 이를 사용하고, 없으면 새로운 인증 절차를 시작합니다.
    """
    creds = None
    if refresh_token:
        creds = Credentials(
            token=None,
            refresh_token=refresh_token,
            token_uri="https://oauth2.googleapis.com/token",
            client_id=client_id,
            client_secret=client_secret
        )
    
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_config(
            {
                "installed": {
                    "client_id": client_id,
                    "client_secret": client_secret,
                    "redirect_uris": [redirect_uri],
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token"
                }
            },
            SCOPES
        )
        authorization_url, _ = flow.authorization_url(prompt='consent')
        st.markdown(f"[Google 계정 연동하기]({authorization_url})")
        auth_code = st.text_input("위 링크를 클릭하여 인증 코드를 입력하세요.")
        if auth_code:
            try:
                flow.fetch_token(code=auth_code)
                creds = flow.credentials
                # Refresh token 저장
                st.session_state["google_refresh_token"] = creds.refresh_token
                st.success("Google Calendar에 성공적으로 연동되었습니다.")
                st.experimental_rerun()
            except Exception as e:
                st.error(f"인증 실패: {e}")
        else:
            return None

    if creds and creds.valid:
        return build('calendar', 'v3', credentials=creds)
    return None

def create_event(service, start_time, end_time, summary, description):
    """
    Google Calendar에 이벤트를 생성합니다.
    """
    event = {
        'summary': summary,
        'description': description,
        'start': {
            'dateTime': start_time.isoformat(),
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': end_time.isoformat(),
            'timeZone': 'Asia/Seoul',
        },
    }
    try:
        event = service.events().insert(calendarId='primary', body=event).execute()
        st.success(f"Google Calendar에 이벤트가 생성되었습니다: {event.get('htmlLink')}")
    except Exception as e:
        st.error(f"이벤트 생성 실패: {e}")

# --- 엑셀 처리 관련 상수 및 함수 ---
sheet_keyword_to_department_map = {
    '치과보철과': '보철', '보철과': '보철', '보철': '보철',
    '치과교정과' : '교정', '교정과': '교정', '교정': '교정',
    '구강 악안면외과' : '외과', '구강악안면외과': '외과', '외과': '외과',
    '구강 내과' : '내과', '구강내과': '내과', '내과': '내과',
    '치과보존과' : '보존', '보존과': '보존', '보존': '보존',
    '소아치과': '소치', '소치': '소치', '소아 치과': '소치',
    '원내생진료센터': '원내생', '원내생': '원내생','원내생 진료센터': '원내생','원진실':'원내생',
    '원스톱 협진센터' : '원스톱', '원스톱협진센터': '원스톱', '원스톱': '원스톱',
    '임플란트 진료센터' : '임플란트', '임플란트진료센터': '임플란트', '임플란트': '임플란트',
    '임플' : '임플란트', '치주과': '치주', '치주': '치주',
    '임플실': '임플란트', '원진실': '원내생', '병리': '병리'
}

professors_dict = {
    '소치': ['김현태', '장기택', '김정욱', '현홍근', '김영재', '신터전', '송지수'],
    '보존': ['이인복', '금기연', '이우철', '유연지', '서덕규', '이창하', '김선영', '손원준'],
    '외과': ['최진영', '서병무', '명훈', '김성민', '박주영', '양훈주', '한정준', '권익재'],
    '치주': ['구영', '이용무', '설양조', '구기태', '김성태', '조영단'],
    '보철': ['곽재영', '김성균', '임영준', '김명주', '권호범', '여인성', '윤형인', '박지만', '이재현', '조준호'],
    '교정': [], '내과': [], '원내생': [], '원스톱': [], '임플란트': [], '병리': []
}

def process_excel_file_and_style(file_bytes_io):
    """
    업로드된 엑셀 파일을 처리하고 스타일을 적용하여 새 파일을 반환합니다.
    """
    file_bytes_io.seek(0)
    try:
        wb_raw = load_workbook(filename=file_bytes_io, keep_vba=False, data_only=True)
    except Exception as e:
        raise ValueError(f"엑셀 워크북 로드 실패: {e}")

    processed_sheets_dfs = {}
    for sheet_name_raw in wb_raw.sheetnames:
        sheet_name_lower = sheet_name_raw.strip().lower()
        sheet_key = None
        for keyword, department_name in sorted(sheet_keyword_to_department_map.items(), key=lambda item: len(item[0]), reverse=True):
            if keyword.lower() in sheet_name_lower:
                sheet_key = department_name
                break
        if not sheet_key:
            st.warning(f"시트 '{sheet_name_raw}'을(를) 인식할 수 없습니다. 건너킵니다.")
            continue
        ws = wb_raw[sheet_name_raw]
        values = list(ws.values)
        while values and (values[0] is None or all((v is None or str(v).strip() == "") for v in values[0])):
            values.pop(0)
        if len(values) < 2:
            st.warning(f"시트 '{sheet_name_raw}'에 유효한 데이터가 충분하지 않습니다. 건너깁니다.")
            continue
        df = pd.DataFrame(values)
        df.columns = df.iloc[0]
        df = df.drop([0]).reset_index(drop=True)
        df = df.fillna("").astype(str)
        if '예약의사' in df.columns:
            df['예약의사'] = df['예약의사'].str.strip().str.replace(" 교수님", "", regex=False)
        else:
            st.warning(f"시트 '{sheet_name_raw}': '예약의사' 컬럼이 없습니다. 이 시트는 처리되지 않습니다.")
            continue
        professors_list = professors_dict.get(sheet_key, [])
        try:
            # `process_sheet_v8` 함수는 이 예제 코드에 포함되지 않아, 간단한 정렬만 적용
            processed_df = df.sort_values(by=['예약의사', '예약시간'], ascending=[True, True])
            processed_sheets_dfs[sheet_name_raw] = processed_df
        except KeyError as e:
            st.error(f"시트 '{sheet_name_raw}' 처리 중 컬럼 오류: {e}. 이 시트는 건너깁니다.")
            continue
        except Exception as e:
            st.error(f"시트 '{sheet_name_raw}' 처리 중 알 수 없는 오류: {e}. 이 시트는 건너깁니다.")
            continue

    if not processed_sheets_dfs:
        st.info("처리된 시트가 없습니다.")
        return None, None
    output_buffer_for_styling = io.BytesIO()
    with pd.ExcelWriter(output_buffer_for_styling, engine='openpyxl') as writer:
        for sheet_name_raw, df in processed_sheets_dfs.items():
            df.to_excel(writer, sheet_name=sheet_name_raw, index=False)
    output_buffer_for_styling.seek(0)
    wb_styled = load_workbook(output_buffer_for_styling, keep_vba=False, data_only=True)
    # 스타일링 로직 (코드레용.txt의 내용 기반)
    for sheet_name in wb_styled.sheetnames:
        ws = wb_styled[sheet_name]
        if sheet_name.strip() == "교정":
            header = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}
            if '진료내역' in header:
                idx = header['진료내역'] - 1
                for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
                    if len(row) > idx:
                        cell = row[idx]
                        text = str(cell.value).strip().lower()
                        if ('bonding' in text or '본딩' in text) and 'debonding' not in text:
                            cell.font = Font(bold=True)
    final_output_bytes = io.BytesIO()
    wb_styled.save(final_output_bytes)
    final_output_bytes.seek(0)
    return processed_sheets_dfs, final_output_bytes

# --- Streamlit UI ---
st.set_page_config(layout="wide")
st.title("OCS 환자 알림 시스템")
st.caption("환자 정보 관리 및 진료 알림을 위한 앱입니다.")
st.markdown("---")

# --- 세션 상태 초기화 ---
if 'logged_in_as_admin' not in st.session_state:
    st.session_state.logged_in_as_admin = False

# --- Firebase에서 환자 데이터 가져오기 ---
ref = db.reference('/')
patients_ref_for_user = ref.child('patients')
patients_data = patients_ref_for_user.get()
existing_patient_data = patients_data if patients_data else {}
existing_pids = list(val['진료번호'] for val in existing_patient_data.values() if '진료번호' in val)

# 사용자 이름 입력 필드
user_name = st.text_input("사용자 이름을 입력하세요 (예시: 홍길동 또는 admin)")

# Admin 계정 확인 로직
is_admin_input = (user_name.strip().lower() == "admin")
if is_admin_input:
    admin_password = st.text_input("관리자 비밀번호를 입력하세요", type="password")
    if admin_password == st.secrets["admin"]["password"]:
        st.session_state.logged_in_as_admin = True
        st.info("관리자 계정으로 로그인했습니다.")
    else:
        st.error("비밀번호가 올바르지 않습니다.")
        st.session_state.logged_in_as_admin = False

elif user_name:
    st.session_state.logged_in_as_admin = False
    st.info(f"**{user_name}** 님으로 로그인되었습니다.")
elif not user_name:
    st.warning("사용자 이름을 입력해주세요.")
    st.stop()

tab1, tab2, tab3 = st.tabs(["환자 관리", "환자 상태 확인 및 알림", "이메일 알림"])

with tab1:
    st.header("환자 등록/삭제")
    st.info("여기에 환자를 등록하거나 삭제할 수 있습니다.")

    with st.container(border=True):
        st.markdown("### 등록된 환자 목록")
        if existing_patient_data:
            for key, val in existing_patient_data.items():
                col1, col2 = st.columns([0.9, 0.1])
                with col1:
                    st.markdown(f"**{val.get('환자명', '미지정')}** / {val.get('진료번호', '미지정')} / {val.get('등록과', '미지정')}")
                with col2:
                    if st.button("X", key=f"delete_button_{key}"):
                        patients_ref_for_user.child(key).delete()
                        st.experimental_rerun()
        else:
            st.info("등록된 환자가 없습니다.")
    
    st.markdown("---")

    with st.form("register_form"):
        st.markdown("### 신규 환자 등록")
        name = st.text_input("환자명")
        pid = st.text_input("진료번호")
        selected_department = st.selectbox("등록 과", ["내과", "외과", "소아과", "미지정"])
        submitted = st.form_submit_button("등록")

        if submitted:
            if not name or not pid:
                st.warning("모든 항목을 입력해주세요.")
            elif any(v.get("진료번호") == pid for v in existing_patient_data.values()):
                st.error("이미 등록된 진료번호입니다.")
            else:
                patients_ref_for_user.push().set({"환자명": name, "진료번호": pid, "등록과": selected_department})
                st.success(f"환자 '{name}'이(가) 등록되었습니다.")
                time.sleep(1)
                st.experimental_rerun()

with tab2:
    st.header("환자 상태 확인 및 알림")
    st.info("엑셀 파일에서 등록된 환자의 상태를 확인하고, 진료 일정을 캘린더에 추가할 수 있습니다.")

    uploaded_file = st.file_uploader("보호된 엑셀 파일(.xlsx)을 업로드하세요", type="xlsx")
    if uploaded_file:
        password = st.text_input("엑셀 파일 비밀번호를 입력하세요", type="password")
        if password:
            df = decrypt_and_read_excel(uploaded_file, password)
            if df is not None:
                st.write("### 업로드된 엑셀 파일 미리보기")
                st.dataframe(df.head())

                st.write("### 등록된 환자 진료 상태")
                pid_list = existing_pids
                if not pid_list:
                    st.warning("등록된 환자가 없습니다. '환자 관리' 탭에서 환자를 등록해주세요.")
                else:
                    found_patients_df = find_pid_in_dataframe(df, pid_list)
                    if not found_patients_df.empty:
                        st.dataframe(found_patients_df)
                    else:
                        st.info("업로드된 엑셀 파일에서 등록된 환자 정보를 찾을 수 없습니다.")
    
    st.markdown("---")
    
    st.markdown("### Google Calendar 연동")
    google_refresh_token = st.session_state.get("google_refresh_token", None)

    if st.session_state.logged_in_as_admin:
        # 관리자용 기능
        if not google_refresh_token:
            if st.button("Google Calendar 연동 시작 (관리자용)"):
                service = get_google_calendar_service(refresh_token=None)
                if service:
                    st.session_state["google_refresh_token"] = service._http.credentials.refresh_token
                    st.success("Google Calendar에 성공적으로 연동되었습니다.")
                    st.experimental_rerun()
        else:
            st.success("Google Calendar와 연동되어 있습니다.")
            with st.form("calendar_event_form"):
                st.subheader("새로운 진료 이벤트 추가")
                event_summary = st.text_input("이벤트 제목", "환자 진료 일정")
                event_description = st.text_area("이벤트 설명", "진료 관련 내용")
                event_start_date = st.date_input("시작 날짜", datetime.date.today())
                event_start_time = st.time_input("시작 시간", datetime.time(9, 0))
                event_end_date = st.date_input("종료 날짜", datetime.date.today())
                event_end_time = st.time_input("종료 시간", datetime.time(10, 0))

                submitted_event = st.form_submit_button("캘린더에 추가")
                if submitted_event:
                    try:
                        start_datetime = datetime.datetime.combine(event_start_date, event_start_time)
                        end_datetime = datetime.datetime.combine(event_end_date, event_end_time)
                        if start_datetime >= end_datetime:
                            st.error("종료 시간이 시작 시간보다 빠를 수 없습니다.")
                        else:
                            service = get_google_calendar_service(refresh_token=google_refresh_token)
                            if service:
                                create_event(service, start_datetime, end_datetime, event_summary, event_description)
                    except Exception as e:
                        st.error(f"날짜/시간 입력 오류: {e}")
    else:
        # 일반 사용자용 기능
        if not google_refresh_token:
            st.info("Google Calendar를 연동하여 진료 일정을 확인할 수 있습니다.")
            if st.button("Google Calendar 연동 시작"):
                # 일반 사용자는 연동만 가능하도록 함
                service = get_google_calendar_service(refresh_token=None)
                if service:
                    st.session_state["google_refresh_token"] = service._http.credentials.refresh_token
                    st.success("Google Calendar에 성공적으로 연동되었습니다.")
                    st.experimental_rerun()
        else:
            st.success("Google Calendar와 연동되어 있습니다.")
            # 일반 사용자는 이벤트를 추가할 수 없음
            st.info("관리자만 일정을 추가할 수 있습니다.")


with tab3:
    st.header("이메일 알림")
    st.info("환자 상태에 대한 이메일 알림을 보낼 수 있습니다.")
    
    with st.form("email_form"):
        st.subheader("이메일 보내기")
        to_email = st.text_input("수신자 이메일 주소")
        subject = st.text_input("제목", "OCS 환자 알림")
        body = st.text_area("내용", "진료 상태 확인 부탁드립니다.")

        submitted_email = st.form_submit_button("이메일 전송")
        if submitted_email:
            if not is_valid_email(to_email):
                st.error("유효한 이메일 주소를 입력해주세요.")
            else:
                if send_email(to_email, subject, body):
                    st.success("이메일이 성공적으로 발송되었습니다.")
