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

# Google Calendar API 관련 라이브러리 추가
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import datetime
import base64

# --- 이메일 유효성 검사 함수 ---
def is_valid_email(email):
    """
    이메일 주소의 유효성을 검사합니다.
    """
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$"
    return re.match(email_regex, email) is not None

# Firebase 초기화
if not firebase_admin._apps:
    try:
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
    """
    이메일 주소를 Firebase Realtime Database 키로 사용할 수 있도록 변환합니다.
    """
    return email.replace(".", "_dot_").replace("@", "_at_")

# 이메일 주소 복원 (Firebase 안전 키에서 원래 이메일로)
def recover_email(safe_id: str) -> str:
    """
    Firebase Realtime Database 키를 원래 이메일 주소로 복원합니다.
    """
    email = safe_id.replace("_at_", "@").replace("_dot_", ".").replace("_com", ".com")
    return email

# 암호화된 엑셀 파일인지 확인
def is_encrypted_excel(file):
    """
    파일이 MS Office 암호화 파일인지 확인합니다.
    """
    try:
        file.seek(0)
        return msoffcrypto.OfficeFile(file).is_encrypted()
    except Exception:
        return False

# 엑셀 파일 로드 및 복호화
def load_excel(file, password=None):
    """
    암호화된 엑셀 파일을 복호화하여 Pandas ExcelFile 객체를 반환합니다.
    """
    try:
        file.seek(0)
        office_file = msoffcrypto.OfficeFile(file)
        if office_file.is_encrypted():
            if not password:
                raise ValueError("암호화된 파일입니다. 비밀번호를 입력해주세요.")
            decrypted = io.BytesIO()
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            return pd.ExcelFile(decrypted), decrypted
        else:
            return pd.ExcelFile(file), file
    except Exception as e:
        raise ValueError(f"엑셀 로드 또는 복호화 실패: {e}")

# 이메일 전송 함수
def send_email(receiver, rows, sender, password, date_str=None, custom_message=None):
    """
    HTML 형식으로 데이터를 포함한 이메일을 보냅니다.
    """
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver

        if custom_message:
            msg['Subject'] = "단체 메일 알림"
            body = custom_message
        else:
            subject_prefix = ""
            if date_str:
                subject_prefix = f"{date_str}일에 내원하는 "
            msg['Subject'] = f"{subject_prefix}등록 환자 내원 알림"
            
            html_table = rows.to_html(index=False, escape=False)
            
            style = """
            <style>
                table {
                    width: 100%;
                    max-width: 100%;
                    border-collapse: collapse;
                    font-family: Arial, sans-serif;
                    font-size: 14px;
                    table-layout: fixed;
                }
                th, td {
                    border: 1px solid #dddddd;
                    text-align: left;
                    padding: 8px;
                    vertical-align: top;
                    word-wrap: break-word;
                    word-break: break-word;
                }
                th {
                    background-color: #f2f2f2;
                    font-weight: bold;
                    white-space: nowrap;
                }
                tr:nth-child(even) {
                    background-color: #f9f9f9;
                }
                .table-container {
                    overflow-x: auto;
                    -webkit-overflow-scrolling: touch;
                }
            </style>
            """
            body = f"다음 토탈 환자가 내일 내원예정입니다:<br><br><div class='table-container'>{style}{html_table}</div>"
        
        msg.attach(MIMEText(body, 'html'))
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        return str(e)

# --- Google Calendar API 관련 함수 ---
# 사용할 스코프 정의. 캘린더 이벤트 생성 권한
SCOPES = ["https://www.googleapis.com/auth/calendar.events"]

def create_calendar_event(service, patient_name, pid, department, event_date_str, event_time_str):
    """
    구글 캘린더에 이벤트를 생성합니다.
    """
    try:
        # 날짜와 시간을 합쳐서 RFC3339 형식으로 변환
        event_datetime_str = f"{event_date_str}T{event_time_str}:00"
        
        event = {
            'summary': f"내원: {patient_name} ({department})",
            'description': f"진료번호: {pid}\n환자명: {patient_name}\n등록과: {department}",
            'start': {
                'dateTime': event_datetime_str,
                'timeZone': 'Asia/Seoul',
            },
            'end': {
                'dateTime': event_datetime_str,
                'timeZone': 'Asia/Seoul',
            },
            'reminders': {
                'useDefault': False,
                'overrides': [
                    {'method': 'email', 'minutes': 24 * 60},  # 1일 전 이메일
                    {'method': 'popup', 'minutes': 30},       # 30분 전 팝업
                ],
            },
        }
        service.events().insert(calendarId='primary', body=event).execute()
        st.success(f"Google Calendar에 {patient_name}님 ({event_date_str} {event_time_str}) 일정이 추가되었습니다.")
    except Exception as e:
        st.error(f"Google Calendar 이벤트 생성 실패: {e}")

def get_google_calendar_service(user_id_safe):
    """
    사용자별로 Google Calendar 서비스 객체를 반환하거나 인증 URL을 표시합니다.
    Streamlit 세션 상태를 활용하여 인증 정보를 관리합니다.
    """
    creds = st.session_state.get(f"google_creds_{user_id_safe}")

    # secrets.toml에서 클라이언트 설정 불러오기
    client_config = {
        "web": {
            "client_id": st.secrets["google_calendar"]["client_id"],
            "client_secret": st.secrets["google_calendar"]["client_secret"],
            "redirect_uris": [st.secrets["google_calendar"]["redirect_uri"]],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs"
        }
    }
    
    # 인증 플로우 생성
    flow = InstalledAppFlow.from_client_config(client_config, SCOPES, redirect_uri=st.secrets["google_calendar"]["redirect_uri"])
    
    if not creds:
        auth_code = st.query_params.get("code")
        
        if auth_code:
            # 인증 코드를 사용하여 토큰을 교환
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            st.session_state[f"google_creds_{user_id_safe}"] = creds
            st.success("Google Calendar 인증이 완료되었습니다.")
            st.query_params.clear()
            st.rerun()
        else:
            auth_url, _ = flow.authorization_url(prompt='consent')
            st.warning("Google Calendar 연동을 위해 인증이 필요합니다. 아래 링크를 클릭하여 권한을 부여하세요.")
            st.markdown(f"**[Google Calendar 인증 링크]({auth_url})**", unsafe_allow_html=True)
            return None

    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        st.session_state[f"google_creds_{user_id_safe}"] = creds

    try:
        service = build('calendar', 'v3', credentials=creds)
        return service
    except HttpError as error:
        st.error(f'Google Calendar 서비스 생성 실패: {error}')
        st.session_state.pop(f"google_creds_{user_id_safe}", None)
        return None

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

# 엑셀 시트 데이터 처리 (교수님/비교수님, 시간/의사별 정렬)
def process_sheet_v8(df, professors_list, sheet_key):
    """
    엑셀 시트의 데이터를 교수님/비교수님으로 분리하고, 예약시간/의사별로 정렬합니다.
    """
    df = df.drop(columns=['예약일시'], errors='ignore')
    if '예약의사' not in df.columns or '예약시간' not in df.columns:
        st.error(f"시트 처리 오류: '예약의사' 또는 '예약시간' 컬럼이 DataFrame에 없습니다.")
        return pd.DataFrame(columns=['진료번호', '예약시간', '환자명', '예약의사', '진료내역'])

    df = df.sort_values(by=['예약의사', '예약시간'])
    professors = df[df['예약의사'].isin(professors_list)]
    non_professors = df[~df['예약의사'].isin(professors_list)]

    if sheet_key != '보철':
        non_professors = non_professors.sort_values(by=['예약시간', '예약의사'])
    else:
        non_professors = non_professors.sort_values(by=['예약의사', '예약시간'])

    final_rows = []
    current_time = None
    current_doctor = None

    for _, row in non_professors.iterrows():
        if sheet_key != '보철':
            if current_time != row['예약시간']:
                if current_time is not None:
                    final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
                current_time = row['예약시간']
        else:
            if current_doctor != row['예약의사']:
                if current_doctor is not None:
                    final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
                current_doctor = row['예약의사']
        final_rows.append(row)

    final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
    final_rows.append(pd.Series(["<교수님>"] + [" "] * (len(df.columns) - 1), index=df.columns))

    current_professor = None
    for _, row in professors.iterrows():
        if current_professor != row['예약의사']:
            if current_professor is not None:
                final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
            current_professor = row['예약의사']
        final_rows.append(row)

    final_df = pd.DataFrame(final_rows, columns=df.columns)
    required_cols = ['진료번호', '예약시간', '환자명', '예약의사', '진료내역']
    final_df = final_df[[col for col in required_cols if col in final_df.columns]]
    return final_df

# 엑셀 파일 전체 처리 및 스타일 적용
def process_excel_file_and_style(file_bytes_io):
    """
    업로드된 엑셀 파일을 처리하고, 시트별로 데이터를 정리하고 스타일을 적용합니다.
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
            processed_df = process_sheet_v8(df, professors_list, sheet_key)
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

    for sheet_name in wb_styled.sheetnames:
        ws = wb_styled[sheet_name]
        header = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
            if row[0].value == "<교수님>":
                for cell in row:
                    if cell.value:
                        cell.font = Font(bold=True)
            if sheet_name.strip() == "교정" and '진료내역' in header:
                idx = header['진료내역'] - 1
                if len(row) > idx:
                    cell = row[idx]
                    text = str(cell.value).strip().lower()
                    if ('bonding' in text or '본딩' in text) and 'debonding' not in text:
                        cell.font = Font(bold=True)

    final_output_bytes = io.BytesIO()
    wb_styled.save(final_output_bytes)
    final_output_bytes.seek(0)
    return processed_sheets_dfs, final_output_bytes

# --- Streamlit 애플리케이션 시작 ---
st.set_page_config(layout="wide")

# 제목에 링크 추가 및 초기화 로직
st.markdown("""
<style>
.title-link {
    text-decoration: none;
    color: inherit;
}
</style>
<h1>
    <a href="." class="title-link">환자 내원 확인 시스템</a>
</h1>
""", unsafe_allow_html=True)
st.markdown("---")
st.markdown("<p style='text-align: left; color: grey; font-size: small;'>directed by HSY</p>", unsafe_allow_html=True)

# --- 세션 상태 초기화 ---
if "clear" in st.query_params and st.query_params["clear"] == "true":
    st.session_state.clear()
    st.query_params["clear"] = "false"
    st.rerun()
if 'email_change_mode' not in st.session_state:
    st.session_state.email_change_mode = False
if 'user_id_input_value' not in st.session_state:
    st.session_state.user_id_input_value = ""
if 'found_user_email' not in st.session_state:
    st.session_state.found_user_email = ""
if 'current_firebase_key' not in st.session_state:
    st.session_state.current_firebase_key = ""
if 'current_user_name' not in st.session_state:
    st.session_state.current_user_name = ""
if 'logged_in_as_admin' not in st.session_state:
    st.session_state.logged_in_as_admin = False
if 'admin_password_correct' not in st.session_state:
    st.session_state.admin_password_correct = False
if 'select_all_users' not in st.session_state:
    st.session_state.select_all_users = False
if 'google_calendar_auth_needed' not in st.session_state:
    st.session_state.google_calendar_auth_needed = False
if 'google_creds_validated' not in st.session_state:
    st.session_state.google_creds_validated = False

users_ref = db.reference("users")

# --- 사용 설명서 PDF 다운로드 버튼 추가 ---
pdf_file_path = "manual.pdf"
pdf_display_name = "사용 설명서"
if os.path.exists(pdf_file_path):
    with open(pdf_file_path, "rb") as pdf_file:
        st.download_button(
            label=f"{pdf_display_name} 다운로드",
            data=pdf_file,
            file_name=pdf_file_path,
            mime="application/pdf"
        )
else:
    st.warning(f"⚠️ {pdf_display_name} 파일을 찾을 수 없습니다. (경로: {pdf_file_path})")

# 사용자 이름 입력 필드
user_name = st.text_input("사용자 이름을 입력하세요 (예시: 홍길동)", value=st.session_state.user_id_input_value)

# Admin 계정 확인 로직
is_admin_input = (user_name.strip().lower() == "admin")

# user_name이 입력되었을 때 기존 사용자 검색
if user_name and not is_admin_input and not st.session_state.email_change_mode:
    all_users_meta = users_ref.get()
    matched_users_by_name = []
    if all_users_meta:
        for safe_key, user_info in all_users_meta.items():
            if user_info and user_info.get("name") == user_name:
                matched_users_by_name.append({"safe_key": safe_key, "email": user_info.get("email", ""), "name": user_info.get("name", "")})

    if matched_users_by_name:
        st.session_state.current_user_name = user_name
        st.session_state.current_firebase_key = matched_users_by_name[0]["safe_key"]
        st.session_state.found_user_email = matched_users_by_name[0]["email"]
        st.info(f"'{user_name}'님으로 로그인되었습니다.")
        st.session_state.user_id_input_value = user_name
        st.rerun()
    else:
        st.warning(f"'{user_name}'(으)로 등록된 사용자가 없습니다. 이메일 주소를 입력하여 새로운 사용자로 등록하세요.")
        st.session_state.email_change_mode = True
        st.session_state.user_id_input_value = user_name

# Admin 로그인 로직
if is_admin_input:
    admin_password = st.secrets["admin"]["password"]
    
    if st.session_state.logged_in_as_admin:
        st.success("Admin 계정으로 로그인되었습니다.")
    else:
        admin_input_password = st.text_input("Admin 비밀번호를 입력하세요:", type="password")
        if admin_input_password:
            if admin_input_password == admin_password:
                st.session_state.logged_in_as_admin = True
                st.session_state.admin_password_correct = True
                st.session_state.current_user_name = "Admin"
                st.session_state.user_id_input_value = "Admin"
                st.success("Admin 계정으로 로그인되었습니다.")
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")
                st.session_state.admin_password_correct = False

# 사용자 로그인/등록 UI
if st.session_state.email_change_mode:
    new_user_email = st.text_input("이메일 주소를 입력하고 엔터를 누르세요:", key="email_input")
    
    if new_user_email:
        if is_valid_email(new_user_email):
            safe_key = sanitize_path(new_user_email)
            user_data = users_ref.child(safe_key).get()

            if user_data:
                # 기존 사용자가 이름을 바꾸려고 할 때
                st.error(f"'{new_user_email}'(으)로 이미 등록된 계정이 있습니다. 이름을 변경하려면 다른 이름을 입력하세요.")
                st.session_state.user_id_input_value = ""
                st.session_state.email_change_mode = False
                st.rerun()
            else:
                # 신규 사용자 등록
                new_user_info = {"name": st.session_state.user_id_input_value, "email": new_user_email}
                users_ref.child(safe_key).set(new_user_info)
                st.success(f"새로운 사용자 '{st.session_state.user_id_input_value}'님 ({new_user_email})으로 등록되었습니다. 페이지를 새로고침합니다.")
                st.session_state.email_change_mode = False
                st.session_state.current_firebase_key = safe_key
                st.session_state.found_user_email = new_user_email
                st.rerun()
        else:
            st.error("유효하지 않은 이메일 주소입니다.")

# --- 메인 애플리케이션 로직 ---
if st.session_state.current_firebase_key or st.session_state.logged_in_as_admin:
    st.markdown("---")
    st.subheader(f"✨ {st.session_state.current_user_name}님의 환자 관리 시스템")

    tab1, tab2, tab3 = st.tabs(["엑셀 파일 업로드", "환자 등록", "메일 발송"])

    # 탭 1: 엑셀 파일 업로드 및 처리
    with tab1:
        st.markdown("#### 엑셀 파일 업로드")
        uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요 (xlsx)", type=["xlsx"])
        
        # 파일 업로드 관련 로직
        if uploaded_file:
            st.info("파일 업로드가 완료되었습니다. 엑셀 파일을 처리하고 있습니다...")

            is_encrypted = is_encrypted_excel(uploaded_file)
            password = None
            if is_encrypted:
                password = st.text_input("암호화된 파일입니다. 비밀번호를 입력하세요:", type="password")
                if not password:
                    st.warning("비밀번호를 입력해야 파일을 처리할 수 있습니다.")
                    st.stop()

            try:
                excel_data, decrypted_file_bytes_io = load_excel(uploaded_file, password)
                processed_dfs, styled_excel_bytes = process_excel_file_and_style(decrypted_file_bytes_io)

                if processed_dfs:
                    st.success("✅ 엑셀 파일 처리가 완료되었습니다!")
                    st.write("---")

                    # 처리된 엑셀 파일 다운로드 버튼
                    st.download_button(
                        label="처리된 엑셀 파일 다운로드",
                        data=styled_excel_bytes,
                        file_name=f"처리된_환자_목록_{time.strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.markdown("---")
                    st.subheader("엑셀 파일 미리보기")

                    # 탭을 사용하여 시트별 데이터 미리보기
                    sheet_tabs = st.tabs(processed_dfs.keys())
                    for sheet_name, df in processed_dfs.items():
                        with sheet_tabs[list(processed_dfs.keys()).index(sheet_name)]:
                            st.subheader(f"시트: {sheet_name}")
                            st.dataframe(df, use_container_width=True)

            except ValueError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"예상치 못한 오류가 발생했습니다: {e}")

    # 탭 2: 환자 등록
    with tab2:
        st.markdown("#### 환자 등록")
        if st.session_state.logged_in_as_admin:
            st.info("Admin 계정은 환자를 등록할 수 없습니다. 개별 사용자 계정으로 로그인해주세요.")
            st.stop()
        
        patients_ref_for_user = db.reference(f"patients/{st.session_state.current_firebase_key}")
        existing_patient_data = patients_ref_for_user.get()

        st.markdown("##### 등록된 환자 목록")
        if existing_patient_data:
            cols_header = st.columns([0.8, 0.2])
            with cols_header[0]:
                st.markdown("**환자명 / 진료번호 / 등록 과**")
            with cols_header[1]:
                st.markdown("**삭제**")

            for key, val in existing_patient_data.items():
                cols = st.columns([0.8, 0.2])
                with cols[0]:
                    st.markdown(f"**{val['환자명']}** / {val['진료번호']} / {val.get('등록과', '미지정')}")
                    
                    # 캘린더 이벤트 추가 버튼
                    service = get_google_calendar_service(st.session_state.current_firebase_key)
                    if service:
                        if st.button("캘린더에 일정 추가", key=f"add_calendar_{key}"):
                            # 캘린더 이벤트 생성 함수 호출
                            today_date = datetime.date.today().strftime("%Y-%m-%d")
                            now_time = datetime.datetime.now().strftime("%H:%M")
                            create_calendar_event(service, val['환자명'], val['진료번호'], val.get('등록과', '미지정'), today_date, now_time)
                
                with cols[1]:
                    if st.button("X", key=f"delete_button_{key}"):
                        patients_ref_for_user.child(key).delete()
                        st.rerun()
        else:
            st.info("등록된 환자가 없습니다.")
        st.markdown("---")

        with st.form("register_form"):
            name = st.text_input("환자명")
            pid = st.text_input("진료번호")

            departments_for_registration = sorted(list(set(sheet_keyword_to_department_map.values())))
            selected_department = st.selectbox("등록 과", departments_for_registration)

            submitted = st.form_submit_button("등록")
            if submitted:
                if not name or not pid:
                    st.warning("모든 항목을 입력해주세요.")
                elif existing_patient_data and any(
                    v["환자명"] == name and v["진료번호"] == pid and v.get("등록과") == selected_department
                    for v in existing_patient_data.values()):
                    st.error("이미 등록된 환자입니다.")
                else:
                    patients_ref_for_user.push().set({"환자명": name, "진료번호": pid, "등록과": selected_department})
                    st.success(f"{name} ({pid}) 환자가 등록되었습니다.")
                    st.rerun()

    # 탭 3: 메일 발송
    with tab3:
        st.markdown("#### 메일 발송")
        if st.session_state.logged_in_as_admin:
            st.info("Admin 계정은 메일을 발송할 수 없습니다. 개별 사용자 계정으로 로그인해주세요.")
            st.stop()
        
        email_sender = st.session_state.found_user_email
        if not email_sender:
            st.error("로그인된 사용자의 이메일 정보가 없습니다. 다시 로그인해주세요.")
        else:
            st.info(f"메일 발송 계정: {email_sender}")
            email_password = st.text_input("Gmail 앱 비밀번호를 입력하세요:", type="password")

            email_tab1, email_tab2 = st.tabs(["등록 환자 내원 알림", "단체 메일"])

            # 탭 3-1: 등록 환자 내원 알림
            with email_tab1:
                st.markdown("##### 등록된 환자 중 내일 내원 예정 환자에게 메일 발송")
                tomorrow = datetime.date.today() + datetime.timedelta(days=1)
                tomorrow_str = tomorrow.strftime("%Y-%m-%d")

                st.info(f"**{tomorrow_str}**에 내원 예정인 환자를 찾습니다.")

                uploaded_file_for_email = st.file_uploader("최신 엑셀 파일을 업로드하세요.", type=["xlsx"], key="email_excel")
                
                if uploaded_file_for_email and email_password:
                    try:
                        excel_data, decrypted_file_bytes_io = load_excel(uploaded_file_for_email, password=None)
                        all_registered_patients = db.reference(f"patients/{st.session_state.current_firebase_key}").get()
                        
                        if not all_registered_patients:
                            st.warning("등록된 환자 정보가 없습니다. 먼저 환자를 등록해주세요.")
                        else:
                            all_registered_patients_df = pd.DataFrame.from_dict(all_registered_patients, orient='index')
                            
                            found_patients = []
                            for sheet_name in excel_data.sheet_names:
                                df = pd.read_excel(excel_data, sheet_name=sheet_name, header=1) # 첫 행은 제목이라 가정하고
                                df = df.fillna("").astype(str)
                                
                                # '예약일시' 컬럼이 있고, 날짜가 내일인 경우 필터링
                                if '예약일시' in df.columns:
                                    tomorrow_patients = df[df['예약일시'].str.contains(tomorrow_str, na=False)]
                                    
                                    for _, registered_patient in all_registered_patients_df.iterrows():
                                        matched_rows = tomorrow_patients[
                                            (tomorrow_patients['환자명'] == registered_patient['환자명']) &
                                            (tomorrow_patients['진료번호'] == registered_patient['진료번호'])
                                        ]
                                        if not matched_rows.empty:
                                            found_patients.append(matched_rows.iloc[0])
                            
                            if found_patients:
                                found_patients_df = pd.DataFrame(found_patients)
                                st.success(f"{tomorrow_str}에 내원 예정인 등록 환자 {len(found_patients_df)}명 발견!")
                                st.dataframe(found_patients_df[['환자명', '진료번호', '예약일시', '예약의사', '진료내역']])
                                
                                if st.button("메일 발송 시작"):
                                    with st.spinner("메일을 발송하고 있습니다..."):
                                        result = send_email(email_sender, found_patients_df, email_sender, email_password, date_str=tomorrow_str)
                                        if result is True:
                                            st.success("메일 발송이 성공적으로 완료되었습니다!")
                                        else:
                                            st.error(f"메일 발송 실패: {result}")
                            else:
                                st.info(f"내일 ({tomorrow_str}) 내원 예정인 등록 환자가 없습니다.")
                    except ValueError as e:
                        st.error(f"파일 처리 오류: {e}")
                    except Exception as e:
                        st.error(f"예상치 못한 오류: {e}")
            
            # 탭 3-2: 단체 메일
            with email_tab2:
                st.markdown("##### 단체 메일 발송")
                receiver_email = st.text_input("수신자 이메일 주소", value=email_sender)
                custom_subject = st.text_input("제목", value="단체 메일 알림")
                custom_message = st.text_area("메일 내용")
                
                if st.button("단체 메일 발송", key="send_bulk_email"):
                    if not custom_message:
                        st.warning("메일 내용을 입력해주세요.")
                    elif not is_valid_email(receiver_email):
                        st.warning("유효한 수신자 이메일 주소를 입력해주세요.")
                    elif not email_password:
                        st.warning("Gmail 앱 비밀번호를 입력해주세요.")
                    else:
                        with st.spinner("메일을 발송하고 있습니다..."):
                            result = send_email(receiver_email, pd.DataFrame(), email_sender, email_password, custom_message=custom_message)
                            if result is True:
                                st.success(f"'{receiver_email}'님에게 단체 메일이 성공적으로 발송되었습니다!")
                            else:
                                st.error(f"메일 발송 실패: {result}")

else:
    st.info("로그인하거나 사용자 이름을 입력하여 등록해주세요.")
