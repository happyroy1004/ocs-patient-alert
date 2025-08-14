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
import pickle # 인증 토큰 저장을 위해 사용

# --- 이메일 유효성 검사 함수 ---
def is_valid_email(email):
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
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
    return email.replace(".", "_dot_").replace("@", "_at_")

# 이메일 주소 복원 (Firebase 안전 키에서 원래 이메일로)
def recover_email(safe_id: str) -> str:
    email = safe_id.replace("_at_", "@").replace("_dot_", ".").replace("_com", ".com")
    return email

# 암호화된 엑셀 파일인지 확인
def is_encrypted_excel(file):
    try:
        file.seek(0)
        return msoffcrypto.OfficeFile(file).is_encrypted()
    except Exception:
        return False

# 엑셀 파일 로드 및 복호화
def load_excel(file, password=None):
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

# SCOPES = ["https://www.googleapis.com/auth/calendar"]
# NOTE: 이 코드는 secrets.toml에 client_secrets.json 내용이 있어야 동작합니다.
# NOTE: Streamlit 환경에서는 token.json 파일 저장이 어려우므로, 실제 배포 시에는 별도의 파일 시스템 또는 DB에 저장해야 합니다.

def get_google_calendar_service(user_email):
    """사용자별로 Google Calendar 서비스 객체를 반환합니다."""
    creds = None
    # NOTE: Streamlit 환경에서는 세션 상태를 활용하여 토큰을 저장하는 방식이 더 적합할 수 있습니다.
    # 예시: st.session_state.get(f'google_token_{user_email}')
    # 여기서는 pickle 파일을 사용하는 방식을 개념적으로만 보여줍니다.
    token_file = f'token_{sanitize_path(user_email)}.pickle'
    
    if os.path.exists(token_file):
        with open(token_file, 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # NOTE: 실제 환경에서는 Streamlit 앱 외부에서 이 인증 URL을 생성하고 사용자에게 보여줘야 합니다.
            # 예시: flow = InstalledAppFlow.from_client_secrets_file('client_secrets.json', SCOPES)
            # st.markdown(f"[Google 계정으로 로그인](https://accounts.google.com/o/oauth2/auth?...)")
            # 인증 후 redirect URL로 받은 코드를 처리하는 로직 필요
            # 이 코드는 로컬 개발 환경에서만 동작하는 예시입니다.
            st.info("Google Calendar 연동을 위해 인증이 필요합니다.")
            return None

    try:
        service = build('calendar', 'v3', credentials=creds)
        return service
    except HttpError as error:
        st.error(f'An error occurred: {error}')
        return None

def create_calendar_event(service, event_info):
    """Google Calendar에 이벤트를 생성합니다."""
    try:
        event = service.events().insert(calendarId='primary', body=event_info).execute()
        st.success(f"이벤트 생성 완료: {event.get('htmlLink')}")
        return event.get('htmlLink')
    except HttpError as error:
        st.error(f'Google Calendar 이벤트 생성 실패: {error}')
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
# URL 쿼리 매개변수에 'clear'가 있을 경우 초기화
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
user_name = st.text_input("사용자 이름을 입력하세요 (예시: 홍길동)")

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

    if len(matched_users_by_name) == 1:
        st.session_state.found_user_email = matched_users_by_name[0]["email"]
        st.session_state.user_id_input_value = matched_users_by_name[0]["email"]
        st.session_state.current_firebase_key = matched_users_by_name[0]["safe_key"]
        st.session_state.current_user_name = user_name
        st.info(f"**{user_name}**님으로 로그인되었습니다. 이메일 주소: **{st.session_state.found_user_email}**")
    elif len(matched_users_by_name) > 1:
        st.warning("동일한 이름의 사용자가 여러 명 있습니다. 정확한 이메일 주소를 입력해주세요.")
        st.session_state.found_user_email = ""
        st.session_state.user_id_input_value = ""
        st.session_state.current_firebase_key = ""
        st.session_state.current_user_name = ""
    else:
        st.info("새로운 사용자이거나 등록되지 않은 이름입니다. 이메일 주소를 입력해주세요.")
        st.session_state.found_user_email = ""
        st.session_state.user_id_input_value = ""
        st.session_state.current_firebase_key = ""
        st.session_state.current_user_name = ""

# 이메일 입력 필드
if not is_admin_input:
    if st.session_state.email_change_mode or not st.session_state.found_user_email:
        user_id_input = st.text_input("아이디를 입력하세요 (예시: example@gmail.com)", value=st.session_state.user_id_input_value)
        if user_id_input != st.session_state.user_id_input_value:
            st.session_state.user_id_input_value = user_id_input
    else:
        st.text_input("아이디 (등록된 이메일)", value=st.session_state.found_user_email, disabled=True)
        if st.button("이메일 주소 변경"):
            st.session_state.email_change_mode = True
            st.rerun()

# 이메일 변경 모드일 때 변경 완료 버튼 표시
if st.session_state.email_change_mode:
    if st.button("이메일 주소 변경 완료"):
        if is_valid_email(st.session_state.user_id_input_value):
            st.session_state.email_change_mode = False
            old_firebase_key = st.session_state.current_firebase_key
            new_email = st.session_state.user_id_input_value
            new_firebase_key = sanitize_path(new_email)

            if old_firebase_key and old_firebase_key != new_firebase_key:
                users_ref.child(new_firebase_key).update({"name": st.session_state.current_user_name, "email": new_email})
                old_patient_data = db.reference(f"patients/{old_firebase_key}").get()
                if old_patient_data:
                    db.reference(f"patients/{new_firebase_key}").set(old_patient_data)
                    db.reference(f"patients/{old_firebase_key}").delete()
                users_ref.child(old_firebase_key).delete()
                st.session_state.current_firebase_key = new_firebase_key
                st.session_state.found_user_email = new_email
                st.success(f"이메일 주소가 **{new_email}**로 성공적으로 변경되었습니다.")
            elif not old_firebase_key:
                st.session_state.current_firebase_key = new_firebase_key
                st.session_state.found_user_email = new_email
                st.success(f"새로운 사용자 정보가 등록되었습니다: {st.session_state.current_user_name} ({new_email})")
            else:
                st.success("이메일 주소 변경사항이 없습니다.")
            st.rerun()
        else:
            st.error("올바른 이메일 주소 형식이 아닙니다.")

# --- Admin 모드 로그인 처리 ---
if is_admin_input:
    st.session_state.logged_in_as_admin = True
    st.session_state.found_user_email = "admin"
    st.session_state.current_user_name = "admin"
    
    # 엑셀 업로드 섹션 - 비밀번호 없이도 접근 가능
    st.subheader("💻 Excel File Processor")
    uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])
    
    # 엑셀 업로드 로직
    if uploaded_file:
        uploaded_file.seek(0)
        
        password = st.text_input("엑셀 파일 비밀번호 입력", type="password") if is_encrypted_excel(uploaded_file) else None
        if is_encrypted_excel(uploaded_file) and not password:
            st.info("암호화된 파일입니다. 비밀번호를 입력해주세요.")
            st.stop()
        
        try:
            file_name = uploaded_file.name
            date_match = re.search(r'(\d{4})', file_name)
            extracted_date = date_match.group(1) if date_match else None

            xl_object, raw_file_io = load_excel(uploaded_file, password)
            excel_data_dfs, styled_excel_bytes = process_excel_file_and_style(raw_file_io)

            if excel_data_dfs is None or styled_excel_bytes is None:
                st.warning("엑셀 파일 처리 중 문제가 발생했거나 처리할 데이터가 없습니다.")
                st.stop()
            
            sender = st.secrets["gmail"]["sender"]
            sender_pw = st.secrets["gmail"]["app_password"]

            all_users_meta = users_ref.get()
            all_patients_data = db.reference("patients").get()

            if not all_users_meta and not all_patients_data:
                st.warning("Firebase에 등록된 사용자 또는 환자 데이터가 없습니다. 이메일 전송은 불가능합니다.")
            elif not all_users_meta:
                st.warning("Firebase users 노드에 등록된 사용자 메타 정보가 없습니다. 이메일 전송 시 이름 대신 이메일이 사용됩니다.")
            elif not all_patients_data:
                st.warning("Firebase patients 노드에 등록된 환자 데이터가 없습니다. 매칭할 수 없습니다.")

            matched_users = []
            
            if all_patients_data:
                for uid_safe, registered_patients_for_this_user in all_patients_data.items():
                    user_email = recover_email(uid_safe)
                    user_display_name = user_email
                    
                    if all_users_meta and uid_safe in all_users_meta:
                        user_meta = all_users_meta[uid_safe]
                        if "name" in user_meta:
                            user_display_name = user_meta["name"]
                        if "email" in user_meta:
                            user_email = user_meta["email"]
                    
                    registered_patients_data = []
                    if registered_patients_for_this_user:
                        for key, val in registered_patients_for_this_user.items():
                            registered_patients_data.append({
                                "환자명": val["환자명"].strip(),
                                "진료번호": val["진료번호"].strip().zfill(8),
                                "등록과": val.get("등록과", "")
                            })
                    
                    matched_rows_for_user = []

                    for sheet_name_excel_raw, df_sheet in excel_data_dfs.items():
                        excel_sheet_name_lower = sheet_name_excel_raw.strip().lower()

                        excel_sheet_department = None
                        for keyword, department_name in sorted(sheet_keyword_to_department_map.items(), key=lambda item: len(item[0]), reverse=True):
                            if keyword.lower() in excel_sheet_name_lower:
                                excel_sheet_department = department_name
                                break
                        
                        if not excel_sheet_department:
                            continue
                            
                        for _, excel_row in df_sheet.iterrows():
                            excel_patient_name = excel_row["환자명"].strip()
                            excel_patient_pid = excel_row["진료번호"].strip().zfill(8)
                            
                            for registered_patient in registered_patients_data:
                                if (registered_patient["환자명"] == excel_patient_name and
                                        registered_patient["진료번호"] == excel_patient_pid and
                                        registered_patient["등록과"] == excel_sheet_department):
                                    
                                    matched_row_copy = excel_row.copy()
                                    matched_row_copy["시트"] = sheet_name_excel_raw
                                    matched_rows_for_user.append(matched_row_copy)
                                    break
                                    
                    if matched_rows_for_user:
                        combined_matched_df = pd.DataFrame(matched_rows_for_user)
                        matched_users.append({"email": user_email, "name": user_display_name, "data": combined_matched_df})

            if matched_users:
                st.success(f"{len(matched_users)}명의 사용자와 일치하는 환자 발견됨.")
                
                for user_match_info in matched_users:
                    st.markdown(f"**수신자:** {user_match_info['name']} ({user_match_info['email']})")
                    st.dataframe(user_match_info['data'])
                
                mail_col, calendar_col = st.columns(2)
                
                with mail_col:
                    if st.button("매칭된 환자에게 메일 보내기"):
                        for user_match_info in matched_users:
                            real_email = user_match_info['email']
                            df_matched = user_match_info['data']
                            result = send_email(real_email, df_matched, sender, sender_pw, date_str=extracted_date)
                            if result is True:
                                st.success(f"**{user_match_info['name']}** ({real_email}) 전송 완료")
                            else:
                                st.error(f"**{user_match_info['name']}** ({real_email}) 전송 실패: {result}")
                
                with calendar_col:
                    # 관리자용 구글 캘린더 일정 추가 버튼
                    if st.button("Google Calendar 일정 추가"):
                        # NOTE: 관리자 계정의 Google Calendar 서비스 객체를 가져오는 로직이 필요합니다.
                        # 여기서는 임시로 None을 사용합니다.
                        # admin_service = get_google_calendar_service("admin@example.com")
                        admin_service = None # 실제 구현 시 위 함수를 사용하여 인증
                        
                        if admin_service:
                            for user_match_info in matched_users:
                                df_matched = user_match_info['data']
                                if not df_matched.empty:
                                    for index, row in df_matched.iterrows():
                                        try:
                                            # 날짜와 시간을 조합하여 이벤트 시작/종료 시간 설정
                                            # NOTE: '예약시간' 컬럼의 형식이 정확히 'HH:mm'이어야 함
                                            start_time_str = f"{extracted_date}T{row['예약시간']}:00"
                                            end_time_str = start_time_str # 편의상 시작 시간과 동일하게 설정
                                            
                                            event_summary = f"[내원 예정] 환자: {row['환자명']} ({row['진료번호']})"
                                            event_description = f"예약의사: {row['예약의사']}, 진료내역: {row['진료내역']}"
                                            
                                            event = {
                                                'summary': event_summary,
                                                'description': event_description,
                                                'start': {
                                                    'dateTime': start_time_str,
                                                    'timeZone': 'Asia/Seoul', # 시간대 설정
                                                },
                                                'end': {
                                                    'dateTime': end_time_str,
                                                    'timeZone': 'Asia/Seoul',
                                                },
                                            }
                                            create_calendar_event(admin_service, event)
                                            st.success(f"{row['환자명']} 환자의 일정을 캘린더에 추가했습니다.")
                                        except Exception as e:
                                            st.error(f"{row['환자명']} 환자의 일정 추가 실패: {e}")
                        else:
                            st.error("Google Calendar 서비스 인증이 필요합니다. 관리자 계정으로 다시 시도해주세요.")
                            
            else:
                st.info("엑셀 파일 처리 완료. 매칭된 환자가 없습니다.")
                
            output_filename = uploaded_file.name.replace(".xlsx", "_processed.xlsx").replace(".xlsm", "_processed.xlsm")
            st.download_button(
                "처리된 엑셀 다운로드",
                data=styled_excel_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except ValueError as ve:
            st.error(f"파일 처리 실패: {ve}")
        except Exception as e:
            st.error(f"예상치 못한 오류 발생: {e}")

    # 관리자 비밀번호 입력 섹션 - 별도 분리
    st.markdown("---")
    st.subheader("🛠️ Administer password")
    admin_password_input = st.text_input("관리자 비밀번호를 입력하세요", type="password", key="admin_password")

    # secrets.toml에서 비밀번호 불러오기
    try:
        secret_admin_password = st.secrets["admin"]["password"]
    except KeyError:
        secret_admin_password = None
        st.error("⚠️ secrets.toml 파일에 'admin.password' 설정이 없습니다. 개발자에게 문의하세요.")
    
    if admin_password_input and admin_password_input == secret_admin_password:
        st.session_state.admin_password_correct = True
        st.success("관리자 권한이 활성화되었습니다.")
    elif admin_password_input and admin_password_input != secret_admin_password:
        st.error("비밀번호가 틀렸습니다.")
        st.session_state.admin_password_correct = False
    
    # 비밀번호가 맞았을 때만 추가 기능 표시
    if st.session_state.admin_password_correct:
        st.markdown("---")
        st.subheader("📦 메일 발송") # 제목 변경
        
        all_users_meta = users_ref.get()
        user_list_for_dropdown = [f"{user_info.get('name', '이름 없음')} ({user_info.get('email', '이메일 없음')})" 
                                    for user_info in (all_users_meta.values() if all_users_meta else [])]
        
        # '모든 사용자 선택' 체크박스 추가
        select_all_users_button = st.button("모든 사용자 선택/해제", key="select_all_btn")
        if select_all_users_button:
            st.session_state.select_all_users = not st.session_state.select_all_users

        default_selection = user_list_for_dropdown if st.session_state.select_all_users else []

        selected_users_for_mail = st.multiselect("보낼 사용자 선택", user_list_for_dropdown, default=default_selection, key="mail_multiselect")
        
        custom_message = st.text_area("보낼 메일 내용", height=200)
        if st.button("메일 보내기"): # 버튼 이름 변경
            if custom_message:
                sender = st.secrets["gmail"]["sender"]
                sender_pw = st.secrets["gmail"]["app_password"]
                
                email_list = []
                if selected_users_for_mail:
                    for user_str in selected_users_for_mail:
                        match = re.search(r'\((.*?)\)', user_str)
                        if match:
                            email_list.append(match.group(1))
                
                if email_list:
                    with st.spinner("메일 전송 중..."):
                        for email in email_list:
                            result = send_email(email, pd.DataFrame(), sender, sender_pw, custom_message=custom_message)
                            if result is True:
                                st.success(f"{email}로 메일 전송 완료!")
                            else:
                                st.error(f"{email}로 메일 전송 실패: {result}")
                else:
                    st.warning("메일 내용을 입력했으나, 선택된 사용자가 없습니다. 전송이 진행되지 않았습니다.")
            else:
                st.warning("메일 내용을 입력해주세요.")
        
        st.markdown("---")
        st.subheader("🗑️ 사용자 삭제")
        users_to_delete = st.multiselect("삭제할 사용자 선택", user_list_for_dropdown, key="delete_user_multiselect")
        if st.button("선택한 사용자 삭제"):
            if users_to_delete:
                for user_to_del_str in users_to_delete:
                    match = re.search(r'\((.*?)\)', user_to_del_str)
                    if match:
                        email_to_del = match.group(1)
                        safe_key_to_del = sanitize_path(email_to_del)
                        
                        db.reference(f"users/{safe_key_to_del}").delete()
                        db.reference(f"patients/{safe_key_to_del}").delete()
                        st.success(f"사용자 {user_to_del_str} 삭제 완료.")
                st.rerun()
            else:
                st.warning("삭제할 사용자를 선택해주세요.")
    


# --- 일반 사용자 모드 ---
else: # is_admin_input이 False일 때
    # 최종적으로 사용할 Firebase 키
    user_id_final = st.session_state.user_id_input_value if st.session_state.email_change_mode or not st.session_state.found_user_email else st.session_state.found_user_email
    firebase_key = sanitize_path(user_id_final) if user_id_final else ""

    if not user_name or not user_id_final:
        st.info("내원 알람 노티를 받을 이메일 주소와 사용자 이름을 입력해주세요.")
        st.stop()

    patients_ref_for_user = db.reference(f"patients/{firebase_key}")

    # 사용자 정보 (이름, 이메일) Firebase 'users' 노드에 저장 또는 업데이트
    if not st.session_state.email_change_mode:
        current_user_meta_data = users_ref.child(firebase_key).get()
        if not current_user_meta_data or current_user_meta_data.get("name") != user_name or current_user_meta_data.get("email") != user_id_final:
            users_ref.child(firebase_key).update({"name": user_name, "email": user_id_final})
            st.success(f"사용자 정보가 업데이트되었습니다: {user_name} ({user_id_final})")
            # 세션 상태 업데이트 (새로운 등록 또는 정보 변경 시)
            st.session_state.current_firebase_key = firebase_key
            st.session_state.current_user_name = user_name
            st.session_state.found_user_email = user_id_final

    st.subheader(f"{user_name}님의 등록 환자 목록")
    
    # 일반 사용자용 구글 캘린더 권한 부여 버튼 추가
    if st.button("Google Calendar 권한 부여"):
        # NOTE: 이 버튼 클릭 시 Google OAuth 2.0 인증 절차를 시작해야 합니다.
        # 실제 구현에서는 redirect URL을 통해 인증 코드를 받아와 토큰을 생성해야 합니다.
        st.warning("Google Calendar 연동을 위한 인증 절차를 시작합니다. (실제 환경에서는 별도 인증 창이 열립니다.)")
        # 예시:
        # flow = InstalledAppFlow.from_client_secrets_file('client_secrets.json', SCOPES)
        # auth_url, _ = flow.authorization_url(prompt='consent')
        # st.markdown(f"[이 링크를 클릭하여 Google Calendar에 접근 권한을 부여하세요]({auth_url})")

    existing_patient_data = patients_ref_for_user.get()

    if existing_patient_data:
        desired_order = ['소치', '외과', '보철', '내과', '교정']
        order_map = {dept: i for i, dept in enumerate(desired_order)}
        patient_list = list(existing_patient_data.items())
        sorted_patient_list = sorted(patient_list, key=lambda item: order_map.get(item[1].get('등록과', '미지정'), float('inf')))

        cols_count = 3
        cols = st.columns(cols_count)
        
        for idx, (key, val) in enumerate(sorted_patient_list):
            with cols[idx % cols_count]:
                with st.container(border=True):
                    info_col, btn_col = st.columns([4, 1])
                    
                    with info_col:
                        st.markdown(f"**{val['환자명']}** / {val['진료번호']} / {val.get('등록과', '미지정')}")
                    
                    with btn_col:
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
                st.success(f"{name} ({pid}) [{selected_department}] 환자 등록 완료")
                st.rerun()
