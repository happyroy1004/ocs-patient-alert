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
import datetime
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

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
        st.info("secrets.toml 파일의 Firebase 설정을 확인해주세요.")
        st.stop()

# Firebase-safe 경로 변환 (이메일을 Firebase 키로 사용하기 위해)
def sanitize_path(email):
    return email.replace(".", "_dot_").replace("@", "_at_")

# 이메일 주소 복원 (Firebase 안전 키에서 원래 이메일로)
def recover_email(safe_id: str) -> str:
    email = safe_id.replace("_at_", "@").replace("_dot_", ".")
    if not email.endswith(".com") and email.endswith("_com"):
        email = email[:-4] + ".com"
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
def send_email(receiver, rows, sender, password, date_str=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver

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

# --- 엑셀 처리 관련 상수 및 함수 ---
sheet_keyword_to_department_map = {
    '치과보철과': '보철', '보철과': '보철', '보철': '보철',
    '치과교정과' : '교정', '교정과': '교정', '교정': '교정',
    '구강 악안면외과' : '외과', '구강악안면외과': '외과', '외과': '외과',
    '구강 내과' : '내과', '구강내과': '내과', '내과': '내과',
    '치과보존과' : '보존', '보존과': '보존', '보존': '보존',
    '소아치과': '소치', '소치': '소치', '소아 치과': '소치',
    '원내생진료센터': '원내생', '원내생': '원내생','원내생 진료센터': '원내생',
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
            st.warning(f"시트 '{sheet_name_raw}'에 유효한 데이터가 충분하지 않습니다. 건너킵니다.")
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
                    text = str(cell.value)
                    if any(keyword in text for keyword in ['본딩', 'bonding']):
                        cell.font = Font(bold=True)

    final_output_bytes = io.BytesIO()
    wb_styled.save(final_output_bytes)
    final_output_bytes.seek(0)

    return processed_sheets_dfs, final_output_bytes

# --- 구글 캘린더 관련 전역 변수 설정 ---
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

def get_google_calendar_service():
    """
    구글 캘린더 API 서비스 객체를 반환.
    Streamlit Cloud 배포 환경에 최적화된 인증 로직을 사용합니다.
    """
    creds = None
    
    # Streamlit secrets에서 클라이언트 정보 가져오기
    client_config = {
        "web": {
            "client_id": st.secrets["google_calendar"]["client_id"],
            "client_secret": st.secrets["google_calendar"]["client_secret"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs"
        }
    }
    
    # Streamlit 세션 상태에 토큰을 저장하여 재사용
    token_info = st.session_state.get('google_calendar_token', None)
    
    if token_info:
        creds = Credentials.from_authorized_user_info(token_info, SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_config(
                client_config, 
                SCOPES, 
                redirect_uri=st.secrets["google_calendar"]["redirect_uri"]
            )
            auth_url, _ = flow.authorization_url(prompt='consent')
            
            st.warning("⚠️ Google 캘린더 연동이 필요합니다.")
            st.markdown(f"**[Google 계정 연동하기]({auth_url})**")
            st.info("위 링크를 클릭하여 Google 계정에 로그인하고 권한을 허용하세요. 이후, 페이지 URL의 'code=' 뒤에 있는 코드를 복사하여 아래에 붙여넣어주세요.")
            
            auth_code = st.text_input("인증 코드 붙여넣기", type="password")
            
            if auth_code:
                try:
                    flow.fetch_token(code=auth_code)
                    creds = flow.credentials
                    st.session_state['google_calendar_token'] = json.loads(creds.to_json())
                    st.success("Google 계정 연동에 성공했습니다! 페이지를 새로고침하거나 다시 시도해주세요.")
                    st.rerun()
                except Exception as e:
                    st.error(f"토큰을 가져오는 데 실패했습니다: {e}")
                    st.stop()
            else:
                st.stop()
    
    return build('calendar', 'v3', credentials=creds)

def create_calendar_event(service, receiver_email, rows, date_str):
    """
    구글 캘린더에 이벤트를 생성하는 함수
    :param service: 구글 캘린더 API 서비스 객체
    :param receiver_email: 이벤트 참석자로 추가할 이메일 주소
    :param rows: DataFrame 형태의 환자 데이터
    :param date_str: 예약 날짜 (예: "2025-08-15")
    :return: 생성된 이벤트 ID 또는 에러 메시지
    """
    event_list = []
    
    for _, row in rows.iterrows():
        summary = f"{row['환자명']} ({row['진료번호']})"
        description = f"진료내역: {row['진료내역']}\n예약의사: {row['예약의사']}\n시트: {row['시트']}"
        
        try:
            start_time = datetime.datetime.strptime(f"{date_str} {row['예약시간']}", "%Y%m%d %H:%M")
            end_time = start_time + datetime.timedelta(hours=1)
        except ValueError:
            start_time = datetime.datetime.now()
            end_time = start_time + datetime.timedelta(hours=1)
            st.warning(f"시간 형식 오류로 현재 시간으로 캘린더 이벤트를 생성했습니다. (환자: {row['환자명']})")

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
            'attendees': [
                {'email': receiver_email},
            ],
        }
        event_list.append(event)
    
    created_events = []
    for event in event_list:
        try:
            event = service.events().insert(calendarId='primary', body=event).execute()
            created_events.append(event.get('htmlLink'))
        except HttpError as error:
            st.error(f"캘린더 이벤트 생성 실패: {error}")
            return str(error)
    
    return created_events

# --- Streamlit 애플리케이션 시작 ---
st.title("환자 내원 확인 시스템")
st.markdown("---")
st.markdown("<p style='text-align: left; color: grey; font-size: small;'>directed by HSY</p>", unsafe_allow_html=True)

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

# 사용자 입력 필드
user_name = st.text_input("사용자 이름을 입력하세요 (예시: 홍길동)")

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

users_ref = db.reference("users")

is_admin_mode = (user_name.strip().lower() == "admin")

if user_name and not is_admin_mode and not st.session_state.email_change_mode:
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

if st.session_state.email_change_mode or not st.session_state.found_user_email or is_admin_mode:
    user_id_input = st.text_input("아이디를 입력하세요 (예시: example@gmail.com)", value=st.session_state.user_id_input_value)
    if user_id_input != st.session_state.user_id_input_value:
        st.session_state.user_id_input_value = user_id_input
else:
    st.text_input("아이디 (등록된 이메일)", value=st.session_state.found_user_email, disabled=True)
    if st.button("이메일 주소 변경"):
        st.session_state.email_change_mode = True
        st.rerun()

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
            
user_id_final = st.session_state.user_id_input_value if st.session_state.email_change_mode or not st.session_state.found_user_email else st.session_state.found_user_email

if not user_name or (not user_id_final and not is_admin_mode):
    st.info("내원 알람 노티를 받을 이메일 주소와 사용자 이름을 입력해주세요.")
    st.stop()

firebase_key = sanitize_path(user_id_final) if user_id_final else ""

if not is_admin_mode:
    patients_ref_for_user = db.reference(f"patients/{firebase_key}")

    if not st.session_state.email_change_mode:
        current_user_meta_data = users_ref.child(firebase_key).get()
        if not current_user_meta_data or current_user_meta_data.get("name") != user_name or current_user_meta_data.get("email") != user_id_final:
            users_ref.child(firebase_key).update({"name": user_name, "email": user_id_final})
            st.success(f"사용자 정보가 업데이트되었습니다: {user_name} ({user_id_final})")
            st.session_state.current_firebase_key = firebase_key
            st.session_state.current_user_name = user_name
            st.session_state.found_user_email = user_id_final

if not is_admin_mode:
    st.subheader(f"{user_name}님의 등록 환자 목록")
    patients_ref_for_user = db.reference(f"patients/{firebase_key}")
    existing_patient_data = patients_ref_for_user.get()

    st.markdown("""
    <style>
    .patient-list-container {
        display: flex;
        flex-wrap: wrap;
        gap: 1rem;
        justify-content: flex-start;
    }
    .patient-item {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 0.5rem;
        background-color: #f0f2f6;
        border-radius: 0.5rem;
        flex-grow: 1;
        min-width: 250px;
        margin-bottom: 0.5rem;
        word-break: break-all;
    }
    .patient-info {
        flex-grow: 1;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        padding-right: 10px;
    }
    .small-delete-button {
        background-color: #e6e6e6;
        color: #000000;
        border: none;
        padding: 0.2rem 0.5rem;
        border-radius: 0.3rem;
        cursor: pointer;
        font-size: 0.75rem;
        width: auto;
        flex-shrink: 0;
    }
    .small-delete-button:hover {
        background-color: #cccccc;
    }

    @media (min-width: 260px) {
        .patient-list-container {
            justify-content: space-between;
        }
        .patient-item {
            width: 32%;
        }
    }
    </style>
    """, unsafe_allow_html=True)

    if existing_patient_data:
        st.markdown('<div class="patient-list-container">', unsafe_allow_html=True)
        for key, val in existing_patient_data.items():
            st.markdown('<div class="patient-item">', unsafe_allow_html=True)
            info_col, btn_col = st.columns([0.8, 0.2])
            with info_col:
                st.markdown(f'<div class="patient-info">{val["환자명"]} / {val["진료번호"]} / {val.get("등록과", "미지정")}</div>', unsafe_allow_html=True)
            with btn_col:
                st.markdown(
                    f"""
                    <form action="" method="post" style="display:inline-block; margin:0; padding:0;">
                        <input type="hidden" name="delete_key" value="{key}">
                        <button type="submit" class="small-delete-button">삭제</button>
                    </form>
                    """,
                    unsafe_allow_html=True
                )
            st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        if "delete_key" in st.query_params:
            key_to_delete = st.query_params["delete_key"]
            patients_ref_for_user.child(key_to_delete).delete()
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

else:
    st.subheader("💻 관리자 모드 💻")
    uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])

    if uploaded_file:
        uploaded_file.seek(0)

        password = None
        if is_encrypted_excel(uploaded_file):
            password = st.text_input("엑셀 파일 비밀번호 입력", type="password")
            if not password:
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

                if st.button("매칭된 환자에게 메일 보내기"):
                    for user_match_info in matched_users:
                        real_email = user_match_info['email']
                        df_matched = user_match_info['data']
                        
                        result_email = send_email(real_email, df_matched, sender, sender_pw, date_str=extracted_date)
                        if result_email is True:
                            st.success(f"**{user_match_info['name']}** ({real_email}) 이메일 전송 완료")
                        else:
                            st.error(f"**{user_match_info['name']}** ({real_email}) 이메일 전송 실패: {result_email}")

                        try:
                            calendar_service = get_google_calendar_service()
                            event_links = create_calendar_event(calendar_service, real_email, df_matched, date_str=extracted_date)
                            if isinstance(event_links, list):
                                st.success(f"**{user_match_info['name']}** ({real_email}) 캘린더 이벤트 생성 완료")
                            else:
                                st.error(f"**{user_match_info['name']}** ({real_email}) 캘린더 이벤트 생성 실패: {event_links}")
                        except Exception as e:
                            st.error(f"**{user_match_info['name']}** ({real_email}) 캘린더 연동 오류: {e}")

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
