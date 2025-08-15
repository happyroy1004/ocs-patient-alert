# #1. Imports, Validation Functions, and Firebase Initialization
import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os
import re
import smtplib
from email.mime.text import MIMEText
import pandas as pd
from datetime import datetime, timedelta
import pytz
import base64
import io
from openpyxl import load_workbook
from openpyxl.styles import Font
import msoffcrypto

# --- Firebase 초기화 ---
if not firebase_admin._apps:
    cred = credentials.Certificate({
        "type": "service_account",
        "project_id": st.secrets["firebase"]["project_id"],
        "private_key_id": st.secrets["firebase"]["private_key_id"],
        "private_key": st.secrets["firebase"]["private_key"].replace('\\n', '\n'),
        "client_email": st.secrets["firebase"]["client_email"],
        "client_id": st.secrets["firebase"]["client_id"],
        "auth_uri": st.secrets["firebase"]["auth_uri"],
        "token_uri": st.secrets["firebase"]["token_uri"],
        "auth_provider_x509_cert_url": st.secrets["firebase"]["auth_provider_x509_cert_url"],
        "client_x509_cert_url": st.secrets["firebase"]["client_x509_cert_url"]
    })
    firebase_admin.initialize_app(cred, {
        'databaseURL': st.secrets["firebase"]["database_url"]
    })
db = db

# --- 이메일 유효성 검사 함수 ---
def is_valid_email(email):
    # 정규표현식에서 \\. -> \. 로 수정
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(email_regex, email) is not None

# --- 경로 정리 함수 ---
def sanitize_path(s):
    return s.replace(".", "_dot_").replace("@", "_at_")

# --- 이메일 복원 함수 ---
def recover_email(s):
    return s.replace("_dot_", ".").replace("_at_", "@")

# --- 구글 캘린더 관련 함수 ---
SCOPES = ['https://www.googleapis.com/auth/calendar']

def load_google_creds_from_firebase(user_id):
    creds_ref = db.reference(f"tokens/{user_id}")
    token_info = creds_ref.get()
    
    if not token_info:
        return None

    try:
        creds = Credentials(
            token_info.get("token"),
            refresh_token=token_info.get("refresh_token"),
            id_token=token_info.get("id_token"),
            token_uri=token_info.get("token_uri"),
            client_id=st.secrets["google_calendar"]["client_id"],
            client_secret=st.secrets["google_calendar"]["client_secret"],
            scopes=SCOPES
        )
        return creds
    except Exception as e:
        st.error(f"Failed to load credentials: {e}")
        return None

def save_google_creds_to_firebase(user_id, creds):
    creds_ref = db.reference(f"tokens/{user_id}")
    creds_ref.set({
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "id_token": creds.id_token,
        "token_uri": creds.token_uri
    })

def create_calendar_event(service, patient_name, pid, department, start_date, start_time, doctor_name, summary="내원 환자"):
    try:
        event = {
            'summary': f'{summary} ({patient_name}, {department})',
            'location': '서울대학교 치과병원',
            'description': f'환자명: {patient_name}\n진료번호: {pid}\n진료과: {department}\n예약의사: {doctor_name}',
            'start': {
                'dateTime': f'{start_date}T{start_time}:00',
                'timeZone': 'Asia/Seoul',
            },
            'end': {
                'dateTime': (datetime.strptime(f'{start_date}T{start_time}', '%Y-%m-%dT%H:%M') + timedelta(minutes=30)).isoformat(),
                'timeZone': 'Asia/Seoul',
            },
        }
        event = service.events().insert(calendarId='primary', body=event).execute()
        st.success(f"Google Calendar에 '{patient_name}' 환자 일정이 추가되었습니다.")
    except Exception as e:
        st.error(f"Google Calendar에 일정을 추가하지 못했습니다: {e}")

# --- 메일 전송 함수 ---
def send_email(to_email, df_matched, sender, sender_pw, custom_message=None, date_str=None):
    from_email = sender
    msg = MIMEText(f"{custom_message}\n\n환자 정보:\n{df_matched.to_string(index=False)}")
    msg['Subject'] = f'{date_str} 내원 환자 알림' if date_str else '내원 환자 알림'
    msg['From'] = from_email
    msg['To'] = to_email

    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender, sender_pw)
        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"메일 전송 실패: {e}")
        return False

# --- 세션 상태 초기화 ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "current_user_name" not in st.session_state:
    st.session_state.current_user_name = None
if "found_user_email" not in st.session_state:
    st.session_state.found_user_email = None
if "admin_password_correct" not in st.session_state:
    st.session_state.admin_password_correct = False
if "select_all_users" not in st.session_state:
    st.session_state.select_all_users = False
if "processed_excel_data_dfs" not in st.session_state:
    st.session_state.processed_excel_data_dfs = None
if "processed_styled_bytes" not in st.session_state:
    st.session_state.processed_styled_bytes = None
if 'google_calendar_service' not in st.session_state:
    st.session_state.google_calendar_service = None
if "email_change_mode" not in st.session_state:
    st.session_state.email_change_mode = False
if "user_id_input_value" not in st.session_state:
    st.session_state.user_id_input_value = ""

#2. Excel and Email Processing Functions
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
                    width: 100%; max-width: 100%;
                    border-collapse: collapse;
                    font-family: Arial, sans-serif;
                    font-size: 14px;
                    table-layout: fixed;
                }
                th, td {
                    border: 1px solid #dddddd; text-align: left;
                    padding: 8px;
                    vertical-align: top;
                    word-wrap: break-word;
                    word-break: break-word;
                }
                th {
                    background-color: #f2f2f2; font-weight: bold;
                    white-space: nowrap;
                }
                tr:nth-child(even) {
                    background-color: #f9f9f9;
                }
                .table-container {
                    overflow-x: auto; -webkit-overflow-scrolling: touch;
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


#3. Google Calendar API Functions
# --- Google Calendar API 관련 함수 (수정) ---

# 사용할 스코프 정의. 캘린더 이벤트 생성 권한
SCOPES = ["https://www.googleapis.com/auth/calendar.events"]

# 수정 코드 (Revised Code)
def get_google_calendar_service(user_id_safe):
    """
    사용자별로 Google Calendar 서비스 객체를 반환하거나 인증 URL을 표시합니다. Streamlit 세션 상태와 Firebase를 활용하여 인증 정보를 관리합니다.
    """
    creds = st.session_state.get(f"google_creds_{user_id_safe}")
    
    if not creds:
        creds = load_google_creds_from_firebase(user_id_safe)
        if creds:
            st.session_state[f"google_creds_{user_id_safe}"] = creds

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
            # Store credentials in Firebase
            save_google_creds_to_firebase(user_id_safe, creds)
            st.success("Google Calendar 인증이 완료되었습니다.")
            st.query_params.clear()
            st.rerun()
        else:
            auth_url, _ = flow.authorization_url(prompt='consent')
            st.warning("Google Calendar 연동을 위해 인증이 필요합니다. 아래 링크를 클릭하여 권한을 부여하세요.")
            st.markdown(f"**[Google Calendar 인증 링크]({auth_url})**")
            return None

    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        st.session_state[f"google_creds_{user_id_safe}"] = creds
        # Update credentials in Firebase
        save_google_creds_to_firebase(user_id_safe, creds)

    try:
        service = build('calendar', 'v3', credentials=creds)
        return service
    except HttpError as error:
        st.error(f'Google Calendar 서비스 생성 실패: {error}')
        st.session_state.pop(f"google_creds_{user_id_safe}", None)
        # Clear invalid credentials from Firebase
        db.reference(f"users/{user_id_safe}/google_creds").delete()
        return None

def create_calendar_event(service, patient_name, pid, department, reservation_date_str, reservation_time_str, doctor_name, treatment_details):
    """
    Google Calendar에 이벤트를 생성합니다. 예약 날짜와 시간을 기반으로 30분 일정을 만들고 의사 이름과 진료내역을 추가합니다.
    """
    seoul_tz = datetime.timezone(datetime.timedelta(hours=9))

    # 예약 날짜와 시간을 사용하여 이벤트 시작/종료 시간 설정
    try:
        date_time_str = f"{reservation_date_str} {reservation_time_str}"
        
        # Naive datetime 객체 생성 후 한국 시간대(KST)로 로컬라이즈
        naive_start = datetime.datetime.strptime(date_time_str, "%Y-%m-%d %H:%M")
        event_start = naive_start.replace(tzinfo=seoul_tz)
        event_end = event_start + datetime.timedelta(minutes=30)
        
    except ValueError as e:
        # 날짜 형식 파싱 실패 시 현재 시간 사용 (예외 처리)
        st.warning(f"'{patient_name}' 환자의 날짜/시간 형식 파싱 실패: {e}. 현재 시간으로 일정을 추가합니다.")
        event_start = datetime.datetime.now(seoul_tz)
        event_end = event_start + datetime.timedelta(minutes=30)
    
    # 캘린더 이벤트 요약(summary)을 새로운 형식으로 변경
    summary_text = f'내원예정: {patient_name} ({department}, {doctor_name})' if doctor_name else f'내원예정: {patient_name} ({department})'

    event = {
        'summary': summary_text,
        'location': f'진료번호: {pid}',
        'description': f'환자명: {patient_name}\n진료번호: {pid}\n등록 과: {department}\n진료내역: {treatment_details}',
        'start': {
            'dateTime': event_start.isoformat(),
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': event_end.isoformat(),
            'timeZone': 'Asia/Seoul',
        },
    }
    
    try:
        event = service.events().insert(calendarId='primary', body=event).execute()
        st.success(f"'{patient_name}' 환자 내원 일정이 캘린더에 추가되었습니다.")
    except HttpError as error:
        st.error(f"캘린더 이벤트 생성 중 오류 발생: {error}")
        st.warning("구글 캘린더 인증 권한을 다시 확인해주세요.")
    except Exception as e:
        st.error(f"알 수 없는 오류 발생: {e}")

# #4. Excel Processing Constants and Functions
# --- 엑셀 처리 관련 상수 및 함수 ---
# 필요한 라이브러리 추가
import pandas as pd
import openpyxl
from openpyxl.styles import Font
from openpyxl import load_workbook
import msoffcrypto
import re
import datetime
import io
import streamlit as st
import os

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

def process_sheet_v8(df, professors_list, sheet_key):
    """
    각 시트의 데이터를 처리하는 함수 (버전 8).
    - '교수님'으로 표기된 행의 '예약의사'를 "<교수님>"으로 변경
    - '교수님'으로 표기되지 않은 행의 '예약의사'를 그대로 유지
    - 최종 컬럼: '환자명', '진료번호', '예약의사', '진료내역', '예약일시', '예약시간', '진료과', '담당의사'
    """
    # 예약의사 컬럼이 없는 경우 처리
    if '예약의사' not in df.columns:
        st.warning("경고: '예약의사' 컬럼이 존재하지 않습니다.")
        return pd.DataFrame()

    df['진료과'] = sheet_key

    # 예약의사가 교수님 리스트에 포함되는 경우, 예약의사를 <교수님>으로 변경
    df['예약의사'] = df.apply(
        lambda row: '<교수님>' if row['예약의사'] in professors_list else row['예약의사'],
        axis=1
    )

    # 필요한 컬럼만 선택
    required_cols = ['환자명', '진료번호', '예약의사', '진료내역', '예약일시', '예약시간', '진료과']
    
    # 누락된 컬럼 처리
    for col in required_cols:
        if col not in df.columns:
            df[col] = ''
    
    # 최종 데이터프레임 구성
    df_final = df[required_cols]

    return df_final

def load_excel(file, password=None):
    """
    암호화된 엑셀 파일을 로드하는 함수.
    """
    if password:
        file.seek(0)
        temp_decrypted_file = io.BytesIO()
        officefile = msoffcrypto.OfficeFile(file)
        try:
            officefile.load_key(password=password)
            officefile.decrypt(temp_decrypted_file)
            temp_decrypted_file.seek(0)
            return temp_decrypted_file
        except msoffcrypto.exceptions.InvalidKeyError:
            st.error("잘못된 비밀번호입니다.")
            return None
        except Exception as e:
            st.error(f"엑셀 파일 복호화 중 오류가 발생했습니다: {e}")
            return None
    else:
        return file

def is_encrypted_excel(file):
    """
    엑셀 파일이 암호화되었는지 확인하는 함수.
    """
    file.seek(0)
    try:
        msoffcrypto.OfficeFile(file).is_encrypted()
        file.seek(0)
        return True
    except:
        file.seek(0)
        return False

def process_excel_file_and_style(raw_file_io):
    raw_file_io.seek(0)
    try:
        wb_raw = load_workbook(filename=raw_file_io, keep_vba=False, data_only=True)
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
                        
            # --- 교정과 'Bonding' 텍스트 굵게 처리 로직 추가 ---
            if sheet_name.strip() == "교정" and '진료내역' in header:
                idx = header['진료내역'] - 1
                if len(row) > idx:
                    cell = row[idx]
                    text = str(cell.value).strip().lower()
                    
                    if ('bonding' in text or '본딩' in text) and 'debonding' not in text:
                        cell.font = Font(bold=True)
            # --- 교정과 'Bonding' 텍스트 굵게 처리 로직 추가 끝 ---

    final_output_bytes = io.BytesIO()
    wb_styled.save(final_output_bytes)
    final_output_bytes.seek(0)
    
    return processed_sheets_dfs, final_output_bytes

# --- OCS 분석 함수 추가 ---
def analyze_ocs_data_for_tabs(processed_sheets_dfs, professors_dict):
    """
    업로드된 OCS 데이터를 분석하여 소치, 보존, 교정 현황을 출력합니다.
    """
    with st.spinner("OCS 현황을 분석 중입니다..."):
        # 소아치과 단타 분석
        if '소치' in processed_sheets_dfs:
            st.subheader("소아치과 현황 (단타)")
            df_sochi = processed_sheets_dfs['소치']
            professors = professors_dict.get('소치', [])
            
            # 교수님 진료 제외
            df_non_prof = df_sochi[~df_sochi['예약의사'].isin(professors)]
            
            # 예약시간을 datetime.time 객체로 변환
            try:
                df_non_prof.loc[:, '예약시간'] = pd.to_datetime(df_non_prof['예약시간'], format='%H:%M').dt.time
                morning_count = df_non_prof[df_non_prof['예약시간'] <= datetime.time(12, 30)].shape[0]
                afternoon_count = df_non_prof[df_non_prof['예약시간'] >= datetime.time(12, 50)].shape[0]
            except Exception as e:
                st.error(f"소아치과 시간 분석 오류: {e}")
                morning_count = '오류'
                afternoon_count = '오류'

            total_count = df_non_prof.shape[0]
            st.markdown(f"총 단타 환자 수: **{total_count}명**")
            st.markdown(f"- 오전 진료 (08:00~12:30): **{morning_count}명**")
            st.markdown(f"- 오후 진료 (12:50 이후): **{afternoon_count}명**")
        else:
            st.info("소아치과 시트가 발견되지 않았습니다.")

        # 보존과 단타 분석
        if '보존' in processed_sheets_dfs:
            st.subheader("보존과 현황 (단타)")
            df_bojon = processed_sheets_dfs['보존']
            professors = professors_dict.get('보존', [])
            
            # 교수님 진료 제외
            df_non_prof = df_bojon[~df_bojon['예약의사'].isin(professors)]
            
            try:
                df_non_prof.loc[:, '예약시간'] = pd.to_datetime(df_non_prof['예약시간'], format='%H:%M').dt.time
                morning_count = df_non_prof[df_non_prof['예약시간'] <= datetime.time(12, 30)].shape[0]
                afternoon_count = df_non_prof[df_non_prof['예약시간'] >= datetime.time(12, 50)].shape[0]
            except Exception as e:
                st.error(f"보존과 시간 분석 오류: {e}")
                morning_count = '오류'
                afternoon_count = '오류'

            total_count = df_non_prof.shape[0]
            st.markdown(f"총 단타 환자 수: **{total_count}명**")
            st.markdown(f"- 오전 진료 (08:00~12:30): **{morning_count}명**")
            st.markdown(f"- 오후 진료 (12:50 이후): **{afternoon_count}명**")
        else:
            st.info("보존과 시트가 발견되지 않았습니다.")

        # 교정과 Bonding 갯수 분석
        if '교정' in processed_sheets_dfs:
            st.subheader("교정과 현황 (Bonding)")
            df_kyo = processed_sheets_dfs['교정']

            # 진료내역에 'bonding' 또는 '본딩'이 포함되면서 'debonding' 또는 '탈부착'이 없는 경우만 필터링
            df_bonding = df_kyo[
                ((df_kyo['진료내역'].str.contains('bonding', case=False, na=False)) |
                 (df_kyo['진료내역'].str.contains('본딩', case=False, na=False))) &
                (~(df_kyo['진료내역'].str.contains('debonding', case=False, na=False)) &
                 ~(df_kyo['진료내역'].str.contains('탈부착', case=False, na=False)))
            ]

            try:
                df_bonding.loc[:, '예약시간'] = pd.to_datetime(df_bonding['예약시간'], format='%H:%M').dt.time
                morning_count = df_bonding[df_bonding['예약시간'] <= datetime.time(12, 30)].shape[0]
                afternoon_count = df_bonding[df_bonding['예약시간'] >= datetime.time(12, 50)].shape[0]
            except Exception as e:
                st.error(f"교정과 시간 분석 오류: {e}")
                morning_count = '오류'
                afternoon_count = '오류'

            total_count = df_bonding.shape[0]
            st.markdown(f"총 Bonding 환자 수: **{total_count}명**")
            st.markdown(f"- 오전 Bonding: **{morning_count}명**")
            st.markdown(f"- 오후 Bonding: **{afternoon_count}명**")
        else:
            st.info("교정과 시트가 발견되지 않았습니다.")
            
# #5. Main User Mode
if st.session_state.logged_in and st.session_state.current_user_name != "admin":
    st.title("환자 내원 확인 시스템")
    
    # 탭 구성
    tab1, tab2, tab3 = st.tabs(["진료내역 확인", "OCS 분석 결과", "환자 등록"])

    with tab1:
        st.subheader(f"📅 {st.session_state.current_user_name}님의 내원 환자 정보")
        user_id_safe = sanitize_path(st.session_state.found_user_email)
        
        # 캘린더 서비스 초기화
        creds = load_google_creds_from_firebase(user_id_safe)
        if creds and creds.valid:
            try:
                service = build('calendar', 'v3', credentials=creds)
                st.session_state.google_calendar_service = service
            except Exception as e:
                st.error(f"캘린더 서비스 로드 실패: {e}")
                st.session_state.google_calendar_service = None
        else:
            st.session_state.google_calendar_service = None
            if st.button("Google Calendar 연동하기"):
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
                flow = InstalledAppFlow.from_client_config(client_config, SCOPES, redirect_uri=st.secrets["google_calendar"]["redirect_uri"])
                auth_url, _ = flow.authorization_url(prompt='consent')
                st.markdown(f"[Google Calendar 인증 링크]({auth_url})")

    with tab2:
        st.subheader("📊 OCS 분석 결과")
        if st.session_state.processed_excel_data_dfs:
            analyze_ocs_data_for_tabs(st.session_state.processed_excel_data_dfs, professors_dict)
        else:
            st.info("아직 분석된 OCS 데이터가 없습니다. 관리자가 먼저 엑셀 파일을 업로드해야 합니다.")

    with tab3:
        st.subheader("📝 환자 수동 등록")
        patients_ref_for_user = db.reference(f"patients/{user_id_safe}")
        existing_patient_data = patients_ref_for_user.get()

        if existing_patient_data:
            st.write("---")
            st.write("**이미 등록된 환자 목록**")
            for key, val in existing_patient_data.items():
                col1, col2 = st.columns([0.8, 0.2])
                with col1:
                    st.markdown(f"- **{val['환자명']}** / {val['진료번호']} / {val.get('등록과', '미지정')}")
                
                with col2:
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
            # 교정과는 제외
            if '교정' in departments_for_registration:
                departments_for_registration.remove('교정')
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
                    
                    if st.session_state.google_calendar_service:
                        create_calendar_event(st.session_state.google_calendar_service, name, pid, selected_department,
                                               datetime.date.today().strftime("%Y-%m-%d"), datetime.datetime.now().strftime("%H:%M"), "수동등록", "환자 수동 등록")

                    st.rerun()
#6. User and Admin Login and User Management
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

#7. Admin Mode Functionality
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
            
            # --- 엑셀 파일 이름에서 예약 날짜 정보 추출 (수정) ---
            # 'ocs_0812' -> 8월 12일 -> 2024-08-12
            date_match = re.search(r'_(\d{2})(\d{2})', file_name)
            reservation_date_excel = None
            if date_match:
                month_str = date_match.group(1)
                day_str = date_match.group(2)
                current_year = datetime.datetime.now().year
                reservation_date_excel = f"{current_year}-{month_str}-{day_str}"
            else:
                st.warning("엑셀 파일 이름에서 예약 날짜를 추출할 수 없습니다. 캘린더 일정은 현재 날짜로 설정됩니다.")
                reservation_date_excel = datetime.datetime.now().strftime("%Y-%m-%d")
            
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
                        matched_users.append({"email": user_email, "name": user_display_name, "data": combined_matched_df, "safe_key": uid_safe})

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
                            result = send_email(real_email, df_matched, sender, sender_pw, date_str=reservation_date_excel) # 추출된 날짜 사용
                            if result is True:
                                st.success(f"**{user_match_info['name']}** ({real_email}) 전송 완료")
                            else:
                                st.error(f"**{user_match_info['name']}** ({real_email}) 전송 실패: {result}")
                
                with calendar_col:
                    if st.button("Google Calendar 일정 추가"):
                        for user_match_info in matched_users:
                            user_safe_key = user_match_info['safe_key']
                            user_email = user_match_info['email']
                            user_name = user_match_info['name']
                            df_matched = user_match_info['data']
                            
                            # Check for user-specific Google Calendar credentials
                            creds = load_google_creds_from_firebase(user_safe_key)
                            
                            if creds and creds.valid and not creds.expired:
                                try:
                                    service = build('calendar', 'v3', credentials=creds)
                                    if not df_matched.empty:
                                        for _, row in df_matched.iterrows():
                                            # create_calendar_event 호출 시 날짜, 시간, 의사 이름 인자 전달 (수정)
                                            # 엑셀 파일에 '예약의사' 컬럼이 있다고 가정합니다.
                                            doctor_name = row.get('예약의사', '')
                                            treatment_details = row.get('진료내역', '')
                                            create_calendar_event(service, row['환자명'], row['진료번호'], row.get('시트', ''), 
                                                reservation_date_str=reservation_date_excel, reservation_time_str=row.get('예약시간'), doctor_name=doctor_name, treatment_details=treatment_details)
                                    st.success(f"**{user_name}**님의 캘린더에 일정을 추가했습니다.")
                                except Exception as e:
                                    st.error(f"**{user_name}**님의 캘린더 일정 추가 실패: {e}")
                            else:
                                # If credentials are not found, send an email with the authorization link
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
                                flow = InstalledAppFlow.from_client_config(client_config, SCOPES, redirect_uri=st.secrets["google_calendar"]["redirect_uri"])
                                auth_url, _ = flow.authorization_url(prompt='consent')
                                
                                custom_message = f"""
                                    안녕하세요, {user_name}님.<br><br>
                                    환자 내원 확인 시스템의 구글 캘린더 연동을 위해 인증이 필요합니다.<br>
                                    아래 링크를 클릭하여 권한을 부여해주세요.<br><br>
                                    **<a href="{auth_url}">Google Calendar 인증 링크</a>**<br><br>
                                    감사합니다.
                                """
                                sender = st.secrets["gmail"]["sender"]
                                sender_pw = st.secrets["gmail"]["app_password"]
                                result = send_email(user_email, pd.DataFrame(), sender, sender_pw, custom_message=custom_message)

                                if result is True:
                                    st.success(f"**{user_name}**님 ({user_email})께 캘린더 권한 설정을 위한 메일 전송 완료!")
                                else:
                                    st.error(f"**{user_name}**님 ({user_email})께 메일 전송 실패: {result}")
                            
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

    st.markdown("---")
    st.subheader("🛠️ Administer password")
    admin_password_input = st.text_input("관리자 비밀번호를 입력하세요", type="password", key="admin_password")

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
    
    if st.session_state.admin_password_correct:
        st.markdown("---")
        st.subheader("📦 메일 발송")
        
        all_users_meta = users_ref.get()
        user_list_for_dropdown = [f"{user_info.get('name', '이름 없음')} ({user_info.get('email', '이메일 없음')})" 
                                        for user_info in (all_users_meta.values() if all_users_meta else [])]
        
        select_all_users_button = st.button("모든 사용자 선택/해제", key="select_all_btn")
        if select_all_users_button:
            st.session_state.select_all_users = not st.session_state.select_all_users

        default_selection = user_list_for_dropdown if st.session_state.select_all_users else []

        selected_users_for_mail = st.multiselect("보낼 사용자 선택", user_list_for_dropdown, default=default_selection, key="mail_multiselect")
        
        custom_message = st.text_area("보낼 메일 내용", height=200)
        if st.button("메일 보내기"):
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
                
#8. Regular User Mode
# --- 일반 사용자 모드 ---
else:
    user_id_final = st.session_state.user_id_input_value if st.session_state.email_change_mode or not st.session_state.found_user_email else st.session_state.found_user_email
    firebase_key = sanitize_path(user_id_final) if user_id_final else ""

    if not user_name or not user_id_final:
        st.info("내원 알람 노티를 받을 이메일 주소와 사용자 이름을 입력해주세요.")
        st.stop()

    patients_ref_for_user = db.reference(f"patients/{firebase_key}")

    if not st.session_state.email_change_mode:
        current_user_meta_data = users_ref.child(firebase_key).get()
        if not current_user_meta_data or current_user_meta_data.get("name") != user_name or current_user_meta_data.get("email") != user_id_final:
            users_ref.child(firebase_key).update({"name": user_name, "email": user_id_final})
            st.success(f"사용자 정보가 업데이트되었습니다: {user_name} ({user_id_final})")
        st.session_state.current_firebase_key = firebase_key
        st.session_state.current_user_name = user_name
        st.session_state.found_user_email = user_id_final
    
    # --- 구글 캘린더 연동 섹션 ---
    st.subheader("Google Calendar 연동")
    st.info("환자 등록 시 입력된 이메일 계정의 구글 캘린더에 자동으로 일정이 추가됩니다.")

    if 'google_calendar_service' not in st.session_state:
        st.session_state.google_calendar_service = None
    
    # 구글 캘린더 서비스 객체 가져오기
    google_calendar_service = get_google_calendar_service(firebase_key)
    st.session_state.google_calendar_service = google_calendar_service

    # Display calendar integration status
    if google_calendar_service:
        st.success("✅ 캘린더 추가 기능이 허용되어 있습니다.")
    else:
        # get_google_calendar_service already shows the link
        pass

    st.markdown("---")
    st.subheader(f"{user_name}님의 등록 환자 목록")
    
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
                
                if st.session_state.google_calendar_service:
                     # Manual registration does not have reservation date/time.
                     # The function will use the current time as a fallback.
                    create_calendar_event(st.session_state.google_calendar_service, name, pid, selected_department)
                # ... (rest of the block) ...

                st.rerun()
