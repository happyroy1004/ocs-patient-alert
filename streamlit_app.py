#1. Imports, Validation Functions, and Firebase Initialization
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
import openpyxl # 추가
import datetime # 추가

# Google Calendar API 관련 라이브러리 추가
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import base64

# --- 이메일 유효성 검사 함수 ---
def is_valid_email(email):
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

# 수정 코드 (Revised Code)
# Firebase-safe 경로 변환 (이메일을 Firebase 키로 사용하기 위해)
def sanitize_path(email):
    return email.replace(".", "_dot_").replace("@", "_at_")

# 이메일 주소 복원 (Firebase 안전 키에서 원래 이메일로)
def recover_email(safe_id: str) -> str:
    email = safe_id.replace("_at_", "@").replace("_dot_", ".").replace("_com", ".com")
    return email

# 구글 캘린더 인증 정보를 Firebase에 저장
def save_google_creds_to_firebase(user_id_safe, creds):
    try:
        creds_ref = db.reference(f"users/{user_id_safe}/google_creds")
        creds_ref.set({
            'token': creds.token,
            'refresh_token': creds.refresh_token,
            'token_uri': creds.token_uri,
            'client_id': creds.client_id,
            'client_secret': creds.client_secret,
            'scopes': creds.scopes,
            'id_token': creds.id_token
        })
        return True
    except Exception as e:
        st.error(f"Failed to save Google credentials: {e}")
        return False

# Firebase에서 구글 캘린더 인증 정보를 불러오기
def load_google_creds_from_firebase(user_id_safe):
    try:
        creds_ref = db.reference(f"users/{user_id_safe}/google_creds")
        creds_data = creds_ref.get()
        if creds_data and 'token' in creds_data:
            creds = Credentials(
                token=creds_data.get('token'),
                refresh_token=creds_data.get('refresh_token'),
                token_uri=creds_data.get('token_uri'),
                client_id=creds_data.get('client_id'),
                client_secret=creds_data.get('client_secret'),
                scopes=creds_data.get('scopes'),
                id_token=creds_data.get('id_token')
            )
            return creds
        return None
    except Exception as e:
        st.error(f"Failed to load Google credentials: {e}")
        return None

# --- OCS 분석 관련 함수 추가 ---

# 엑셀 파일 암호화 여부 확인
def is_encrypted_excel(file_path):
    try:
        with openpyxl.open(file_path, read_only=True) as wb:
            return False
    except openpyxl.utils.exceptions.InvalidFileException:
        return True
    except Exception:
        return False

# 엑셀 파일 로드
def load_excel(uploaded_file, password=None):
    try:
        # Streamlit uploaded_file은 io.BytesIO 객체와 유사
        file_io = io.BytesIO(uploaded_file.getvalue())
        wb = load_workbook(file_io, data_only=True)
        return wb, file_io
    except Exception as e:
        st.error(f"엑셀 파일 로드 중 오류 발생: {e}")
        return None, None
    
# 데이터 처리 및 스타일링
def process_excel_file_and_style(file_io):
    try:
        # 파일을 다시 읽어서 raw data를 가져옴
        raw_df = pd.read_excel(file_io)
        
        # DataFrame을 사용하여 각 시트 데이터를 처리
        excel_data_dfs = pd.read_excel(file_io, sheet_name=None)
        
        return excel_data_dfs, raw_df.to_excel(index=False, header=True, engine='xlsxwriter')
    except Exception as e:
        st.error(f"엑셀 데이터 처리 및 스타일링 중 오류 발생: {e}")
        return None, None
    
# OCS 분석 함수
def run_analysis(df_dict, professors_dict):
    analysis_results = {}

    # 딕셔너리로 시트 이름과 부서 맵핑 정의
    sheet_department_map = {
        '소치': '소치',
        '소아치과': '소치',
        '소아 치과': '소치',
        '보존': '보존',
        '보존과': '보존',
        '치과보존과': '보존',
        '교정': '교정',
        '교정과': '교정',
        '치과교정과': '교정'
    }

    # 맵핑된 데이터프레임을 저장할 딕셔너리
    mapped_dfs = {}
    for sheet_name, df in df_dict.items():
        # 공백 제거 및 소문자 변환
        processed_sheet_name = sheet_name.replace(" ", "").lower()
        
        # 맵핑 딕셔너리에서 부서 이름 찾기
        for key, dept in sheet_department_map.items():
            if processed_sheet_name == key.replace(" ", "").lower():
                mapped_dfs[dept] = df
                break

    
    # 소아치과 분석
    if '소치' in mapped_dfs:
        df = mapped_dfs['소치']
        non_professors_df = df[~df['예약의사'].isin(professors_dict.get('소치', []))]
        
        # 🐛 오류 수정: '예약시간'을 문자열로 비교하기 전 유효하지 않은 값 필터링
        non_professors_df['예약시간'] = non_professors_df['예약시간'].astype(str).str.strip()
        non_professors_df = non_professors_df[non_professors_df['예약시간'] != 'nan']
        
        # 오류 수정: '예약시간'을 문자열로 비교
        non_professors_df['예약시간'] = non_professors_df['예약시간'].astype(str).str.strip()
        
        morning_patients = non_professors_df[
            (non_professors_df['예약시간'] >= '08:00') & 
            (non_professors_df['예약시간'] <= '12:50')
        ].shape[0]
        
        afternoon_patients = non_professors_df[
            non_professors_df['예약시간'] >= '13:00'
        ].shape[0]

        # ⚠️ 계산된 값에서 1을 빼는 로직 추가
        if afternoon_patients > 0:
            afternoon_patients -= 1
        analysis_results['소치'] = {'오전': morning_patients, '오후': afternoon_patients}

    # 보존과 분석
    if '보존' in mapped_dfs:
        df = mapped_dfs['보존']
        non_professors_df = df[~df['예약의사'].isin(professors_dict.get('보존', []))]
        
        # 🐛 오류 수정: '예약시간'을 문자열로 비교하기 전 유효하지 않은 값 필터링
        non_professors_df['예약시간'] = non_professors_df['예약시간'].astype(str).str.strip()
        non_professors_df = non_professors_df[non_professors_df['예약시간'] != 'nan']
        
        # 오류 수정: '예약시간'을 문자열로 비교
        non_professors_df['예약시간'] = non_professors_df['예약시간'].astype(str).str.strip()
        
        morning_patients = non_professors_df[
            (non_professors_df['예약시간'] >= '08:00') & 
            (non_professors_df['예약시간'] <= '12:30')
        ].shape[0]
        
        afternoon_patients = non_professors_df[
            non_professors_df['예약시간'] >= '12:50'
        ].shape[0]
# ⚠️ 계산된 값에서 1을 빼는 로직 추가
        if afternoon_patients > 0:
            afternoon_patients -= 1
        analysis_results['보존'] = {'오전': morning_patients, '오후': afternoon_patients}

    # 교정과 분석 (Bonding)
    if '교정' in mapped_dfs:
        df = mapped_dfs['교정']
        bonding_patients_df = df[
            df['진료내역'].str.contains('bonding|본딩', case=False, na=False) &
            ~df['진료내역'].str.contains('debonding', case=False, na=False)
        ]
        
        # 오류 수정: '예약시간'을 문자열로 비교
        bonding_patients_df['예약시간'] = bonding_patients_df['예약시간'].astype(str).str.strip()
        
        morning_bonding_patients = bonding_patients_df[
            (bonding_patients_df['예약시간'] >= '08:00') & 
            (bonding_patients_df['예약시간'] <= '12:30')
        ].shape[0]
        
        afternoon_bonding_patients = bonding_patients_df[
            bonding_patients_df['예약시간'] >= '12:50'
        ].shape[0]
        
        analysis_results['교정'] = {'오전': morning_bonding_patients, '오후': afternoon_bonding_patients}
        
    return analysis_results

# --- 세션 상태 초기화 ---
if "clear" in st.query_params and st.query_params["clear"] == "true":
    st.session_state.clear()
    st.query_params["clear"] = "false"
    st.rerun()

if 'email_change_mode' not in st.session_state:
    st.session_state.email_change_mode = False
if 'last_email_change_time' not in st.session_state:
    st.session_state.last_email_change_time = 0
if 'email_change_sent' not in st.session_state:
    st.session_state.email_change_sent = False
if 'user_logged_in' not in st.session_state:
    st.session_state.user_logged_in = False
if 'found_user_email' not in st.session_state:
    st.session_state.found_user_email = None
if 'user_role' not in st.session_state:
    st.session_state.user_role = 'user'  # 기본값은 'user'
if 'google_creds' not in st.session_state:
    st.session_state['google_creds'] = {}

# 추가된 세션 상태 변수
if 'last_processed_file_name' not in st.session_state:
    st.session_state.last_processed_file_name = None
if 'last_processed_data' not in st.session_state:
    st.session_state.last_processed_data = None

users_ref = db.reference("users")

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

#4. Excel Processing Constants and Functions
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

#5. Streamlit App Start and Session State
# --- Streamlit 애플리케이션 시작 ---
st.set_page_config(layout="wide")

# 제목에 링크 추가 및 초기화 로직
st.markdown("""
    <style>
    .title-link {
        text-decoration: none; color: inherit;
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
if 'google_creds' not in st.session_state:
    st.session_state['google_creds'] = {}

users_ref = db.reference("users")

# 6. User and Admin Login and User Management
import os
import streamlit as st

# 세션 상태 초기화
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'user_logged_in' not in st.session_state:
    st.session_state.user_logged_in = False
if 'current_firebase_key' not in st.session_state:
    st.session_state.current_firebase_key = ""
if 'current_user_name' not in st.session_state:
    st.session_state.current_user_name = ""

# --- 사용 설명서 PDF 다운로드 버튼 ---
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

# 로그인 폼
with st.container():
    st.subheader("로그인")
    user_name = st.text_input("사용자 이름을 입력하세요 (예: 홍길동)")
    password_input = st.text_input("비밀번호를 입력하세요", type="password")
    
    # user_name 변수가 정의된 후에 is_admin_input을 정의
    is_admin_input = (user_name.strip().lower() == "admin")
    
    login_button = st.button("로그인")

if login_button:
    if not user_name:
        st.error("사용자 이름을 입력해주세요.")
    elif not password_input:
        st.error("비밀번호를 입력해주세요.")
    else:
        all_users_meta = users_ref.get()
        found = False
        if all_users_meta:
            for safe_key, user_info in all_users_meta.items():
                if user_info and user_info.get("name") == user_name:
                    # Case 1: 비밀번호가 없는 기존 사용자
                    if "password" not in user_info or user_info.get("password") is None:
                        users_ref.child(safe_key).update({"password": password_input})
                        st.session_state.user_logged_in = True
                        st.session_state.found_user_email = user_info.get("email")
                        st.session_state.current_firebase_key = safe_key
                        st.session_state.current_user_name = user_name
                        st.session_state.logged_in = True
                        st.success(f"**{user_name}**님으로 로그인되었습니다. 새로운 비밀번호를 설정하세요.")
                        found = True
                        break
                    # Case 2: 비밀번호가 있는 사용자
                    elif user_info.get("password") == password_input:
                        st.session_state.user_logged_in = True
                        st.session_state.found_user_email = user_info.get("email")
                        st.session_state.current_firebase_key = safe_key
                        st.session_state.current_user_name = user_name
                        st.session_state.logged_in = True
                        st.success(f"**{user_name}**님으로 로그인되었습니다.")
                        found = True
                        break
                    else:
                        st.error("비밀번호가 틀렸습니다.")
                        found = True
                        break
        
        if not found:
            # 새로운 사용자 등록 로직 (기존과 동일)
            new_email = "" 
            new_firebase_key = sanitize_path(user_name) if user_name else ""
            if new_firebase_key:
                users_ref.child(new_firebase_key).set({
                    "name": user_name,
                    "email": new_email,
                    "password": "1234"
                })
                st.session_state.user_logged_in = True
                st.session_state.found_user_email = new_email
                st.session_state.current_firebase_key = new_firebase_key
                st.session_state.current_user_name = user_name
                st.session_state.logged_in = True
                st.success(f"새로운 사용자 **{user_name}**이(가) 등록되었습니다. 초기 비밀번호는 **1234**입니다.")
# 로그인 상태에 따라 다른 화면 표시
if st.session_state.logged_in:
    st.markdown("---")
    st.success("로그인 성공! 이제 나머지 기능을 이용할 수 있습니다.")
    
    # 비밀번호 수정 기능 추가
    st.subheader("비밀번호 수정")
    new_password = st.text_input("새로운 비밀번호를 입력하세요", type="password")
    confirm_password = st.text_input("새로운 비밀번호를 다시 입력하세요", type="password")
    
    if st.button("비밀번호 변경"):
        if new_password and new_password == confirm_password:
            users_ref.child(st.session_state.current_firebase_key).update({"password": new_password})
            st.success("비밀번호가 성공적으로 변경되었습니다!")
        else:
            st.error("새로운 비밀번호가 일치하지 않습니다.")
            
#7. Admin Mode Functionality
# --- Admin 모드 로그인 처리 ---
if is_admin_input:
    st.session_state.logged_in_as_admin = True
    st.session_state.found_user_email = "admin"
    st.session_state.current_user_name = "admin"
    
    # 두 개의 탭 생성 (추가)
    excel_processor_tab, analysis_tab = st.tabs(['💻 Excel File Processor', '📈 OCS 분석 결과'])
    
    with excel_processor_tab:
        # 엑셀 업로드 섹션 - 비밀번호 없이도 접근 가능
        st.subheader("💻 Excel File Processor")
        uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])
        
        # 엑셀 업로드 로직
        if uploaded_file:
            file_name = uploaded_file.name
            
            uploaded_file.seek(0)
            password = st.text_input("엑셀 파일 비밀번호 입력", type="password") if is_encrypted_excel(uploaded_file) else None
            if is_encrypted_excel(uploaded_file) and not password:
                st.info("암호화된 파일입니다. 비밀번호를 입력해주세요.")
                st.stop()
            
            try:
                xl_object, raw_file_io = load_excel(uploaded_file, password)
                excel_data_dfs, styled_excel_bytes = process_excel_file_and_style(raw_file_io)

                # Firebase에 OCS 분석 결과 영구 저장 (가장 최신값으로 덮어쓰기)
                professors_dict = {
                    '소치': ['김현태', '장기택', '김정욱', '현홍근', '김영재', '신터전', '송지수'],
                    '보존': ['이인복', '금기연', '이우철', '유연지', '서덕규', '이창하', '김선영', '손원준']
                }
                analysis_results = run_analysis(excel_data_dfs, professors_dict)
                
                # 'yyyy-mm-dd' 형식의 키 생성
                today_date_str = datetime.datetime.now().strftime("%Y-%m-%d")
                db.reference("ocs_analysis/latest_result").set(analysis_results)
                db.reference("ocs_analysis/latest_date").set(today_date_str)
                db.reference("ocs_analysis/latest_file_name").set(file_name)
                
                st.session_state.last_processed_data = excel_data_dfs
                st.session_state.last_processed_file_name = file_name
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

    with analysis_tab:
        st.header("📈 OCS 분석 결과")

    # Firebase에서 최신 OCS 분석 결과 로드
        all_analysis_data = db.reference("ocs_analysis").get()
        if all_analysis_data:
            latest_date = sorted(all_analysis_data.keys(), reverse=True)[0]
            latest_file_name = db.reference("ocs_analysis/latest_file_name").get()
            analysis_results = all_analysis_data[latest_date]
            
            st.markdown(f"**<h3 style='text-align: left;'>{latest_file_name} 분석 결과</h3>**", unsafe_allow_html=True)
            st.markdown("---")
            
            # 소아치과 현황
            if '소치' in analysis_results:
                st.subheader("소아치과 현황 (단타)")
                st.info(f"오전: **{analysis_results['소치']['오전']}명**")
                st.info(f"오후: **{analysis_results['소치']['오후']}명**")
            else:
                st.warning("소아치과 데이터가 엑셀 파일에 없습니다.")
            st.markdown("---")
            
            # 보존과 현황
            if '보존' in analysis_results:
                st.subheader("보존과 현황 (단타)")
                st.info(f"오전: **{analysis_results['보존']['오전']}명**")
                st.info(f"오후: **{analysis_results['보존']['오후']}명**")
            else:
                st.warning("보존과 데이터가 엑셀 파일에 없습니다.")
            st.markdown("---")

            # 교정과 현황 (Bonding)
            if '교정' in analysis_results:
                st.subheader("교정과 현황 (Bonding)")
                st.info(f"오전: **{analysis_results['교정']['오전']}명**")
                st.info(f"오후: **{analysis_results['교정']['오후']}명**")
            else:
                st.warning("교정과 데이터가 엑셀 파일에 없습니다.")
            st.markdown("---")
        else:
            st.info("💡 분석 결과가 없습니다. 엑셀 파일을 업로드하면 표시됩니다.")

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
    
    # 두 개의 탭 생성
    analysis_tab, registration_tab = st.tabs(['📈 OCS 분석 결과', '✅ 환자 등록 및 관리'])
    
    with analysis_tab:
        st.header("📈 OCS 분석 결과")

        # Firebase에서 최신 OCS 분석 결과 및 파일명 로드
        analysis_results = db.reference("ocs_analysis/latest_result").get()
        latest_file_name = db.reference("ocs_analysis/latest_file_name").get()

        if analysis_results and latest_file_name:
            st.markdown(f"**<h3 style='text-align: left;'>{latest_file_name} 분석 결과</h3>**", unsafe_allow_html=True)
            st.markdown("---")
            
            # 소아치과 현황
            if '소치' in analysis_results:
                st.subheader("소아치과 현황 (단타)")
                st.info(f"오전: **{analysis_results['소치']['오전']}명**")
                st.info(f"오후: **{analysis_results['소치']['오후']}명**")
            else:
                st.warning("소아치과 데이터가 엑셀 파일에 없습니다.")
            st.markdown("---")
            
            # 보존과 현황
            if '보존' in analysis_results:
                st.subheader("보존과 현황 (단타)")
                st.info(f"오전: **{analysis_results['보존']['오전']}명**")
                st.info(f"오후: **{analysis_results['보존']['오후']}명**")
            else:
                st.warning("보존과 데이터가 엑셀 파일에 없습니다.")
            st.markdown("---")

            # 교정과 현황 (Bonding)
            if '교정' in analysis_results:
                st.subheader("교정과 현황 (Bonding)")
                st.info(f"오전: **{analysis_results['교정']['오전']}명**")
                st.info(f"오후: **{analysis_results['교정']['오후']}명**")
            else:
                st.warning("교정과 데이터가 엑셀 파일에 없습니다.")
            st.markdown("---")

        else:
            st.info("💡 분석 결과가 없습니다. 관리자가 엑셀 파일을 업로드하면 표시됩니다.")
            
    with registration_tab:
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
                    

                    st.rerun()
