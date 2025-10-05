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

# --- 전역 상수 정의 (새로운 구조에 맞게 정의) ---
# 환자 데이터의 진료과 플래그 키 목록 (DB에 저장되는 T/F 플래그)
PATIENT_DEPT_FLAGS = ["보철", "외과", "내과", "소치", "교정", "원진실", "보존"] # 원진실, 보존 추가
# 등록 시 선택할 수 있는 모든 진료과
DEPARTMENTS_FOR_REGISTRATION = ["교정", "내과", "보존", "보철", "소치", "외과", "치주", "원진실"]

# --- 1. Imports, Validation Functions, and Firebase Initialization ---

def is_daily_schedule(file_name):
    """
    파일명이 'ocs_MMDD.xlsx' 또는 'ocs_MMDD.xlsm' 형식인지 확인합니다.
    """
    # 'ocs_날짜(4자리).확장자' 패턴을 찾음 (예: ocs_0815.xlsx)
    pattern = r'^ocs_\d{4}\.(?:xlsx|xlsm)$'
    return re.match(pattern, file_name, re.IGNORECASE) is not None
    

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
        # file_path는 file object (BytesIO)로 간주하고 seek(0)를 수행
        file_path.seek(0)
        # openpyxl은 암호화된 파일을 열 때 InvalidFileException을 발생시킬 수 있음
        try:
             load_workbook(file_path, read_only=True)
             return False
        except openpyxl.utils.exceptions.InvalidFileException:
             # openpyxl이 아닌 msoffcrypto를 사용하여 암호화 확인
             file_path.seek(0)
             return msoffcrypto.OfficeFile(file_path).is_encrypted()
        except Exception:
             file_path.seek(0)
             return False # 기타 오류 시 암호화되지 않은 것으로 가정
    except Exception:
        return False


# 엑셀 파일 로드
def load_excel_openpyxl(uploaded_file, password=None):
    try:
        # Streamlit uploaded_file은 io.BytesIO 객체와 유사
        uploaded_file.seek(0)
        file_io = io.BytesIO(uploaded_file.getvalue())
        wb = load_workbook(file_io, data_only=True)
        return wb, file_io
    except Exception as e:
        st.error(f"엑셀 파일 로드 중 오류 발생: {e}")
        return None, None
    
# 데이터 처리 및 스타일링
def process_excel_file_and_style(file_bytes_io):
    file_bytes_io.seek(0)

    try:
        # data_only=True로 수식 결과만 가져오도록 로드
        wb_raw = load_workbook(filename=file_bytes_io, keep_vba=False, data_only=True) 
    except Exception as e:
        raise ValueError(f"엑셀 워크북 로드 실패: {e}")

    processed_sheets_dfs = {}
    
    # run_analysis에서 사용될 원본 DataFrame 딕셔너리 생성
    file_bytes_io.seek(0)
    all_sheet_dfs = pd.read_excel(file_bytes_io, sheet_name=None)
    
    # 엑셀 시트명 키워드와 진료과 매핑 정의 (Admin 모드와 동일하게 유지)
    sheet_keyword_to_department_map = {
        '치과보철과': '보철', '보철과': '보철', '보철': '보철', '치과교정과' : '교정', '교정과': '교정', '교정': '교정',
        '구강 악안면외과' : '외과', '구강악안면외과': '외과', '외과': '외과', '구강 내과' : '내과', '구강내과': '내과', '내과': '내과',
        '치과보존과' : '보존', '보존과': '보존', '보존': '보존', '소아치과': '소치', '소치': '소치', '소아 치과': '소치',
        '원내생진료센터': '원내생', '원내생': '원내생','원내생 진료센터': '원내생','원진실':'원내생',
        '원스톱 협진센터' : '원스톱', '원스톱협진센터': '원스톱', '원스톱': '원스톱',
        '임플란트 진료센터' : '임플란트', '임플란트진료센터': '임플란트', '임플란트': '임플란트',
        '임플' : '임플란트', '치주과': '치주', '치주': '치주', '임플실': '임플란트', '병리': '병리'
    }

    for sheet_name_raw in wb_raw.sheetnames:
        sheet_name_lower = sheet_name_raw.strip().lower()

        sheet_key = None
        for keyword, department_name in sorted(sheet_keyword_to_department_map.items(), key=lambda item: len(item[0]), reverse=True):
            if keyword.lower() in sheet_name_lower:
                sheet_key = department_name
                break

        if not sheet_key:
            continue

        ws = wb_raw[sheet_name_raw]
        values = list(ws.values)
        while values and (values[0] is None or all((v is None or str(v).strip() == "") for v in values[0])):
            values.pop(0)
        if len(values) < 2:
            continue

        # DataFrame 생성 로직
        df = pd.DataFrame(values)
        if df.empty or df.iloc[0].isnull().all():
             continue

        df.columns = df.iloc[0]
        df = df.drop([0]).reset_index(drop=True)
        df = df.fillna("").astype(str)

        if '예약의사' not in df.columns:
            continue

        df['예약의사'] = df['예약의사'].str.strip().str.replace(" 교수님", "", regex=False)

        professors_dict_v8 = {
            '소치': ['김현태', '장기택', '김정욱', '현홍근', '김영재', '신터전', '송지수'],
            '보존': ['이인복', '금기연', '이우철', '유연지', '서덕규', '이창하', '김선영', '손원준'],
            '외과': ['최진영', '서병무', '명훈', '김성민', '박주영', '양훈주', '한정준', '권익재'],
            '치주': ['구영', '이용무', '설양조', '구기태', '김성태', '조영단'],
            '보철': ['곽재영', '김성균', '임영준', '김명주', '권호범', '여인성', '윤형인', '박지만', '이재현', '조준호'],
            '교정': [], '내과': [], '원진실': [], '원스톱': [], '임플란트': [], '병리': []
        }
        professors_list = professors_dict_v8.get(sheet_key, [])
        
        try:
            processed_df = process_sheet_v8(df, professors_list, sheet_key)
            processed_sheets_dfs[sheet_name_raw] = processed_df
        except Exception as e:
            st.error(f"시트 '{sheet_name_raw}' 처리 중 오류: {e}")
            continue

    if not processed_sheets_dfs:
        return all_sheet_dfs, None

    output_buffer_for_styling = io.BytesIO()
    with pd.ExcelWriter(output_buffer_for_styling, engine='openpyxl') as writer:
        for sheet_name_raw, df in processed_sheets_dfs.items():
            df.to_excel(writer, sheet_name=sheet_name_raw, index=False)

    output_buffer_for_styling.seek(0)
    wb_styled = load_workbook(output_buffer_for_styling, keep_vba=False, data_only=True)

    # 스타일링 로직
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

    # run_analysis에서 사용될 원본 DataFrame 딕셔너리와, 스타일링된 파일의 BytesIO 객체를 반환
    return all_sheet_dfs, final_output_bytes

# OCS 분석 함수
def run_analysis(df_dict, professors_dict):
    analysis_results = {}

    sheet_department_map = {
        '소치': '소치', '소아치과': '소치', '소아 치과': '소치', '보존': '보존', '보존과': '보존', '치과보존과': '보존',
        '교정': '교정', '교정과': '교정', '치과교정과': '교정'
    }

    mapped_dfs = {}
    for sheet_name, df in df_dict.items():
        processed_sheet_name = sheet_name.replace(" ", "").lower()
        for key, dept in sheet_department_map.items():
            if processed_sheet_name == key.replace(" ", "").lower():
                # 컬럼 이름이 유효한지 확인하고 복사본을 저장
                if all(col in df.columns for col in ['예약의사', '예약시간']):
                     mapped_dfs[dept] = df.copy()
                break

    # --- 분석 로직 (기존과 동일) ---
    if '소치' in mapped_dfs:
        df = mapped_dfs['소치']
        non_professors_df = df[~df['예약의사'].isin(professors_dict.get('소치', []))]
        non_professors_df['예약시간'] = non_professors_df['예약시간'].astype(str).str.strip()
        non_professors_df = non_professors_df[non_professors_df['예약시간'] != 'nan']
        morning_patients = non_professors_df[(non_professors_df['예약시간'] >= '08:00') & (non_professors_df['예약시간'] <= '12:50')].shape[0]
        afternoon_patients = non_professors_df[non_professors_df['예약시간'] >= '13:00'].shape[0]
        if afternoon_patients > 0:
            afternoon_patients -= 1
        analysis_results['소치'] = {'오전': morning_patients, '오후': afternoon_patients}

    if '보존' in mapped_dfs:
        df = mapped_dfs['보존']
        non_professors_df = df[~df['예약의사'].isin(professors_dict.get('보존', []))]
        non_professors_df['예약시간'] = non_professors_df['예약시간'].astype(str).str.strip()
        non_professors_df = non_professors_df[non_professors_df['예약시간'] != 'nan']
        morning_patients = non_professors_df[(non_professors_df['예약시간'] >= '08:00') & (non_professors_df['예약시간'] <= '12:30')].shape[0]
        afternoon_patients = non_professors_df[non_professors_df['예약시간'] >= '12:50'].shape[0]
        if afternoon_patients > 0:
            afternoon_patients -= 1
        analysis_results['보존'] = {'오전': morning_patients, '오후': afternoon_patients}

    if '교정' in mapped_dfs:
        df = mapped_dfs['교정']
        bonding_patients_df = df[df['진료내역'].str.contains('bonding|본딩', case=False, na=False) & ~df['진료내역'].str.contains('debonding', case=False, na=False)]
        bonding_patients_df['예약시간'] = bonding_patients_df['예약시간'].astype(str).str.strip()
        morning_bonding_patients = bonding_patients_df[(bonding_patients_df['예약시간'] >= '08:00') & (bonding_patients_df['예약시간'] <= '12:30')].shape[0]
        afternoon_bonding_patients = bonding_patients_df[bonding_patients_df['예약시간'] >= '12:50'].shape[0]
        analysis_results['교정'] = {'오전': morning_bonding_patients, '오후': afternoon_bonding_patients}
        
    return analysis_results


# --- 5. Streamlit App Start and Session State ---
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
if 'user_id_input_value' not in st.session_state:
    st.session_state.user_id_input_value = ""
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
if 'last_processed_file_name' not in st.session_state:
    st.session_state.last_processed_file_name = None
if 'last_processed_data' not in st.session_state:
    st.session_state.last_processed_data = None


users_ref = db.reference("users")
doctor_users_ref = db.reference("doctor_users")

# --- 2. Excel and Email Processing Functions ---
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
            # 암호화되지 않은 경우 파일을 복사하여 반환 (pd.read_excel을 여러 번 사용하기 위함)
            file.seek(0)
            return pd.ExcelFile(file), io.BytesIO(file.read())
    except Exception as e:
        raise ValueError(f"엑셀 로드 또는 복호화 실패: {e}")

# 이메일 전송 함수
def send_email(receiver, rows, sender, password, date_str=None, custom_message=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver

        if custom_message:
            msg['Subject'] = "단체 메일 알림" if date_str is None else f"[치과 내원 알림] {date_str} 예약 내역"
            body = custom_message
        else:
            # 환자 정보 테이블 생성
            subject_prefix = f"{date_str}일에 내원하는 " if date_str else ""
            msg['Subject'] = f"{subject_prefix}등록 환자 내원 알림"
            
            if rows is not None and isinstance(rows, list):
                rows_df = pd.DataFrame(rows)
                html_table = rows_df.to_html(index=False, escape=False)
                
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
                body = f"다음 토탈 환자가 내원예정입니다:<br><br><div class='table-container'>{style}{html_table}</div>"
            else:
                 body = "내원 환자 정보가 없습니다."

        msg.attach(MIMEText(body, 'html'))
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        return str(e)


# --- 3. Google Calendar API Functions ---
SCOPES = ["https://www.googleapis.com/auth/calendar.events"]

# (get_google_calendar_service, create_calendar_event 함수는 내용이 길어 위와 동일하게 유지됨)

# --- 7. Admin 모드 로그인 처리 ---
if st.session_state.get('login_mode') == 'admin_mode':
    st.session_state.logged_in_as_admin = True
    st.session_state.found_user_email = "admin"
    st.session_state.current_user_name = "admin"
    
    st.subheader("💻 Excel File Processor")
    uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])
    
    # 엑셀 시트명 키워드와 진료과 매핑 정의 (다양한 표기법을 통합)
    sheet_keyword_to_department_map = {
    '치과보철과': '보철', '보철과': '보철', '보철': '보철', '치과교정과' : '교정', '교정과': '교정', '교정': '교정',
    '구강 악안면외과' : '외과', '구강악안면외과': '외과', '외과': '외과', '구강 내과' : '내과', '구강내과': '내과', '내과': '내과',
    '치과보존과' : '보존', '보존과': '보존', '보존': '보존', '소아치과': '소치', '소치': '소치', '소아 치과': '소치',
    '원내생진료센터': '원내생', '원내생': '원내생','원내생 진료센터': '원내생','원진실':'원내생',
    '원스톱 협진센터' : '원스톱', '원스톱협진센터': '원스톱', '원스톱': '원스톱',
    '임플란트 진료센터' : '임플란트', '임플란트진료센터': '임플란트', '임플란트': '임플란트',
    '임플' : '임플란트', '치주과': '치주', '치주': '치주', '임플실': '임플란트', '병리': '병리'
    }


    if uploaded_file:
        file_name = uploaded_file.name
        is_daily = is_daily_schedule(file_name)
        if is_daily:
            st.info("✔️ '확정된 당일 일정' 파일로 인식되었습니다. 자동 처리를 시작합니다.")
        else:
            st.info("✔️ '예정된 전체 일정' 파일로 인식되었습니다. 자동 처리를 시작합니다.")
            
        # 암호화 확인 및 비밀번호 입력 처리 (필요시)
        uploaded_file.seek(0)
        password = None
        if is_encrypted_excel(uploaded_file):
            password = st.text_input("⚠️ 암호화된 파일입니다. 비밀번호를 입력해주세요.", type="password", key="auto_exec_password")
            if not password:
                st.info("비밀번호 입력 대기 중...")
                st.stop()
        
        # 1. 파일 처리 및 분석 실행
        try:
            xl_object, raw_file_io = load_excel(uploaded_file, password)
            excel_data_dfs, styled_excel_bytes = process_excel_file_and_style(raw_file_io)
            professors_dict = {
                '소치': ['김현태', '장기택', '김정욱', '현홍근', '김영재', '신터전', '송지수'],
                '보존': ['이인복', '금기연', '이우철', '유연지', '서덕규', '이창하', '김선영', '손원준']
            }
            analysis_results = run_analysis(excel_data_dfs, professors_dict)
            
            # DB에 분석 결과 저장
            today_date_str = datetime.datetime.now().strftime("%Y-%m-%d")
            db.reference("ocs_analysis/latest_result").set(analysis_results)
            db.reference("ocs_analysis/latest_date").set(today_date_str)
            db.reference("ocs_analysis/latest_file_name").set(file_name)
            
            st.session_state.last_processed_data = excel_data_dfs
            st.session_state.last_processed_file_name = file_name

            if excel_data_dfs is None or styled_excel_bytes is None:
                st.warning("엑셀 파일 처리 중 문제가 발생했거나 처리할 데이터가 없습니다. 자동 실행을 중단합니다.")
                st.stop()
                
            output_filename = uploaded_file.name.replace(".xlsx", "_processed.xlsx").replace(".xlsm", "_processed.xlsm")
            st.download_button(
                "처리된 엑셀 다운로드",
                data=styled_excel_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("1. 파일 처리 및 분석이 완료되었습니다. 이제 매칭된 사용자에게 알림을 전송합니다.")
            
        except ValueError as ve:
            st.error(f"파일 처리 실패: {ve}")
            st.stop()
        except Exception as e:
            st.error(f"예상치 못한 오류 발생: {e}")
            st.stop()

        # 2. 학생(일반 사용자) 매칭 및 알림 자동 전송 시작
        st.subheader("📚 학생(일반 사용자)에게 알림 자동 전송")
        
        sender = st.secrets["gmail"]["sender"]
        sender_pw = st.secrets["gmail"]["app_password"]
        all_users_meta = db.reference("users").get()
        all_patients_data = db.reference("patients").get()
        matched_users = []

        # (학생 매칭 로직)
        if all_patients_data:
            patient_dept_to_sheet_map = {
                '보철': ['보철', '임플란트'], '치주': ['치주', '임플란트'], '외과': ['외과', '원스톱', '임플란트'],
                '교정': ['교정'], '내과': ['내과'], '보존': ['보존'], '소치': ['소치'], '원내생': ['원내생'], '병리': ['병리']
            }
            
            for uid_safe, registered_patients_for_this_user in all_patients_data.items():
                user_email = recover_email(uid_safe)
                user_display_name = user_email
                if all_users_meta and uid_safe in all_users_meta and "name" in all_users_meta[uid_safe]:
                    user_display_name = all_users_meta[uid_safe]["name"]
                    user_email = all_users_meta[uid_safe]["email"]
                
                # --- [수정된 환자 데이터 구조를 읽는 로직] ---
                registered_patients_data = []
                if registered_patients_for_this_user:
                    for pid_key, val in registered_patients_for_this_user.items(): 
                        registered_depts = [
                            dept.capitalize() 
                            for dept in PATIENT_DEPT_FLAGS + ['보존', '치주', '원진실'] 
                            if val.get(dept.lower()) is True or val.get(dept.lower()) == 'True' or val.get(dept.lower()) == 'true'
                        ]
                        registered_patients_data.append({
                            "환자명": val.get("환자이름", "").strip(), 
                            "진료번호": pid_key.strip().zfill(8),
                            "등록과_리스트": registered_depts
                        })
                # --- (수정된 환자 데이터 구조를 읽는 로직 끝) ---
                
                matched_rows_for_user = []
                for registered_patient in registered_patients_data:
                    registered_depts = registered_patient["등록과_리스트"] 
                    sheets_to_search = set()
                    for dept in registered_depts:
                        sheets_to_search.update(patient_dept_to_sheet_map.get(dept, [dept]))

                    for sheet_name_excel_raw, df_sheet in excel_data_dfs.items():
                        excel_sheet_department = None
                        for keyword, department_name in sheet_keyword_to_department_map.items():
                            if keyword.lower() in sheet_name_excel_raw.strip().lower():
                                excel_sheet_department = department_name
                                break
                        
                        if excel_sheet_department in sheets_to_search:
                            for _, excel_row in df_sheet.iterrows():
                                excel_patient_name = str(excel_row.get("환자명", "")).strip()
                                excel_patient_pid = str(excel_row.get("진료번호", "")).strip().zfill(8)
                                
                                if (registered_patient["환자명"] == excel_patient_name and
                                        registered_patient["진료번호"] == excel_patient_pid):
                                    matched_row_copy = excel_row.copy()
                                    matched_row_copy["시트"] = sheet_name_excel_raw
                                    matched_row_copy["등록과"] = ", ".join(registered_depts)
                                    matched_rows_for_user.append(matched_row_copy)
                                    break
                
                if matched_rows_for_user:
                    combined_matched_df = pd.DataFrame(matched_rows_for_user)
                    matched_users.append({"email": user_email, "name": user_display_name, "data": combined_matched_df, "safe_key": uid_safe})

        if matched_users:
            st.success(f"매칭된 환자가 있는 {len(matched_users)}명의 학생(사용자)을 발견했습니다. 알림을 전송합니다.")
            
            for user_match_info in matched_users:
                real_email = user_match_info['email']
                df_matched = user_match_info['data']
                user_name = user_match_info['name']
                user_safe_key = user_match_info['safe_key']
                
                # 메일 전송
                email_cols = ['환자명', '진료번호', '예약의사', '진료내역', '예약일시', '예약시간', '등록과']
                df_for_mail = df_matched[[col for col in email_cols if col in df_matched.columns]]
                df_html = df_for_mail.to_html(index=False, escape=False)
                rows_as_dict = df_for_mail.to_dict('records')
                
                email_body = f"""
                <p>안녕하세요, {user_name}님.</p>
                <p>{file_name} 분석 결과, 내원 예정인 환자 진료 정보입니다.</p>
                {df_html}
                <p>확인 부탁드립니다.</p>
                """
                
                try:
                    send_email(receiver=real_email, rows=rows_as_dict, sender=sender, password=sender_pw, custom_message=email_body, date_str=file_name) 
                    st.write(f"✔️ **메일:** {user_name} ({real_email})에게 전송 완료.")
                except Exception as e:
                    st.error(f"❌ **메일:** {user_name} ({real_email})에게 전송 실패: {e}")

                # 캘린더 등록
                creds = load_google_creds_from_firebase(user_safe_key)
                if creds and creds.valid and not creds.expired:
                    try:
                        service = build('calendar', 'v3', credentials=creds)
                        for _, row in df_matched.iterrows():
                            reservation_date_raw = row.get('예약일시', '')
                            reservation_time_raw = row.get('예약시간', '')
                            
                            if reservation_date_raw and reservation_time_raw:
                                full_datetime_str = f"{str(reservation_date_raw).strip()} {str(reservation_time_raw).strip()}"
                                reservation_datetime = datetime.datetime.strptime(full_datetime_str, '%Y/%m/%d %H:%M')
                                
                                event_prefix = "✨ 내원 : " if is_daily else "❓내원 : "
                                event_title = f"{event_prefix}{row.get('환자명', 'N/A')} ({row.get('등록과', 'N/A')}, {row.get('예약의사', 'N/A')})"
                                event_description = f"환자명 : {row.get('환자명', 'N/A')}\n진료번호 : {row.get('진료번호', 'N/A')}\n진료내역 : {row.get('진료내역', 'N/A')}"
                                
                                service.events().insert(calendarId='primary', body={
                                    'summary': event_title, 'location': row.get('진료번호', ''), 'description': event_description,
                                    'start': {'dateTime': reservation_datetime.replace(tzinfo=datetime.timezone(datetime.timedelta(hours=9))).isoformat(), 'timeZone': 'Asia/Seoul'},
                                    'end': {'dateTime': (reservation_datetime + datetime.timedelta(minutes=30)).replace(tzinfo=datetime.timezone(datetime.timedelta(hours=9))).isoformat(), 'timeZone': 'Asia/Seoul'}
                                }).execute()
                                
                        st.write(f"✔️ **캘린더:** {user_name}에게 일정 추가 완료.")
                    except Exception as e:
                        st.warning(f"⚠️ **캘린더:** {user_name} 일정 추가 중 오류: 인증/권한 문제일 수 있습니다.")
                else:
                    st.warning(f"⚠️ **캘린더:** {user_name}님은 Google Calendar 계정이 연동되지 않았습니다.")

        else:
            st.info("매칭된 학생(사용자)이 없습니다. 알림 전송을 건너뜁니다.")
        
        # 3. 치과의사 매칭 및 알림 자동 전송 시작
        st.subheader("🧑‍⚕️ 치과의사에게 알림 자동 전송")
        
        all_doctors_meta = db.reference("doctor_users").get()
        doctors = []
        if all_doctors_meta:
            for safe_key, user_info in all_doctors_meta.items():
                if user_info:
                    doctors.append({
                        "safe_key": safe_key,
                        "name": user_info.get("name", "이름 없음"),
                        "email": user_info.get("email", "이메일 없음"),
                        "department": user_info.get("department", "미지정")
                    })
        
        matched_doctors = []
        doctor_dept_to_sheet_map = {
            '보철': ['보철', '임플란트'], '치주': ['치주', '임플란트'], '외과': ['외과', '원스톱', '임플란트'],
            '교정': ['교정'], '내과': ['내과'], '보존': ['보존'], '소치': ['소치'], '원내생': ['원내생'], '병리': ['병리']
        }

        if doctors and excel_data_dfs:
            for res in doctors:
                found_match = False
                doctor_dept = res['department']
                sheets_to_search = doctor_dept_to_sheet_map.get(doctor_dept, [doctor_dept])

                for sheet_name_excel_raw, df_sheet in excel_data_dfs.items():
                    excel_sheet_name_lower = sheet_name_excel_raw.strip().lower().replace(' ', '')
                    
                    excel_sheet_department = None
                    for keyword, department_name in sorted(sheet_keyword_to_department_map.items(), key=lambda item: len(item[0]), reverse=True):
                        if keyword.lower().replace(' ', '') in excel_sheet_name_lower:
                            excel_sheet_department = department_name
                            break
                    if not excel_sheet_department:
                        continue
                    
                    if excel_sheet_department in sheets_to_search:
                        for _, excel_row in df_sheet.iterrows():
                            excel_doctor_name_from_row = str(excel_row.get('예약의사', '')).strip().replace("'", "").replace("‘", "").replace("’", "").strip()
                            
                            if excel_doctor_name_from_row == res['name']:
                                matched_doctors.append(res)
                                found_match = True
                                break 
                    if found_match:
                        break
        
        if matched_doctors:
            st.success(f"매칭된 진료가 있는 {len(matched_doctors)}명의 치과의사를 발견했습니다. 알림을 전송합니다.")

            for res in matched_doctors:
                matched_rows_for_doctor = []
                doctor_dept = res['department']
                sheets_to_search = doctor_dept_to_sheet_map.get(doctor_dept, [doctor_dept])

                if excel_data_dfs:
                    for sheet_name_excel_raw, df_sheet in excel_data_dfs.items():
                        excel_sheet_name_lower = sheet_name_excel_raw.strip().lower().replace(' ', '')
                        excel_sheet_department = None
                        for keyword, department_name in sorted(sheet_keyword_to_department_map.items(), key=lambda item: len(item[0]), reverse=True):
                            if keyword.lower().replace(' ', '') in excel_sheet_name_lower:
                                excel_sheet_department = department_name
                                break
                        
                        if excel_sheet_department in sheets_to_search:
                            for _, excel_row in df_sheet.iterrows():
                                excel_doctor_name_from_row = str(excel_row.get('예약의사', '')).strip().replace("'", "").replace("‘", "").replace("’", "").strip()
                                
                                if excel_doctor_name_from_row == res['name']:
                                    matched_rows_for_doctor.append(excel_row.copy())
                            
                    if matched_rows_for_doctor:
                        df_matched = pd.DataFrame(matched_rows_for_doctor)
                        latest_file_name = db.reference("ocs_analysis/latest_file_name").get()
                        
                        # 메일 전송
                        email_cols = ['환자명', '진료번호', '예약의사', '진료내역', '예약일시', '예약시간']
                        df_for_mail = df_matched[[col for col in email_cols if col in df_matched.columns]]
                        df_html = df_for_mail.to_html(index=False, escape=False)
                        rows_as_dict = df_for_mail.to_dict('records')

                        email_body = f"""
                        <p>안녕하세요, {res['name']} 치과의사님.</p>
                        <p>{latest_file_name}에서 가져온 내원할 환자 정보입니다.</p>
                        {df_html}
                        <p>확인 부탁드립니다.</p>
                        """
                        try:
                            send_email(receiver=res['email'], rows=rows_as_dict, sender=sender, password=sender_pw, custom_message=email_body, date_str=latest_file_name)
                            st.write(f"✔️ **메일:** Dr. {res['name']}에게 전송 완료!")
                        except Exception as e:
                            st.error(f"❌ **메일:** Dr. {res['name']}에게 전송 실패: {e}")

                        # 캘린더 등록
                        creds = load_google_creds_from_firebase(res['safe_key'])
                        if creds and creds.valid and not creds.expired:
                            try:
                                service = build('calendar', 'v3', credentials=creds)
                                for _, row in df_matched.iterrows():
                                    reservation_date_str = row.get('예약일시', '')
                                    reservation_time_str = row.get('예약시간', '')
                                    if reservation_date_str and reservation_time_str:
                                        full_datetime_str = f"{str(reservation_date_str).strip()} {str(reservation_time_str).strip()}"
                                        reservation_datetime = datetime.datetime.strptime(full_datetime_str, '%Y/%m/%d %H:%M')
                                        
                                        event_prefix = "✨:" if is_daily else "?:"
                                        event_title = f"{event_prefix}{row.get('환자명', 'N/A')}({row.get('진료번호', 'N/A')})"
                                        event_description = f"환자명: {row.get('환자명', 'N/A')}\n진료번호: {row.get('진료번호', 'N/A')}\n진료내역: {row.get('진료내역', 'N/A')}"
                                        
                                        service.events().insert(calendarId='primary', body={
                                            'summary': event_title, 'location': row.get('진료번호', ''), 'description': event_description,
                                            'start': {'dateTime': reservation_datetime.replace(tzinfo=datetime.timezone(datetime.timedelta(hours=9))).isoformat(), 'timeZone': 'Asia/Seoul'},
                                            'end': {'dateTime': (reservation_datetime + datetime.timedelta(minutes=30)).replace(tzinfo=datetime.timezone(datetime.timedelta(hours=9))).isoformat(), 'timeZone': 'Asia/Seoul'}
                                        }).execute()
                                
                                st.write(f"✔️ **캘린더:** Dr. {res['name']}에게 일정 추가 완료.")
                            except Exception as e:
                                st.warning(f"⚠️ **캘린더:** Dr. {res['name']} 일정 추가 중 오류: {e}")
                        else:
                            st.warning(f"⚠️ **캘린더:** Dr. {res['name']}님은 Google Calendar 계정이 연동되지 않았습니다.")
                    else:
                        st.warning(f"Dr. {res['name']} 치과의사의 매칭 데이터가 엑셀 파일에 없습니다.")
        else:
            st.info("매칭된 치과의사 계정이 없습니다. 알림 전송을 건너뜁니다.")


        st.success("🎉 모든 자동 알림 프로세스가 완료되었습니다.")
        st.stop() # 자동 처리가 완료되면 Streamlit이 재실행되지 않도록 멈춥니다.
    
    # 4. 수동 관리자 탭 및 비밀번호 변경 폼은 파일 업로드 if 블록 밖에 유지
    if not uploaded_file:
         st.info("엑셀 파일을 업로드하면 자동 분석 및 알림 전송이 시작됩니다.")

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
        if st.session_state.admin_password_correct:
            st.markdown("---")
    
            tab1, tab2 = st.tabs(["일반 사용자 관리", "치과의사 관리"])
            
            # 탭 1: 일반 사용자 관리
            with tab1:
                st.subheader("📦 일반 사용자 메일 발송 & 삭제")
                
                all_users_meta = users_ref.get()
                user_list_for_dropdown = [f"{user_info.get('name', '이름 없음')} ({user_info.get('email', '이메일 없음')})"
                                          for user_info in (all_users_meta.values() if all_users_meta else [])]
                
                # --- 메일 발송 로직 (일반 사용자) ---
                if 'select_all_users_tab1' not in st.session_state:
                    st.session_state.select_all_users_tab1 = False
                
                select_all_users_button_tab1 = st.button("모든 사용자 선택/해제", key="select_all_btn_tab1")
                if select_all_users_button_tab1:
                    st.session_state.select_all_users_tab1 = not st.session_state.select_all_users_tab1
                    st.rerun()
                
                default_selection_tab1 = user_list_for_dropdown if st.session_state.select_all_users_tab1 else []
                selected_users_for_mail_tab1 = st.multiselect("보낼 사용자 선택", user_list_for_dropdown, default=default_selection_tab1, key="mail_multiselect_tab1")
                
                custom_message_tab1 = st.text_area("보낼 메일 내용", height=200, key="mail_content_tab1")
                
                if st.button("메일 보내기", key="send_mail_button_tab1"):
                    if custom_message_tab1 and selected_users_for_mail_tab1:
                        # 메일 전송 로직은 기존과 동일
                        sender = st.secrets["gmail"]["sender"]
                        sender_pw = st.secrets["gmail"]["app_password"]
                        
                        email_list = []
                        for user_str in selected_users_for_mail_tab1:
                            match = re.search(r'\((.*?)\)', user_str)
                            if match:
                                email_list.append(match.group(1))
                
                        if email_list:
                            with st.spinner("메일 전송 중..."):
                                success_count = 0
                                for email in email_list:
                                    # 단체 메일은 rows 없이 custom_message만 전송
                                    result = send_email(
                                        receiver=email,
                                        rows=None, 
                                        sender=sender,
                                        password=sender_pw,
                                        date_str=None, # 이 기능에서는 사용되지 않으므로 None
                                        custom_message=custom_message_tab1
                                    )
                                    if result is True:
                                        success_count += 1
                                        st.success(f"{email}로 메일 전송 완료!")
                                    else:
                                        st.error(f"{email}로 메일 전송 실패! 오류: {result}")
                                st.success(f"총 {success_count}건의 메일 전송 완료!")
                        else:
                            st.warning("메일 내용을 입력했으나, 선택된 사용자가 없습니다. 전송이 진행되지 않았습니다.")
                    else:
                        st.warning("메일 내용과 대상을 모두 선택해주세요.")
                
                st.markdown("---")
                st.subheader("🗑️ 일반 사용자 삭제")
                
                # --- 사용자 삭제 로직 (일반 사용자) ---
                if 'delete_confirm_tab1' not in st.session_state:
                    st.session_state.delete_confirm_tab1 = False
                if 'users_to_delete_tab1' not in st.session_state:
                    st.session_state.users_to_delete_tab1 = []
                
                if not st.session_state.delete_confirm_tab1:
                    users_to_delete_tab1 = st.multiselect("삭제할 사용자 선택", user_list_for_dropdown, key="delete_user_multiselect_tab1")
                    if st.button("선택한 사용자 삭제", key="delete_button_tab1"):
                        if users_to_delete_tab1:
                            st.session_state.delete_confirm_tab1 = True
                            st.session_state.users_to_delete_tab1 = users_to_delete_tab1
                            st.rerun()
                        else:
                            st.warning("삭제할 사용자를 선택해주세요.")
                else:
                    st.warning("정말로 선택한 사용자를 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.")
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button("예, 삭제합니다", key="confirm_delete_tab1"):
                            with st.spinner('삭제 중...'):
                                for user_to_del_str in st.session_state.users_to_delete_tab1:
                                    match = re.search(r'\((.*?)\)', user_to_del_str)
                                    if match:
                                        email_to_del = match.group(1)
                                        safe_key_to_del = sanitize_path(email_to_del)
                                        
                                        # 일반 사용자 데이터베이스에서 삭제
                                        users_ref.child(safe_key_to_del).delete()
                                        # 환자 데이터도 함께 삭제
                                        db.reference(f"patients/{safe_key_to_del}").delete()
                                        
                                st.success(f"사용자 {', '.join(st.session_state.users_to_delete_tab1)} 삭제 완료.")
                                st.session_state.delete_confirm_tab1 = False
                                st.session_state.users_to_delete_tab1 = []
                                st.rerun()
                    with col2:
                        if st.button("아니오, 취소합니다", key="cancel_delete_tab1"):
                            st.session_state.delete_confirm_tab1 = False
                            st.session_state.users_to_delete_tab1 = []
                            st.rerun()
        
            # 탭 2: 치과의사 사용자 관리
            with tab2:
                st.subheader("📦 치과의사 메일 발송 & 삭제")
                
                all_doctors_meta = doctor_users_ref.get()
                doctor_list_for_dropdown = [f"{doc_info.get('name', '이름 없음')} ({doc_info.get('email', '이메일 없음')})"
                                            for doc_info in (all_doctors_meta.values() if all_doctors_meta else [])]
                
                # --- 메일 발송 로직 (치과의사) ---
                if 'select_all_users_tab2' not in st.session_state:
                    st.session_state.select_all_users_tab2 = False
                
                select_all_users_button_tab2 = st.button("모든 사용자 선택/해제", key="select_all_btn_tab2")
                if select_all_users_button_tab2:
                    st.session_state.select_all_users_tab2 = not st.session_state.select_all_users_tab2
                    st.rerun()
                
                default_selection_tab2 = doctor_list_for_dropdown if st.session_state.select_all_users_tab2 else []
                selected_users_for_mail_tab2 = st.multiselect("보낼 사용자 선택", doctor_list_for_dropdown, default=default_selection_tab2, key="mail_multiselect_tab2")
                
                custom_message_tab2 = st.text_area("보낼 메일 내용", height=200, key="mail_content_tab2")
                
                if st.button("메일 보내기", key="send_mail_button_tab2"):
                    if custom_message_tab2 and selected_users_for_mail_tab2:
                        # 메일 전송 로직
                        sender = st.secrets["gmail"]["sender"]
                        sender_pw = st.secrets["gmail"]["app_password"]
                
                        email_list = []
                        for user_str in selected_users_for_mail_tab2:
                            match = re.search(r'\((.*?)\)', user_str)
                            if match:
                                email_list.append(match.group(1))
                
                        if email_list:
                            with st.spinner("메일 전송 중..."):
                                success_count = 0
                                for email in email_list:
                                    result = send_email(
                                        receiver=email,
                                        rows=None,
                                        sender=sender,
                                        password=sender_pw,
                                        date_str=None, # 이 기능에서는 사용되지 않으므로 None
                                        custom_message=custom_message_tab2
                                    )
                                    if result is True:
                                        success_count += 1
                                        st.success(f"{email}로 메일 전송 완료!")
                                    else:
                                        st.error(f"{email}로 메일 전송 실패! 오류: {result}")
                                st.success(f"총 {success_count}건의 메일 전송 완료!")
                        else:
                            st.warning("메일 내용을 입력했으나, 선택된 사용자가 없습니다. 전송이 진행되지 않았습니다.")
                    else:
                        st.warning("메일 내용과 대상을 모두 선택해주세요.")
                        
                st.markdown("---")
                st.subheader("🗑️ 치과의사 삭제")
                
                # --- 사용자 삭제 로직 (치과의사) ---
                if 'delete_confirm_tab2' not in st.session_state:
                    st.session_state.delete_confirm_tab2 = False
                if 'users_to_delete_tab2' not in st.session_state:
                    st.session_state.users_to_delete_tab2 = []
                
                if not st.session_state.delete_confirm_tab2:
                    users_to_delete_tab2 = st.multiselect("삭제할 사용자 선택", doctor_list_for_dropdown, key="delete_user_multiselect_tab2")
                    if st.button("선택한 사용자 삭제", key="delete_button_tab2"):
                        if users_to_delete_tab2:
                            st.session_state.delete_confirm_tab2 = True
                            st.session_state.users_to_delete_tab2 = users_to_delete_tab2
                            st.rerun()
                        else:
                            st.warning("삭제할 사용자를 선택해주세요.")
                else:
                    st.warning("정말로 선택한 사용자를 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.")
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button("예, 삭제합니다", key="confirm_delete_tab2"):
                            with st.spinner('삭제 중...'):
                                for user_to_del_str in st.session_state.users_to_delete_tab2:
                                    match = re.search(r'\((.*?)\)', user_to_del_str)
                                    if match:
                                        email_to_del = match.group(1)
                                        safe_key_to_del = sanitize_path(email_to_del)
                                        
                                        # 치과의사 데이터베이스에서만 삭제
                                        doctor_users_ref.child(safe_key_to_del).delete()
                                        
                                st.success("선택한 치과의사 삭제 완료.")
                                st.session_state.delete_confirm_tab2 = False
                                st.session_state.users_to_delete_tab2 = []
                                st.rerun()
                    with col2:
                        if st.button("아니오, 취소합니다", key="cancel_delete_tab2"):
                            st.session_state.delete_confirm_tab2 = False
                            st.session_state.users_to_delete_tab2 = []
                            st.rerun()

    elif admin_password_input and admin_password_input != secret_admin_password:
        st.error("비밀번호가 틀렸습니다.")
        st.session_state.admin_password_correct = False
        
    # --- 8. Regular User Mode ---
# --- 일반 사용자 & 치과의사 모드 ---
import streamlit as st
import pandas as pd
import io
import re
                        
if st.session_state.get('login_mode') in ['user_mode', 'new_user_registration', 'doctor_mode', 'new_doctor_registration', 'doctor_name_input']:
    user_name = st.session_state.get('current_user_name', "")
    user_id_final = st.session_state.get('found_user_email', "")
    firebase_key = st.session_state.get('current_firebase_key', "")
    user_role = st.session_state.get('current_user_role', 'user')
    
    # 올바른 데이터베이스 참조를 결정
    if user_role == 'doctor':
        target_users_ref = doctor_users_ref
    else:
        target_users_ref = users_ref
    
    if firebase_key: # firebase_key가 있을 때만 이 코드를 실행합니다.
        
        # 이메일 주소 변경 기능으로 인해 유저 정보가 바뀔 수 있으므로 매번 업데이트
        if not st.session_state.get('email_change_mode'):
            current_user_meta_data = target_users_ref.child(firebase_key).get()
            if not current_user_meta_data or current_user_meta_data.get("name") != user_name or current_user_meta_data.get("email") != user_id_final:
                # name, email을 업데이트 (다른 필드는 유지)
                update_data = {"name": user_name, "email": user_id_final}
                target_users_ref.child(firebase_key).update(update_data)
                # st.success(f"사용자 정보가 업데이트되었습니다: {user_name} ({user_id_final})")
            st.session_state.current_firebase_key = firebase_key
            st.session_state.current_user_name = user_name
            st.session_state.found_user_email = user_id_final
            st.session_state.current_user_role = user_role

        if not user_name or not user_id_final:
            st.info("내원 알람 노티를 받을 이메일 주소와 사용자 이름을 입력해주세요.")
            st.stop()
    
        if st.session_state.get('login_mode') == 'doctor_mode' or st.session_state.get('login_mode') == 'new_doctor_registration':
            st.header(f"🧑‍⚕️Dr. {user_name}")
            st.subheader("🗓️ Google Calendar 연동")
            st.info("구글 캘린더와 연동하여 내원 일정을 자동으로 등록할 수 있습니다.")

            if 'google_calendar_service' not in st.session_state:
                st.session_state.google_calendar_service = None
            
            # firebase_key가 존재할 때만 함수를 호출하도록 수정
            if firebase_key:
                try:
                    google_calendar_service = get_google_calendar_service(firebase_key)
                    st.session_state.google_calendar_service = google_calendar_service
                except Exception as e:
                    st.error(f"❌ Google Calendar 서비스 로딩에 실패했습니다: {e}")
                    st.info("로그인/인증 정보가 올바른지 확인해주세요.")
                    st.session_state.google_calendar_service = None

            if st.session_state.google_calendar_service:
                st.success("✅ 캘린더 추가 기능이 허용되어 있습니다.")
            else:
                pass

            st.markdown("---")
            st.header("🔑 비밀번호 변경")
            new_password = st.text_input("새 비밀번호를 입력하세요", type="password", key="res_new_password_input")
            confirm_password = st.text_input("새 비밀번호를 다시 입력하세요", type="password", key="res_confirm_password_input")
            
            if st.button("비밀번호 변경", key="res_password_change_btn"):
                if not new_password or not confirm_password:
                    st.error("새 비밀번호와 확인용 비밀번호를 모두 입력해주세요.")
                elif new_password != confirm_password:
                    st.error("새 비밀번호가 일치하지 않습니다. 다시 확인해주세요.")
                else:
                    try:
                        doctor_users_ref.child(st.session_state.current_firebase_key).update({"password": new_password})
                        st.success("🎉 비밀번호가 성공적으로 변경되었습니다!")
                    except Exception as e:
                        st.error(f"비밀번호 변경 중 오류가 발생했습니다: {e}")
            
        elif st.session_state.get('login_mode') in ['user_mode', 'new_user_registration']:
            patients_ref_for_user = db.reference(f"patients/{firebase_key}")

            registration_tab, analysis_tab = st.tabs(['✅ 환자 등록 및 관리', '📈 OCS 분석 결과'])
        
            with registration_tab:
                st.subheader("Google Calendar 연동")
                st.info("환자 등록 시 입력된 이메일 계정의 구글 캘린더에 자동으로 일정이 추가됩니다.")
                if 'google_calendar_service' not in st.session_state:
                    st.session_state.google_calendar_service = None
                
                try:
                    google_calendar_service = get_google_calendar_service(firebase_key)
                    st.session_state.google_calendar_service = google_calendar_service
                except Exception as e:
                    st.error(f"❌ Google Calendar 서비스 로딩에 실패했습니다: {e}")
                    st.session_state.google_calendar_service = None
        
                if st.session_state.google_calendar_service:
                    st.success("✅ 캘린더 추가 기능이 허용되어 있습니다.")
                else:
                    pass
        
                st.markdown("---")
                st.subheader(f"{user_name}님의 토탈 환자 목록")
                existing_patient_data = patients_ref_for_user.get()
        
                if existing_patient_data:
                    patient_list = list(existing_patient_data.items())
                    # 유효성 검사: 데이터가 딕셔너리 형태가 아닌 손상된 데이터를 제거
                    valid_patient_list = [item for item in patient_list if isinstance(item[1], dict)]
                    # --- [핵심 변경: 진료과 플래그 우선순위 정렬] ---
                    # 1. 소치(0) > 보철(1) > 내과(2) > 교정(3) 순서로 높은 우선순위를 부여
                    # 2. 동일 우선순위 내에서는 환자이름 순으로 정렬
                    sorted_patient_list = sorted(valid_patient_list, key=lambda item: (
                        0 if item[1].get('소치', False) else
                        1 if item[1].get('외과', False) else
                        2 if item[1].get('내과', False) else
                        3 if item[1].get('교정', False) else
                        4 if item[1].get('보철', False) else
                        5, # 나머지 과목 (원진실, 보존 등)은 4순위로 밀립니다.
                        item[1].get('환자이름', 'zzz')
                    ))
                    cols_count = 3
                    cols = st.columns(cols_count)
        
                    for idx, (pid_key, val) in enumerate(sorted_patient_list): # pid_key가 환자번호
                        with cols[idx % cols_count]:
                            with st.container(border=True):
                                info_col, btn_col = st.columns([4, 1])
                                with info_col:
                                    # True인 진료과만 추출하여 표시
                                    registered_depts = [
                                        dept.capitalize() 
                                        for dept in PATIENT_DEPT_FLAGS + ['보존', '치주', '원진실'] # 모든 가능한 과
                                        if val.get(dept.lower()) is True or val.get(dept.lower()) == 'True' or val.get(dept.lower()) == 'true'
                                    ]

                                    depts_str = ", ".join(registered_depts) if registered_depts else "미지정"
                                    
                                    st.markdown(f"**{val.get('환자이름', '이름 없음')}** / {pid_key} / {depts_str}") # pid_key는 진료번호
                                with btn_col:
                                    if st.button("X", key=f"delete_button_{pid_key}"): # pid_key 사용
                                        patients_ref_for_user.child(pid_key).delete()
                                        st.rerun()
                else:
                    st.info("등록된 환자가 없습니다.")
                st.markdown("---")

                # --- 환자 정보 대량 등록 섹션 추가 (구조 변경 반영) ---
                st.subheader("📋 환자 정보 대량 등록")
                st.markdown("엑셀에서 **환자명, 진료번호, 등록과** 순서의 데이터를 그대로 붙여넣어주세요.")
                st.markdown("예시: 홍길동	1046769	보존")
                st.markdown(f"등록 가능 과: {', '.join(DEPARTMENTS_FOR_REGISTRATION)}")
                
                
                paste_area = st.text_area("", height=200, placeholder="여기에 엑셀 데이터를 붙여넣으세요.")
                
                if st.button("붙여넣은 환자 등록"):
                    if paste_area:
                        try:
                            data_io = io.StringIO(paste_area)
                            
                            # header=None으로 헤더가 없음을 명시하고, names로 열 이름을 수동 지정
                            df = pd.read_csv(data_io, sep='\s+', header=None, names=['환자명', '진료번호', '등록과'])
                            
                            # 기존 환자 데이터 가져오기
                            existing_patient_data = patients_ref_for_user.get()
                            if not existing_patient_data:
                                existing_patient_data = {}
                            
                            success_count = 0
                            for index, row in df.iterrows():
                                name = str(row["환자명"]).strip()
                                pid = str(row["진료번호"]).strip()
                                department = str(row["등록과"]).strip()
                                
                                if not name or not pid or not department:
                                    st.warning(f"{index+1}번째 행: 정보가 누락되어 건너킵니다.")
                                    continue
                                
                                pid_key = pid.strip() # 진료번호를 키로 사용
                                dept_key_lower = department.lower()
                                
                                # 1. 새 데이터 딕셔너리 생성 및 초기화 (환자이름, 진료번호, 5개 진료과 플래그)
                                # 기존 데이터가 있으면 불러와서 업데이트
                                new_patient_data = existing_patient_data.get(pid_key, {
                                    "환자이름": name,
                                    "진료번호": pid # 키로 사용되지만 데이터 내부에도 포함
                                })
                                
                                # 진료과 플래그 초기화 및 업데이트 (기존 데이터와 병합)
                                for dept_flag in PATIENT_DEPT_FLAGS + ['보존', '치주', '원진실']: # 모든 가능한 과 플래그 초기화
                                    lower_dept = dept_flag.lower()
                                    if lower_dept not in new_patient_data:
                                        new_patient_data[lower_dept] = False

                                # 2. 등록과에 해당하는 플래그 True로 설정
                                if dept_key_lower in new_patient_data:
                                    new_patient_data[dept_key_lower] = True
                                else:
                                    st.warning(f"{name} ({pid}): 알 수 없는 진료과 '{department}'가 입력되었습니다. 플래그를 설정하지 않습니다.")

                                # 3. 환자번호(pid)를 키로 사용하여 데이터 저장 (덮어쓰기/업데이트)
                                patients_ref_for_user.child(pid_key).set(new_patient_data) # <--- **핵심 변경점**
                                success_count += 1
                                st.success(f"{name} ({pid}) [{department}] 환자 등록/업데이트 완료")
                                
                            
                            if success_count > 0:
                                st.success(f"총 {success_count}명의 환자 정보 등록/업데이트가 완료되었습니다.")
                            st.rerun()
                            
                        except pd.errors.ParserError:
                            st.error("잘못된 형식입니다. 엑셀이나 구글 스프레드시트의 표를 복사하여 붙여넣었는지 확인해주세요. 데이터 구분자가 탭(Tab)이어야 합니다.")
                        except Exception as e:
                            st.error(f"예상치 못한 오류: {e}")
                    else:
                        st.warning("붙여넣을 환자 정보가 없습니다.")
                        
                st.markdown("---")
        
                # --- 환자 정보 일괄 삭제 섹션 추가 (구조 변경 반영) ---
                st.subheader("🗑️ 환자 정보 일괄 삭제")
                
                if 'delete_patient_confirm' not in st.session_state:
                    st.session_state.delete_patient_confirm = False
                if 'patients_to_delete' not in st.session_state:
                    st.session_state.patients_to_delete = []
                if 'select_all_mode' not in st.session_state:
                    st.session_state.select_all_mode = False
                
                all_patients_meta = patients_ref_for_user.get()
                patient_list_for_dropdown = []
                patient_key_map = {}
                
                if all_patients_meta:
                    for pid_key, value in all_patients_meta.items(): # pid_key는 진료번호
                        # True인 진료과를 모두 추출하여 표시
                        registered_depts = [
                            dept.capitalize() 
                            for dept in PATIENT_DEPT_FLAGS + ['보존', '치주', '원진실'] # 모든 가능한 과
                            if value.get(dept.lower()) is True or value.get(dept.lower()) == 'True' or value.get(dept.lower()) == 'true'
                        ]
                        depts_str = ", ".join(registered_depts) if registered_depts else "미지정"
                        
                        display_text = f"{value.get('환자이름', '이름 없음')} ({pid_key}) [{depts_str}]"
                        patient_list_for_dropdown.append(display_text)
                        patient_key_map[display_text] = pid_key # key가 이제 진료번호
                
                # "전체 선택" 버튼 추가
                if st.button("전체 환자 선택/해제", key="select_all_patients_button"):
                    st.session_state.select_all_mode = not st.session_state.select_all_mode # 상태 토글
                    st.rerun()
                
                # '전체 선택' 모드에 따라 multiselect의 기본값 설정
                default_selection = patient_list_for_dropdown if st.session_state.select_all_mode else []
                
                if not st.session_state.delete_patient_confirm:
                    patients_to_delete_multiselect = st.multiselect(
                        "삭제할 환자 선택",
                        patient_list_for_dropdown,
                        default=default_selection, # 기본값 설정
                        key="delete_patient_multiselect"
                    )
                
                    if st.button("선택한 환자 삭제", key="delete_patient_button"):
                        if patients_to_delete_multiselect:
                            st.session_state.delete_patient_confirm = True
                            st.session_state.patients_to_delete = patients_to_delete_multiselect
                            st.session_state.select_all_mode = False # 삭제 버튼 클릭 시 전체 선택 모드 초기화
                            st.rerun()
                        else:
                            st.warning("삭제할 환자를 선택해주세요.")
                else:
                    st.warning("정말로 선택한 환자를 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.")
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button("예, 삭제합니다", key="confirm_delete_patient"):
                            with st.spinner('삭제 중...'):
                                for patient_to_del_str in st.session_state.patients_to_delete:
                                    patient_key_to_del = patient_key_map.get(patient_to_del_str) # patient_key_to_del은 이제 진료번호
                                    if patient_key_to_del:
                                        patients_ref_for_user.child(patient_key_to_del).delete()
                                
                                st.success(f"선택한 환자 {st.session_state.patients_to_delete} 삭제 완료.")
                                st.session_state.delete_patient_confirm = False
                                st.session_state.patients_to_delete = []
                                st.rerun()
                    with col2:
                        if st.button("아니오, 취소합니다", key="cancel_delete_patient"):
                            st.session_state.delete_patient_confirm = False
                            st.session_state.patients_to_delete = []
                            st.rerun()
                
                st.markdown("---")
        
                with st.form("register_form"):
                    name = st.text_input("환자명")
                    pid = st.text_input("진료번호")
                    
                    # --- [핵심 변경: 다중 선택으로 변경] ---
                    selected_departments = st.multiselect("등록할 진료과 (복수 선택 가능)", DEPARTMENTS_FOR_REGISTRATION)
                    submitted = st.form_submit_button("등록")
                    
                    if submitted:
                        if not name or not pid or not selected_departments:
                            st.warning("환자명, 진료번호, 등록할 진료과를 모두 입력/선택해주세요.")
                        else:
                            pid_key = pid.strip()
                            
                            # 기존 데이터 불러오기 (없으면 새로 생성)
                            new_patient_data = existing_patient_data.get(pid_key, {
                                "환자이름": name,
                                "진료번호": pid # 키로 사용되지만 데이터 내부에도 포함
                            })
                            
                            # 기존에 저장된 모든 진료과 플래그를 False로 초기화 (선택되지 않은 과)
                            for dept_flag in PATIENT_DEPT_FLAGS + ['보존', '치주', '원진실']:
                                lower_dept = dept_flag.lower()
                                new_patient_data[lower_dept] = False

                            # 선택된 진료과만 True로 설정
                            for dept in selected_departments:
                                dept_key_lower = dept.lower()
                                if dept_key_lower in new_patient_data:
                                    new_patient_data[dept_key_lower] = True
                                
                            # 진료번호를 키로 사용하여 데이터 저장 (덮어쓰기/업데이트)
                            patients_ref_for_user.child(pid_key).set(new_patient_data) # <--- **핵심 변경점**
                            st.success(f"{name} ({pid}) [{', '.join(selected_departments)}] 환자 등록/업데이트 완료")
                            st.rerun()


            
            with analysis_tab:
                st.header("📈 OCS 분석 결과")
                analysis_results = db.reference("ocs_analysis/latest_result").get()
                latest_file_name = db.reference("ocs_analysis/latest_file_name").get()

                if analysis_results and latest_file_name:
                    st.markdown(f"**<h3 style='text-align: left;'>{latest_file_name} 분석 결과</h3>**", unsafe_allow_html=True)
                    st.markdown("---")
                    
                    if '소치' in analysis_results:
                        st.subheader("소아치과 현황 (단타)")
                        st.info(f"오전: **{analysis_results['소치']['오전']}명**")
                        st.info(f"오후: **{analysis_results['소치']['오후']}명**")
                    else:
                        st.warning("소아치과 데이터가 엑셀 파일에 없습니다.")
                    st.markdown("---")
                    
                    if '보존' in analysis_results:
                        st.subheader("보존과 현황 (단타)")
                        st.info(f"오전: **{analysis_results['보존']['오전']}명**")
                        st.info(f"오후: **{analysis_results['보존']['오후']}명**")
                    else:
                        st.warning("보존과 데이터가 엑셀 파일에 없습니다.")
                    st.markdown("---")

                    if '교정' in analysis_results:
                        st.subheader("교정과 현황 (Bonding)")
                        st.info(f"오전: **{analysis_results['교정']['오전']}명**")
                        st.info(f"오후: **{analysis_results['교정']['오후']}명**")
                    else:
                        st.warning("교정과 데이터가 엑셀 파일에 없습니다.")
                    st.markdown("---")
                else:
                    st.info("💡 분석 결과가 없습니다. 관리자가 엑셀 파일을 업로드하면 표시됩니다.")
                    
                
                st.divider()
                st.header("🔑 비밀번호 변경")
                new_password = st.text_input("새 비밀번호를 입력하세요", type="password", key="user_new_password_input")
                confirm_password = st.text_input("새 비밀번호를 다시 입력하세요", type="password", key="user_confirm_password_input")
                
                if st.button("비밀번호 변경", key="user_password_change_btn"):
                    if not new_password or not confirm_password:
                        st.error("새 비밀번호와 확인용 비밀번호를 모두 입력해주세요.")
                    elif new_password != confirm_password:
                        st.error("새 비밀번호가 일치하지 않습니다. 다시 확인해주세요.")
                    else:
                        try:
                            users_ref.child(st.session_state.current_firebase_key).update({"password": new_password})
                            st.success("🎉 비밀번호가 성공적으로 변경되었습니다!")
                        except Exception as e:
                            st.error(f"비밀번호 변경 중 오류가 발생했습니다: {e}")
