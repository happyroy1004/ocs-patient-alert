# ====================================================================================================
# Streamlit & Firebase 기반 환자 관리 시스템
# 
# 이 스크립트는 Streamlit 웹 애플리케이션을 구현하여 병원 환자 데이터를 관리하고,
# Google Calendar와 연동하여 일정을 자동으로 추가하며, 이메일 알림을 보냅니다.
#
# 주요 기능:
# 1.  Firebase Realtime Database 연동: 환자 정보를 CRUD(생성, 읽기, 업데이트, 삭제)합니다.
# 2.  Google Calendar API 연동: 환자 등록 시 자동으로 캘린더 일정을 생성합니다.
# 3.  이메일 알림 기능: 엑셀 파일 업로드 후 환자 정보를 기반으로 이메일을 보냅니다.
# 4.  엑셀 파일 처리: 암호화된 엑셀 파일을 복호화하고, 시트별로 데이터를 가공하여 재가공된
#     엑셀 파일로 출력합니다.
#
# 이 코드는 Streamlit secrets 관리, Firebase 설정, Google Calendar API 키 설정 등
# 외부 서비스 연동을 위한 환경 설정이 필요합니다.
# secrets.toml 파일에 아래와 같은 형식으로 정보를 저장해야 합니다.
# [firebase]
# FIREBASE_SERVICE_ACCOUNT_JSON = "..."
# database_url = "https://your-database-name.firebaseio.com"
#
# [google_calendar]
# client_id = "..."
# client_secret = "..."
# redirect_uri = "http://localhost:8501" # Streamlit 앱 URL
#
# [email]
# user = "your_email@gmail.com"
# password = "your_app_password" # 앱 비밀번호 사용 권장
#
# ====================================================================================================

# --- 라이브러리 임포트 ---
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
from openpyxl.styles import Font, Alignment
import re
import json
import os
import time

# --- Google Calendar API 관련 라이브러리 추가 ---
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import datetime
import base64

# --- 전역 변수 및 상수 설정 ---
# Google Calendar API 인증을 위한 스코프 정의.
# 이 스코프는 캘린더의 이벤트를 생성, 수정, 삭제하는 권한을 포함합니다.
SCOPES = ["https://www.googleapis.com/auth/calendar.events"]

# 엑셀 시트 이름 키워드와 실제 진료과를 매핑하는 딕셔너리
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

# 진료과별 교수님 명단 딕셔너리
professors_dict = {
    '소치': ['김현태', '장기택', '김정욱', '현홍근', '김영재', '신터전', '송지수'],
    '보존': ['이인복', '금기연', '이우철', '유연지', '서덕규', '이창하', '김선영', '손원준'],
    '외과': ['최진영', '서병무', '명훈', '김성민', '박주영', '양훈주', '한정준', '권익재'],
    '치주': ['구영', '이용무', '설양조', '구기태', '김성태', '조영단'],
    '보철': ['곽재영', '김성균', '임영준', '김명주', '권호범', '여인성', '윤형인', '박지만', '이재현', '조준호'],
    '교정': [], '내과': [], '원내생': [], '원스톱': [], '임플란트': [], '병리': []
}

# --- 유틸리티 함수 모음 ---

def is_valid_email(email):
    """
    이메일 주소의 형식을 정규 표현식을 사용하여 검증합니다.
    
    Args:
        email (str): 검증할 이메일 주소 문자열.
        
    Returns:
        bool: 이메일 형식이 유효하면 True, 그렇지 않으면 False.
    """
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(email_regex, email) is not None

def sanitize_path(email):
    """
    Firebase Realtime Database의 키로 사용할 수 있도록 이메일 주소의 특정 문자를 치환합니다.
    Firebase 키는 '.', '#', '$', '/', '[', ']'와 같은 문자를 포함할 수 없습니다.
    
    Args:
        email (str): 치환할 이메일 주소.
        
    Returns:
        str: Firebase 키로 안전하게 사용할 수 있는 문자열.
    """
    return email.replace(".", "_dot_").replace("@", "_at_")

def recover_email(safe_id: str) -> str:
    """
    Firebase 키로 사용된 안전한 문자열을 원래의 이메일 주소로 복원합니다.
    
    Args:
        safe_id (str): Firebase 키로 치환된 이메일 문자열.
        
    Returns:
        str: 원래의 이메일 주소.
    """
    email = safe_id.replace("_at_", "@").replace("_dot_", ".").replace("_com", ".com")
    return email

# --- 파일 처리 함수 모음 ---

def is_encrypted_excel(file):
    """
    업로드된 파일이 암호화된 엑셀 파일인지 확인합니다.
    
    Args:
        file (UploadedFile): Streamlit에서 업로드된 파일 객체.
        
    Returns:
        bool: 암호화된 파일이면 True, 아니면 False.
    """
    try:
        file.seek(0)
        return msoffcrypto.OfficeFile(file).is_encrypted()
    except Exception:
        return False

def load_excel(file, password=None):
    """
    암호화된 엑셀 파일을 복호화하거나 일반 엑셀 파일을 로드합니다.
    
    Args:
        file (UploadedFile): Streamlit에서 업로드된 파일 객체.
        password (str, optional): 암호화된 파일의 비밀번호. Defaults to None.
        
    Returns:
        tuple: (pandas.ExcelFile 객체, 복호화된 파일 객체).
        
    Raises:
        ValueError: 엑셀 로드 또는 복호화에 실패한 경우.
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

# --- 이메일 전송 함수 모음 ---

def send_email(receiver, rows, sender, password, date_str=None, custom_message=None):
    """
    SMTP를 통해 이메일을 전송하는 함수입니다.
    
    Args:
        receiver (str): 수신자 이메일 주소.
        rows (pd.DataFrame): 이메일 본문에 포함될 환자 데이터.
        sender (str): 발신자 이메일 주소.
        password (str): 발신자 이메일 비밀번호 (앱 비밀번호).
        date_str (str, optional): 이메일 제목에 포함될 날짜. Defaults to None.
        custom_message (str, optional): 맞춤 메시지. Defaults to None.
        
    Returns:
        bool or str: 성공 시 True, 실패 시 에러 메시지 문자열.
    """
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver

        # 이메일 내용 구성
        if custom_message:
            msg['Subject'] = "단체 메일 알림"
            body = custom_message
        else:
            subject_prefix = ""
            if date_str:
                subject_prefix = f"{date_str}일에 내원하는 "
            msg['Subject'] = f"{subject_prefix}등록 환자 내원 알림"
            
            html_table = rows.to_html(index=False, escape=False)
            
            # HTML 테이블 스타일
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
            body = f"""
            <p>안녕하세요. 다음은 내원 예정인 등록 환자 명단입니다.</p>
            <br>
            <div class='table-container'>{style}{html_table}</div>
            <br>
            <p>궁금한 점이 있으시면 언제든지 문의해주세요.</p>
            """
        
        msg.attach(MIMEText(body, 'html'))
        
        # SMTP 서버 연결 및 이메일 전송
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        return str(e)
    
def send_registration_email(patient_name, patient_email):
    """
    환자에게 등록 링크가 포함된 이메일을 보내는 함수.
    실제 이메일 서버 설정이 필요합니다.
    """
    st.info(f"📧 등록되지 않은 {patient_name} 환자에게 {patient_email} 주소로 등록 안내 이메일을 보냈습니다.")
    st.warning("🚨 이메일 전송 로직은 실제 이메일 서버 설정이 필요합니다.")
    # 실제 이메일 전송 로직 예시
    # try:
    #     sender = st.secrets["email"]["user"]
    #     password = st.secrets["email"]["password"]
    #     
    #     msg = MIMEMultipart()
    #     msg['From'] = sender
    #     msg['To'] = patient_email
    #     msg['Subject'] = f"{patient_name}님, 환자 등록을 완료해주세요."
    #     
    #     # HTML 본문
    #     html_body = f"""
    #     <html>
    #     <head></head>
    #     <body>
    #         <p>안녕하세요, <strong>{patient_name}</strong>님.</p>
    #         <p>저희 병원 시스템에 환자 등록을 완료해주시기 바랍니다.</p>
    #         <p>아래 링크를 클릭하여 등록을 진행해주세요.</p>
    #         <br>
    #         <a href="https://your-registration-link.com" style="padding: 10px 20px; background-color: #4CAF50; color: white; text-decoration: none; border-radius: 5px;">등록하기</a>
    #         <br>
    #         <p>감사합니다.</p>
    #     </body>
    #     </html>
    #     """
    #     msg.attach(MIMEText(html_body, 'html'))
    #
    #     server = smtplib.SMTP('smtp.gmail.com', 587)
    #     server.starttls()
    #     server.login(sender, password)
    #     server.send_message(msg)
    #     server.quit()
    #     st.success(f"등록 안내 이메일 전송 성공: {patient_email}")
    # except Exception as e:
    #     st.error(f"이메일 전송 실패: {e}")

# --- Google Calendar API 관련 함수 (수정 및 확장) ---
def get_google_calendar_service(user_id_safe):
    """
    사용자별로 Google Calendar 서비스 객체를 반환하거나, 인증 URL을 표시합니다.
    Streamlit의 세션 상태를 활용하여 인증 정보를 관리합니다.
    
    Args:
        user_id_safe (str): Firebase 키로 안전하게 치환된 사용자 ID.
        
    Returns:
        googleapiclient.discovery.Resource or None: Calendar 서비스 객체 또는 None.
    """
    creds = st.session_state.get(f"google_creds_{user_id_safe}")
    
    try:
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
    except KeyError:
        st.error("`secrets.toml` 파일에 Google Calendar API 설정이 누락되었습니다. `[google_calendar]` 섹션을 확인해주세요.")
        return None
        
    # 인증 플로우 생성
    flow = InstalledAppFlow.from_client_config(client_config, SCOPES, redirect_uri=st.secrets["google_calendar"]["redirect_uri"])
    
    if not creds:
        auth_code = st.query_params.get("code")
        
        if auth_code:
            # 인증 코드를 사용하여 토큰을 교환
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            st.session_state[f"google_creds_{user_id_safe}"] = creds
            st.success("Google Calendar 인증이 완료되었습니다. 잠시 후 페이지가 새로고침됩니다.")
            st.query_params.clear()
            st.rerun()
        else:
            auth_url, _ = flow.authorization_url(prompt='consent')
            st.warning("Google Calendar 연동을 위해 인증이 필요합니다. 아래 링크를 클릭하여 권한을 부여하세요.")
            st.markdown(f"**[Google Calendar 인증 링크]({auth_url})**")
            return None

    if creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            st.session_state[f"google_creds_{user_id_safe}"] = creds
        except Exception as e:
            st.error(f"Google Calendar 토큰 갱신 실패: {e}")
            st.session_state.pop(f"google_creds_{user_id_safe}", None)
            return None

    try:
        service = build('calendar', 'v3', credentials=creds)
        return service
    except HttpError as error:
        st.error(f'Google Calendar 서비스 생성 실패: {error}')
        st.session_state.pop(f"google_creds_{user_id_safe}", None)
        return None

def create_calendar_event(service, patient_name, pid, department):
    """
    Google Calendar에 환자 등록 이벤트를 생성합니다.
    
    Args:
        service (googleapiclient.discovery.Resource): Google Calendar API 서비스 객체.
        patient_name (str): 환자 이름.
        pid (str): 진료번호.
        department (str): 등록 과.
    """
    # 이벤트 시작 및 종료 시간 설정 (현재 시간부터 1시간 후)
    now = datetime.datetime.now(datetime.timezone.utc).astimezone(datetime.timezone(datetime.timedelta(hours=9)))
    event_start_time = now.isoformat()
    event_end_time = (now + datetime.timedelta(hours=1)).isoformat()
    
    event = {
        'summary': f'환자 내원: {patient_name} ({department})',
        'location': f'진료번호: {pid}',
        'description': f'환자명: {patient_name}\n진료번호: {pid}\n등록 과: {department}',
        'start': {
            'dateTime': event_start_time,
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': event_end_time,
            'timeZone': 'Asia/Seoul',
        },
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': 24 * 60}, # 24시간 전 이메일 알림
                {'method': 'popup', 'minutes': 10},      # 10분 전 팝업 알림
            ],
        },
    }
    
    try:
        event = service.events().insert(calendarId='primary', body=event).execute()
        st.success(f"'{patient_name}' 환자 등록 일정이 캘린더에 성공적으로 추가되었습니다.")
    except HttpError as error:
        st.error(f"캘린더 이벤트 생성 중 오류 발생: {error}")
        st.warning("Google Calendar 인증 권한을 다시 확인해주세요.")
    except Exception as e:
        st.error(f"알 수 없는 오류 발생: {e}")

# --- 엑셀 처리 관련 함수 모음 ---
def process_sheet_v8(df, professors_list, sheet_key):
    """
    엑셀 시트 데이터를 진료과별로 분류하고 정렬하여 처리합니다.
    
    Args:
        df (pd.DataFrame): 처리할 데이터가 담긴 DataFrame.
        professors_list (list): 해당 진료과의 교수님 이름 목록.
        sheet_key (str): 진료과 키 (예: '보철', '교정').
        
    Returns:
        pd.DataFrame: 처리 및 정렬이 완료된 DataFrame.
        
    Raises:
        st.error: '예약의사' 또는 '예약시간' 컬럼이 없는 경우.
    """
    df = df.drop(columns=['예약일시'], errors='ignore')
    if '예약의사' not in df.columns or '예약시간' not in df.columns:
        st.error(f"시트 처리 오류: '예약의사' 또는 '예약시간' 컬럼이 DataFrame에 없습니다.")
        return pd.DataFrame(columns=['진료번호', '예약시간', '환자명', '예약의사', '진료내역'])

    # 예약의사 및 예약시간으로 정렬
    df = df.sort_values(by=['예약의사', '예약시간'])
    
    # 교수님과 비-교수님(전공의 등)으로 데이터 분리
    professors = df[df['예약의사'].isin(professors_list)]
    non_professors = df[~df['예약의사'].isin(professors_list)]

    # 보철과와 다른 과의 정렬 기준이 다름
    if sheet_key != '보철':
        non_professors = non_professors.sort_values(by=['예약시간', '예약의사'])
    else:
        non_professors = non_professors.sort_values(by=['예약의사', '예약시간'])

    final_rows = []
    current_time = None
    current_doctor = None

    # 비-교수님 데이터 처리
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

    # 교수님 데이터 섹션 구분자 추가
    final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
    final_rows.append(pd.Series(["<교수님>"] + [" "] * (len(df.columns) - 1), index=df.columns))

    # 교수님 데이터 처리
    current_professor = None
    for _, row in professors.iterrows():
        if current_professor != row['예약의사']:
            if current_professor is not None:
                final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
            current_professor = row['예약의사']
        final_rows.append(row)

    # 최종 DataFrame 생성 및 컬럼 정리
    final_df = pd.DataFrame(final_rows, columns=df.columns)
    required_cols = ['진료번호', '예약시간', '환자명', '예약의사', '진료내역']
    final_df = final_df[[col for col in required_cols if col in final_df.columns]]
    return final_df

def process_excel_file_and_style(file_bytes_io):
    """
    업로드된 엑셀 파일을 처리하고, 데이터를 가공한 후 스타일을 적용하여
    새로운 엑셀 파일을 생성합니다.
    
    Args:
        file_bytes_io (io.BytesIO): 복호화 또는 로드된 엑셀 파일의 BytesIO 객체.
        
    Returns:
        tuple: (dict, io.BytesIO) - 처리된 DataFrame 딕셔너리, 스타일이 적용된 파일 객체.
        
    Raises:
        ValueError: 엑셀 로드 또는 처리 중 오류가 발생한 경우.
    """
    file_bytes_io.seek(0)
    
    try:
        # data_only=True 옵션으로 셀에 있는 수식이 아닌 결과값만 가져옴
        wb_raw = load_workbook(filename=file_bytes_io, keep_vba=False, data_only=True)
    except Exception as e:
        raise ValueError(f"엑셀 워크북 로드 실패: {e}")

    processed_sheets_dfs = {}

    for sheet_name_raw in wb_raw.sheetnames:
        sheet_name_lower = sheet_name_raw.strip().lower()

        # 시트 이름 키워드를 기반으로 진료과 매핑
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
        
        # 첫 번째 유효한 헤더 행 찾기
        while values and (values[0] is None or all((v is None or str(v).strip() == "") for v in values[0])):
            values.pop(0)
            
        if len(values) < 2:
            st.warning(f"시트 '{sheet_name_raw}'에 유효한 데이터가 충분하지 않습니다. 건너깁니다.")
            continue

        # DataFrame으로 변환 및 헤더 설정
        df = pd.DataFrame(values)
        df.columns = df.iloc[0]
        df = df.drop([0]).reset_index(drop=True)
        df = df.fillna("").astype(str)

        # 예약의사 컬럼 정리
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

    # 처리된 데이터를 새로운 엑셀 파일로 저장
    output_buffer_for_styling = io.BytesIO()
    with pd.ExcelWriter(output_buffer_for_styling, engine='openpyxl') as writer:
        for sheet_name_raw, df in processed_sheets_dfs.items():
            df.to_excel(writer, sheet_name=sheet_name_raw, index=False)
            
    # 스타일 적용
    output_buffer_for_styling.seek(0)
    wb_styled = load_workbook(output_buffer_for_styling, keep_vba=False, data_only=True)
    
    for sheet_name in wb_styled.sheetnames:
        ws = wb_styled[sheet_name]
        header = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}
        
        # 헤더 폰트 스타일 적용
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
            if row[0].value == "<교수님>":
                for cell in row:
                    if cell.value:
                        cell.font = Font(bold=True)
            
            # 특정 조건에 따라 폰트 스타일 적용
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
st.set_page_config(layout="wide", page_title="환자 관리 시스템")

# 제목에 링크 추가 및 초기화 로직
st.markdown("""
<style>
    .title-link {
        text-decoration: none;
        color: inherit;
    }
</style>
<h1>
    <a href="." class="title-link">👨‍⚕️ 환자 내원 확인 시스템</a>
</h1>
""", unsafe_allow_html=True)
st.markdown("---")
st.markdown("<p style='text-align: left; color: grey; font-size: small;'>directed by HSY</p>", unsafe_allow_html=True)

# --- Firebase 초기화 ---
# 이 블록은 앱이 처음 로드될 때 한 번만 실행됩니다.
if not firebase_admin._apps:
    try:
        firebase_credentials_json_str = st.secrets["firebase"]["FIREBASE_SERVICE_ACCOUNT_JSON"]
        firebase_credentials_dict = json.loads(firebase_credentials_json_str)

        cred = credentials.Certificate(firebase_credentials_dict)
        firebase_admin.initialize_app(cred, {
            'databaseURL': st.secrets["firebase"]["database_url"]
        })
        st.success("Firebase 초기화 성공!")
    except KeyError:
        st.error("`secrets.toml` 파일에 Firebase 설정이 누락되었습니다. `[firebase]` 섹션을 확인해주세요.")
        st.stop()
    except Exception as e:
        st.error(f"Firebase 초기화 오류: {e}")
        st.info("secrets.toml 파일의 Firebase 설정(FIREBASE_SERVICE_ACCOUNT_JSON 또는 database_url)을 [firebase] 섹션 아래에 올바르게 작성했는지 확인해주세요.")
        st.stop()
        
# --- 세션 상태 초기화 ---
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
if 'excel_data_to_send' not in st.session_state:
    st.session_state.excel_data_to_send = None
if 'processed_excel_dfs' not in st.session_state:
    st.session_state.processed_excel_dfs = {}

users_ref = db.reference("users")

# --- Streamlit 사이드바 ---
with st.sidebar:
    st.header("설정 및 도구")
    
    # 사용 설명서 다운로드
    pdf_file_path = "manual.pdf"
    pdf_display_name = "사용 설명서"
    if os.path.exists(pdf_file_path):
        with open(pdf_file_path, "rb") as pdf_file:
            st.download_button(
                label=f"⬇️ {pdf_display_name} 다운로드",
                data=pdf_file,
                file_name=pdf_file_path,
                mime="application/pdf"
            )
    else:
        st.warning(f"⚠️ {pdf_display_name} 파일을 찾을 수 없습니다.")

    # Firebase 데이터 초기화 버튼 (주의: 모든 데이터 삭제)
    st.markdown("---")
    st.markdown("#### 데이터 관리")
    if st.button("🚨 모든 환자 데이터 삭제 (초기화)"):
        if st.session_state.logged_in_as_admin:
            users_ref.child(st.session_state.current_firebase_key).child("patients").delete()
            st.success("모든 환자 데이터가 삭제되었습니다.")
            st.rerun()
        else:
            st.error("관리자 계정으로 로그인해야 이 기능을 사용할 수 있습니다.")

# --- 사용자 로그인 섹션 ---
st.markdown("### 사용자 로그인")
user_name = st.text_input("사용자 이름을 입력하세요 (예시: 홍길동)", value=st.session_state.user_id_input_value)
is_admin_input = (user_name.strip().lower() == "admin")

# user_name이 입력되었을 때 기존 사용자 검색
if user_name and not is_admin_input and not st.session_state.email_change_mode:
    all_users_meta = users_ref.get()
    matched_users_by_name = []
    if all_users_meta:
        for safe_key, user_info in all_users_meta.items():
            if user_info and user_info.get('name') == user_name:
                matched_users_by_name.append((safe_key, user_info))
    
    if len(matched_users_by_name) == 1:
        safe_key, user_info = matched_users_by_name[0]
        st.session_state.current_user_name = user_name
        st.session_state.current_firebase_key = safe_key
        st.session_state.found_user_email = recover_email(safe_key)
        st.info(f"사용자 '{user_name}'님, 로그인되었습니다.")
    elif len(matched_users_by_name) > 1:
        st.warning(f"'{user_name}' 이름을 가진 사용자가 여러 명 있습니다. 이메일로 로그인해주세요.")
    
    # 이메일로 로그인
    user_email_input = st.text_input("이메일을 입력하세요 (예시: test@example.com)")
    if user_email_input:
        if is_valid_email(user_email_input):
            safe_key = sanitize_path(user_email_input)
            user_info = users_ref.child(safe_key).get()
            if user_info and user_info.get('name') == user_name:
                st.session_state.current_user_name = user_name
                st.session_state.current_firebase_key = safe_key
                st.session_state.found_user_email = user_email_input
                st.success(f"사용자 '{user_name}'님 ({user_email_input}) 로그인되었습니다.")
            else:
                st.error("사용자 이름과 이메일이 일치하지 않습니다.")
        else:
            st.error("유효한 이메일 주소를 입력해주세요.")
            
    # 새 사용자 등록
    if not st.session_state.current_firebase_key and st.button("신규 사용자 등록"):
        st.session_state.email_change_mode = True
        st.session_state.user_id_input_value = user_name
        st.info("신규 사용자 등록 모드입니다. 이메일을 입력하세요.")

if st.session_state.email_change_mode:
    st.markdown("### 신규 사용자 등록")
    new_user_email = st.text_input("등록할 이메일을 입력하세요")
    if st.button("등록 완료"):
        if is_valid_email(new_user_email):
            safe_key = sanitize_path(new_user_email)
            users_ref.child(safe_key).set({
                'name': st.session_state.user_id_input_value,
                'email': new_user_email
            })
            st.success("새 사용자가 등록되었습니다. 로그인해주세요.")
            st.session_state.email_change_mode = False
            st.session_state.user_id_input_value = ""
            st.rerun()
        else:
            st.error("유효한 이메일 주소를 입력하세요.")
            
# --- 어드민 로그인 섹션 ---
if is_admin_input and not st.session_state.logged_in_as_admin:
    st.markdown("### 관리자 로그인")
    admin_password = st.text_input("관리자 비밀번호", type="password")
    if st.button("로그인"):
        # secrets.toml에 있는 실제 비밀번호와 비교
        if admin_password == st.secrets["admin"]["password"]:
            st.session_state.logged_in_as_admin = True
            st.session_state.admin_password_correct = True
            st.session_state.current_user_name = "관리자"
            st.success("관리자 계정으로 로그인되었습니다.")
        else:
            st.error("비밀번호가 틀렸습니다.")

st.markdown("---")

# 로그인 상태에 따라 UI를 다르게 보여줌
if st.session_state.current_user_name:
    st.success(f"현재 로그인: {st.session_state.current_user_name}님")
    
    # 캘린더 인증 섹션
    if not st.session_state.logged_in_as_admin and not st.session_state.google_calendar_auth_needed:
        service = get_google_calendar_service(st.session_state.current_firebase_key)
        if service:
            st.session_state.google_calendar_auth_needed = False
        else:
            st.session_state.google_calendar_auth_needed = True

    # 엑셀 파일 업로드 및 처리 섹션
    st.header("📤 엑셀 파일 업로드")
    uploaded_file = st.file_uploader("암호화된 엑셀 파일을 업로드하세요", type=['xlsx'])
    password_input = None
    
    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        file_io = io.BytesIO(file_bytes)
        
        is_encrypted = is_encrypted_excel(file_io)
        
        if is_encrypted:
            password_input = st.text_input("파일 비밀번호를 입력하세요", type="password")
            if not password_input:
                st.warning("암호화된 파일입니다. 비밀번호를 입력해주세요.")
                st.stop()
            
        if not is_encrypted or password_input:
            try:
                excel_file, decrypted_file_io = load_excel(file_io, password=password_input)
                
                # 파일 처리 및 스타일 적용
                with st.spinner("엑셀 파일 처리 중..."):
                    processed_dfs, styled_file = process_excel_file_and_style(decrypted_file_io)
                    st.session_state.processed_excel_dfs = processed_dfs
                    st.session_state.excel_data_to_send = styled_file
                    
                st.success("엑셀 파일 처리 완료!")
                st.download_button(
                    label="다운로드 (스타일 적용)",
                    data=styled_file,
                    file_name=f"processed_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # 처리된 데이터 미리보기
                with st.expander("처리된 환자 데이터 미리보기"):
                    if st.session_state.processed_excel_dfs:
                        for sheet_name, df in st.session_state.processed_excel_dfs.items():
                            st.markdown(f"#### 시트: {sheet_name}")
                            st.dataframe(df)

            except ValueError as e:
                st.error(f"오류: {e}")
                
    st.markdown("---")

    # 환자 등록 및 캘린더 추가 섹션
    st.header("📝 환자 등록 및 일정 추가")
    
    # Firebase에서 환자 목록 가져오기
    patients_ref_for_user = db.reference(f'users/{st.session_state.current_firebase_key}/patients')
    existing_patient_data = patients_ref_for_user.get()

    if existing_patient_data:
        st.subheader("등록된 환자 목록")
        # 등록된 환자 목록 표시
        with st.expander("목록 보기/삭제"):
            for key, val in existing_patient_data.items():
                with st.container(border=True):
                    info_col, btn_col = st.columns([4, 1])
        
                    with info_col:
                        st.markdown(f"**{val.get('환자명', '이름 없음')}** / {val.get('진료번호', '번호 없음')} / {val.get('등록과', '미지정')}")
        
                    with btn_col:
                        # 삭제 버튼
                        if st.button("X", key=f"delete_button_{key}"):
                            patients_ref_for_user.child(key).delete()
                            st.rerun()
    else:
        st.info("등록된 환자가 없습니다.")
        
    st.subheader("새 환자 등록")
    
    with st.form("register_form"):
        name = st.text_input("환자명")
        pid = st.text_input("진료번호")
        patient_email = st.text_input("환자 이메일 (등록 안내용)")

        # 등록 과 선택 (더미 데이터 사용)
        departments_for_registration = sorted(list(set(sheet_keyword_to_department_map.values())))
        selected_department = st.selectbox("등록 과", departments_for_registration)
    
        submitted = st.form_submit_button("등록 및 일정 추가")
    
        if submitted:
            if not name or not pid or not patient_email:
                st.warning("모든 항목을 입력해주세요.")
            elif not is_valid_email(patient_email):
                st.error("유효한 이메일 주소를 입력해주세요.")
            else:
                is_patient_registered = False
                if existing_patient_data:
                    for key, val in existing_patient_data.items():
                        if val.get('환자명') == name and val.get('진료번호') == pid:
                            is_patient_registered = True
                            break
    
                if is_patient_registered:
                    st.success("✅ 등록된 환자입니다. 캘린더에 일정을 추가합니다.")
                    google_service = get_google_calendar_service(st.session_state.current_firebase_key)
                    if google_service:
                        create_calendar_event(google_service, name, pid, selected_department)
                    else:
                        st.error("Google Calendar 서비스가 초기화되지 않았습니다. 인증을 진행해주세요.")
                else:
                    st.warning("⚠️ 등록되지 않은 환자입니다. 등록 안내 이메일을 보냅니다.")
                    send_registration_email(name, patient_email)
                    
                    # Firebase에 신규 환자 정보 추가
                    patients_ref_for_user.push().set({"환자명": name, "진료번호": pid, "등록과": selected_department, "이메일": patient_email})
                    st.success(f"{name} 환자가 Firebase에 등록되었습니다.")
                    st.rerun()

else:
    st.info("로그인하여 시스템을 사용해주세요.")

st.markdown("---")
st.markdown("### 개발 노트")
st.info("이 앱은 Streamlit, Firebase, Google Calendar API를 활용한 예시입니다. 실제 서비스 환경에 맞게 `secrets.toml` 파일과 각 API 연동 로직을 수정해야 합니다.")
st.markdown("#### 주의 사항")
st.warning("Google Calendar API는 민감한 사용자 데이터를 다루므로, `redirect_uri`를 정확하게 설정하고, 앱을 OAuth 동의 화면에서 '게시' 상태로 전환해야 정상적으로 동작합니다.")
