# #1. Imports, Validation Functions, and Firebase Initialization
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
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$"
    return re.match(email_regex, email) is not None

def sanitize_path(email):
    return email.replace('.', '_').replace('@', '_')

def recover_email(sanitized_path):
    return sanitized_path.replace('_', '.', 1).replace('_', '@', 1)

# Firebase 초기화
if not firebase_admin._apps:
    try:
        firebase_credentials_json_str = st.secrets["firebase"]["FIREBASE_SERVICE_ACCOUNT_JSON"]
        firebase_credentials_dict = json.loads(firebase_credentials_json_str)
        cred = credentials.Certificate(firebase_credentials_dict)
        firebase_admin.initialize_app(cred, {
            'databaseURL': st.secrets["firebase"]["FIREBASE_DATABASE_URL"]
        })
        # print("Firebase App Initialized Successfully")
    except Exception as e:
        st.error(f"Firebase 초기화 실패: {e}")
        st.stop()

# 전역 변수 설정
users_ref = db.reference("users")
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

# --- Google Calendar API 관련 함수 ---
def save_google_creds_to_firebase(uid_safe, creds):
    creds_dict = {
        'token': creds.token,
        'refresh_token': creds.refresh_token,
        'token_uri': creds.token_uri,
        'client_id': creds.client_id,
        'client_secret': creds.client_secret,
        'scopes': creds.scopes
    }
    db.reference(f'google_creds/{uid_safe}').set(creds_dict)

def load_google_creds_from_firebase(uid_safe):
    creds_dict = db.reference(f'google_creds/{uid_safe}').get()
    if creds_dict:
        creds = Credentials(**creds_dict)
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            save_google_creds_to_firebase(uid_safe, creds)
        return creds
    return None
    
def create_calendar_event(service, patient_name, patient_pid, department,
                          reservation_date_str, reservation_time_str, doctor_name, treatment_details):
    try:
        reservation_datetime = datetime.datetime.strptime(f"{reservation_date_str} {reservation_time_str}", "%Y-%m-%d %H:%M")
        start_time = reservation_datetime.isoformat()
        end_time = (reservation_datetime + datetime.timedelta(minutes=30)).isoformat()
    except (ValueError, TypeError):
        # 날짜/시간이 없는 경우 현재 시각으로 대체
        now = datetime.datetime.utcnow()
        start_time = now.isoformat() + 'Z'
        end_time = (now + datetime.timedelta(minutes=30)).isoformat() + 'Z'

    summary = f"({department}) {patient_name}님 내원"
    description = (
        f"진료번호: {patient_pid}\n"
        f"담당의: {doctor_name}\n"
        f"진료내역: {treatment_details}"
    )

    event = {
        'summary': summary,
        'description': description,
        'start': {
            'dateTime': start_time,
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': end_time,
            'timeZone': 'Asia/Seoul',
        },
    }
    try:
        service.events().insert(calendarId='primary', body=event).execute()
    except Exception as e:
        st.error(f"캘린더 일정 추가 실패: {e}")

# --- 이메일 전송 함수 ---
def send_email(to_email, matched_df, sender, sender_pw, date_str=None, custom_message=None):
    if not is_valid_email(to_email):
        return f"유효하지 않은 이메일 주소: {to_email}"

    try:
        msg = MIMEMultipart("alternative")
        msg['From'] = f"KI-OAS <{sender}>"
        msg['To'] = to_email
        msg['Subject'] = "환자 내원 안내"

        if custom_message:
            html_content = f"<html><body>{custom_message}</body></html>"
            msg.attach(MIMEText(html_content, 'html', 'utf-8'))
        else:
            if matched_df.empty:
                return "매칭된 환자 데이터가 없습니다."
            
            email_date_str = date_str if date_str else datetime.date.today().strftime("%Y년 %m월 %d일")
            table_html = matched_df[['예약시간', '환자명', '예약의사', '진료내역']].to_html(index=False)
            
            html_content = f"""
            <html>
                <head>
                    <style>
                        table {{ border-collapse: collapse; width: 100%; }}
                        th, td {{ border: 1px solid #dddddd; text-align: left; padding: 8px; }}
                        th {{ background-color: #f2f2f2; }}
                    </style>
                </head>
                <body>
                    <h4>{email_date_str}의 내원 환자 정보입니다.</h4>
                    {table_html}
                </body>
            </html>
            """
            msg.attach(MIMEText(html_content, 'html', 'utf-8'))

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender, sender_pw)
            smtp.sendmail(sender, to_email, msg.as_string())
        return True
    except smtplib.SMTPAuthenticationError:
        return "SMTP 인증 실패. Gmail 앱 비밀번호를 확인해주세요."
    except Exception as e:
        return f"이메일 전송 오류: {e}"

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

# #2. User Login and Logout
st.title("👨‍⚕️ KI-OAS 환자 내원 확인 시스템")

if not st.session_state.logged_in:
    st.header("로그인")
    login_email = st.text_input("이메일 주소", key="login_email")

    if st.button("로그인"):
        if is_valid_email(login_email):
            user_id_safe = sanitize_path(login_email)
            user_info = users_ref.child(user_id_safe).get()
            
            if user_info and 'name' in user_info:
                st.session_state.logged_in = True
                st.session_state.current_user_name = user_info['name']
                st.session_state.found_user_email = login_email
                st.success(f"로그인 성공: {st.session_state.current_user_name}님 환영합니다!")
                st.rerun()
            elif login_email == st.secrets["admin"]["email"]:
                st.session_state.logged_in = True
                st.session_state.current_user_name = "admin"
                st.session_state.found_user_email = login_email
                st.success(f"로그인 성공: 관리자님 환영합니다!")
                st.rerun()
            else:
                st.error("등록되지 않은 사용자입니다.")
                st.session_state.logged_in = False
        else:
            st.error("유효한 이메일 주소를 입력해주세요.")
else:
    st.sidebar.markdown(f"**로그인 계정:** {st.session_state.current_user_name}")
    if st.sidebar.button("로그아웃"):
        st.session_state.logged_in = False
        st.session_state.current_user_name = None
        st.session_state.found_user_email = None
        st.session_state.admin_password_correct = False
        st.session_state.processed_excel_data_dfs = None
        st.session_state.processed_styled_bytes = None
        st.session_state.google_calendar_service = None
        st.success("로그아웃되었습니다.")
        st.rerun()

# #3. User Registration (visible only if not logged in)
if not st.session_state.logged_in:
    st.header("회원가입")
    with st.form("registration_form"):
        st.write("새로운 사용자를 등록합니다.")
        new_name = st.text_input("이름", key="new_name")
        new_email = st.text_input("이메일 주소", key="new_email")
        submitted = st.form_submit_button("회원가입")

        if submitted:
            if not new_name or not new_email:
                st.warning("이름과 이메일을 모두 입력해주세요.")
            elif not is_valid_email(new_email):
                st.error("유효한 이메일 주소를 입력해주세요.")
            else:
                user_id_safe = sanitize_path(new_email)
                if users_ref.child(user_id_safe).get():
                    st.error("이미 등록된 이메일 주소입니다.")
                else:
                    users_ref.child(user_id_safe).set({"name": new_name, "email": new_email})
                    st.success("회원가입이 완료되었습니다. 로그인해주세요.")

# #4. Excel Processing Constants and Functions
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

def load_excel(file, password=None):
    """암호화된 엑셀 파일을 로드합니다."""
    file.seek(0)
    try:
        if password:
            decrypted_file = io.BytesIO()
            office_file = msoffcrypto.OfficeFile(file)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted_file)
            decrypted_file.seek(0)
            return pd.ExcelFile(decrypted_file), decrypted_file
        else:
            return pd.ExcelFile(file), file
    except msoffcrypto.exceptions.InvalidKeyError:
        raise ValueError("잘못된 비밀번호입니다.")
    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        st.stop()
        
def is_encrypted_excel(file):
    """파일이 암호화되어 있는지 확인합니다."""
    file.seek(0)
    try:
        office_file = msoffcrypto.OfficeFile(file)
        return True
    except Exception:
        return False
    finally:
        file.seek(0)

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
                df_non_prof['예약시간'] = pd.to_datetime(df_non_prof['예약시간'], format='%H:%M').dt.time
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
                df_non_prof['예약시간'] = pd.to_datetime(df_non_prof['예약시간'], format='%H:%M').dt.time
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
                df_bonding['예약시간'] = pd.to_datetime(df_bonding['예약시간'], format='%H:%M').dt.time
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

            departments_for_registration = sorted(list(set(sheet_keyword_to_department_map.values())))\
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

# #6. Oauth2 Callback Functionality
query_params = st.query_params
if 'code' in query_params:
    try:
        user_id_safe = sanitize_path(st.session_state.found_user_email)
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
        flow.fetch_token(code=query_params['code'])
        creds = flow.credentials
        save_google_creds_to_firebase(user_id_safe, creds)
        st.success("Google Calendar 연동이 성공적으로 완료되었습니다!")
        st.session_state.google_calendar_service = build('calendar', 'v3', credentials=creds)
        st.query_params.clear()
        st.rerun()

    except Exception as e:
        st.error(f"Google Calendar 연동 실패: {e}")
        st.query_params.clear()

# #7. Admin Mode Functionality
if st.session_state.logged_in and st.session_state.current_user_name.lower() == "admin":
    st.session_state.logged_in_as_admin = True
    st.session_state.found_user_email = st.secrets["admin"]["email"]
    st.header("관리자 기능")

    # 엑셀 업로드 섹션
    st.subheader("💻 Excel File Processor")
    uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])

    if uploaded_file:
        file_content = uploaded_file.getvalue()
        file_stream = io.BytesIO(file_content)

        password = st.text_input("엑셀 파일 비밀번호 입력", type="password") if is_encrypted_excel(file_stream) else None
        
        file_stream.seek(0)
        
        if is_encrypted_excel(file_stream) and not password:
            st.info("암호화된 파일입니다. 비밀번호를 입력해주세요.")
            st.stop()
        
        try:
            file_name = uploaded_file.name

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

            xl_object, raw_file_io = load_excel(file_stream, password)
            excel_data_dfs, styled_excel_bytes = process_excel_file_and_style(raw_file_io)

            if excel_data_dfs is None or styled_excel_bytes is None:
                st.warning("엑셀 파일 처리 중 문제가 발생했거나 처리할 데이터가 없습니다.")
                st.stop()
            
            filtered_excel_data_dfs = {}
            for sheet_name, df in excel_data_dfs.items():
                department = sheet_keyword_to_department_map.get(sheet_name.strip().lower(), None)
                if department and department in professors_dict:
                    professors_in_dept = professors_dict[department]
                    doctor_col = None
                    for col in ['진료의사', '의사명', '담당의']:
                        if col in df.columns:
                            doctor_col = col
                            break
                    
                    if doctor_col:
                        filtered_df = df[~df[doctor_col].isin(professors_in_dept)]
                        filtered_excel_data_dfs[sheet_name] = filtered_df
                    else:
                        filtered_excel_data_dfs[sheet_name] = df
                else:
                    filtered_excel_data_dfs[sheet_name] = df
            
            st.session_state.processed_excel_data_dfs = filtered_excel_data_dfs
            st.session_state.processed_styled_bytes = styled_excel_bytes

            st.info("기존 OCS 분석 데이터를 삭제하고 새로운 파일로 덮어쓰는 중...")
            processed_data_ref = db.reference("processed_data/ocs_analysis")
            data_to_save = {
                "file_name": file_name,
                "sheets": {sheet_name: df.to_dict('records') for sheet_name, df in filtered_excel_data_dfs.items()}
            }
            processed_data_ref.set(data_to_save)
            st.success("엑셀 분석 데이터가 Firebase에 성공적으로 저장되었습니다.")
            
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
                    for sheet_name_excel_raw, df_sheet in filtered_excel_data_dfs.items():
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

            st.subheader("매칭된 환자 명단")
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
                            result = send_email(real_email, df_matched, sender, sender_pw, date_str=reservation_date_excel)
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

                            creds = load_google_creds_from_firebase(user_safe_key)

                            if creds and creds.valid and not creds.expired:
                                try:
                                    service = build('calendar', 'v3', credentials=creds)
                                    if df_matched is not None and not df_matched.empty:
                                        for _, row in df_matched.iterrows():
                                            doctor_name = row.get('진료의사', '') or row.get('의사명', '') or row.get('담당의', '')
                                            treatment_details = row.get('진료내역', '')
                                            create_calendar_event(service, row['환자명'], row['진료번호'], row.get('시트', ''),
                                                    reservation_date_str=reservation_date_excel, reservation_time_str=row.get('예약시간'), doctor_name=doctor_name, treatment_details=treatment_details)
                                    st.success(f"**{user_name}**님의 캘린더에 일정을 추가했습니다.")
                                except Exception as e:
                                    st.error(f"**{user_name}**님의 캘린더 일정 추가 실패: {e}")
                            else:
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
                data=st.session_state.processed_styled_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except ValueError as ve:
            st.error(f"파일 처리 실패: {ve}")
        except Exception as e:
            st.error(f"예상치 못한 오류 발생: {e}")

    st.markdown("---")
    st.subheader("🛠️ 최고 관리자 권한")
    admin_password_input = st.text_input("관리자 비밀번호를 입력하세요", type="password", key="admin_password")

    try:
        secret_admin_password = st.secrets["admin"]["password"]
    except KeyError:
        secret_admin_password = None
        st.error("⚠️ secrets.toml 파일에 'admin.password' 설정이 없습니다. 개발자에게 문의하세요.")

    if admin_password_input and admin_password_input == secret_admin_password:
        st.session_state.admin_password_correct = True
        st.success("최고 관리자 권한이 활성화되었습니다.")
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
