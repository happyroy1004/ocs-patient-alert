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
import openpyxl 
import datetime 

# Google Calendar API 관련 라이브러리 추가
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import base64

def is_daily_schedule(file_name):
    """
    파일명이 'ocs_MMDD.xlsx' 또는 'ocs_MMDD.xlsm' 형식인지 확인합니다.
    """
    pattern = r'^ocs_\\d{4}\\.(?:xlsx|xlsm)$'
    return re.match(pattern, file_name, re.IGNORECASE) is not None
    
def is_valid_email(email):
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$"
    return re.match(email_regex, email) is not None

def sanitize_path(email):
    return email.replace(".", "_dot_").replace("@", "_at_")

def recover_email(safe_id: str) -> str:
    email = safe_id.replace("_at_", "@").replace("_dot_", ".").replace("_com", ".com")
    return email

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

def is_encrypted_excel(file):
    try:
        file.seek(0)
        return msoffcrypto.OfficeFile(file).is_encrypted()
    except Exception:
        return False

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
        
# --- 세션 상태 초기화 ---
if "clear" in st.query_params and st.query_params["clear"] == "true":
    st.session_state.clear()
    st.query_params["clear"] = "false"
    st.rerun()

if 'auth_status' not in st.session_state:
    st.session_state.auth_status = "unauthenticated"
if 'current_user_email' not in st.session_state:
    st.session_state.current_user_email = ""
if 'current_firebase_key' not in st.session_state:
    st.session_state.current_firebase_key = ""
if 'email_change_mode' not in st.session_state:
    st.session_state.email_change_mode = False
if 'last_email_change_time' not in st.session_state:
    st.session_state.last_email_change_time = 0
if 'email_change_sent' not in st.session_state:
    st.session_state.email_change_sent = False
if 'user_role' not in st.session_state:
    st.session_state.user_role = 'user'
if 'google_creds' not in st.session_state:
    st.session_state['google_creds'] = {}

# OCS 분석 관련 함수 추가
def process_excel_file_and_style(file_io):
    try:
        raw_df = pd.read_excel(file_io)
        excel_data_dfs = pd.read_excel(file_io, sheet_name=None)
        return excel_data_dfs, raw_df.to_excel(index=False, header=True, engine='xlsxwriter')
    except Exception as e:
        st.error(f"엑셀 데이터 처리 및 스타일링 중 오류 발생: {e}")
        return None, None

def run_analysis(df_dict, professors_dict):
    analysis_results = {}
    sheet_department_map = {
        '소치': '소치', '소아치과': '소치', '소아 치과': '소치',
        '보존': '보존', '보존과': '보존', '치과보존과': '보존',
        '교정': '교정', '교정과': '교정', '치과교정과': '교정'
    }
    mapped_dfs = {}
    for sheet_name, df in df_dict.items():
        processed_sheet_name = sheet_name.replace(" ", "").lower()
        for key, dept in sheet_department_map.items():
            if processed_sheet_name == key.replace(" ", "").lower():
                mapped_dfs[dept] = df
                break
    
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

# --- Google Calendar API 관련 함수 (수정) ---
SCOPES = ["https://www.googleapis.com/auth/calendar.events"]

def get_google_calendar_service(user_id_safe):
    creds = st.session_state.get(f"google_creds_{user_id_safe}")
    if not creds:
        creds = load_google_creds_from_firebase(user_id_safe)
        if creds:
            st.session_state[f"google_creds_{user_id_safe}"] = creds

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
    
    if not creds:
        auth_code = st.query_params.get("code")
        if auth_code:
            flow.fetch_token(code=auth_code)
            creds = flow.credentials
            st.session_state[f"google_creds_{user_id_safe}"] = creds
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
        save_google_creds_to_firebase(user_id_safe, creds)

    try:
        service = build('calendar', 'v3', credentials=creds)
        return service
    except HttpError as error:
        st.error(f'Google Calendar 서비스 생성 실패: {error}')
        st.session_state.pop(f"google_creds_{user_id_safe}", None)
        db.reference(f"users/{user_id_safe}/google_creds").delete()
        return None

def create_calendar_event(service, patient_name, pid, department, reservation_datetime, doctor_name, treatment_details):
    seoul_tz = datetime.timezone(datetime.timedelta(hours=9))
    event_start = reservation_datetime.replace(tzinfo=seoul_tz)
    event_end = event_start + datetime.timedelta(minutes=30)
    summary_text = f"환자 내원: {patient_name} ({pid}) / {department} / {doctor_name}"
    description_text = f"진료내역: {treatment_details}"

    event = {
        'summary': summary_text,
        'location': '연세대학교 치과병원',
        'description': description_text,
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
        st.success(f"이벤트 생성 완료: {event.get('htmlLink')}")
    except HttpError as error:
        st.error(f'이벤트 생성 실패: {error}')


#2. User Authentication
def get_user_data(email, password):
    safe_email = sanitize_path(email)
    users_ref = db.reference(f"users/{safe_email}")
    user_data = users_ref.get()
    
    if not user_data:
        return None, None
    
    if user_data.get("password") == password:
        return user_data, safe_email
    else:
        return None, None

def login():
    st.title("로그인")
    with st.form("login_form"):
        email = st.text_input("이메일", key="login_email")
        password = st.text_input("비밀번호", type="password", key="login_password")
        submitted = st.form_submit_button("로그인")
        
        if submitted:
            user_data, user_key = get_user_data(email, password)
            if user_data:
                st.session_state.auth_status = "authenticated"
                st.session_state.current_user_email = email
                st.session_state.current_firebase_key = user_key
                st.session_state.user_role = user_data.get("role", "일반 사용자")
                st.rerun()
            else:
                st.error("이메일 또는 비밀번호가 잘못되었습니다.")

def logout():
    if st.button("로그아웃"):
        for key in st.session_state.keys():
            del st.session_state[key]
        st.rerun()
        
# --- 비밀번호 변경 기능 추가 ---
def change_password_section():
    if st.session_state.get("auth_status") == "authenticated":
        st.divider()
        st.header("🔑 비밀번호 변경")
        
        new_password = st.text_input("새 비밀번호를 입력하세요", type="password", key="new_password_input")
        confirm_password = st.text_input("새 비밀번호를 다시 입력하세요", type="password", key="confirm_password_input")
        
        if st.button("비밀번호 변경"):
            if not new_password or not confirm_password:
                st.error("새 비밀번호와 확인용 비밀번호를 모두 입력해주세요.")
            elif new_password != confirm_password:
                st.error("새 비밀번호가 일치하지 않습니다. 다시 확인해주세요.")
            else:
                try:
                    users_ref = db.reference(f"users/{st.session_state.current_firebase_key}")
                    users_ref.update({"password": new_password})
                    st.success("비밀번호가 성공적으로 변경되었습니다.")
                except Exception as e:
                    st.error(f"비밀번호 변경 중 오류가 발생했습니다: {e}")


#3. Main App UI and Logic
if st.session_state.auth_status == "authenticated":
    st.title(f"👋 환영합니다, {st.session_state.current_user_email}님!")
    st.write(f"현재 역할: {st.session_state.user_role}")
    logout()
    
    st.divider()

    # --- 엑셀 파일 업로드 섹션 ---
    st.header("엑셀 파일 업로드")
    uploaded_file = st.file_uploader("OCS 일일 스케줄 파일을 업로드하세요", type=['xlsx', 'xlsm'])
    
    if uploaded_file and uploaded_file.name != st.session_state.get('last_uploaded_file_name'):
        st.session_state.last_uploaded_file_name = uploaded_file.name
        
        file_name = uploaded_file.name
        if not is_daily_schedule(file_name):
            st.error("OCS 일일 스케줄 파일 형식(ocs_MMDD.xlsx)이 아닙니다. 파일명을 확인해주세요.")
        else:
            try:
                # 엑셀 파일 복호화 및 로드
                excel_file, decrypted_file_io = load_excel(uploaded_file, password=st.secrets["excel_password"])
                
                excel_data_dfs = pd.read_excel(excel_file, sheet_name=None)
                
                # 분석 실행
                professors = st.secrets["professors"]
                analysis_results = run_analysis(excel_data_dfs, professors)
                
                st.session_state.analysis_results = analysis_results
                st.success(f"파일 '{file_name}' 분석 완료!")
                st.session_state.uploaded_file = uploaded_file
            
            except ValueError as ve:
                st.error(f"파일 처리 오류: {ve}")
            except Exception as e:
                st.error(f"예상치 못한 오류 발생: {e}")


    st.divider()

    # --- 탭을 이용한 분리된 기능 섹션 ---
    tab1, tab2, tab3 = st.tabs(["레지던트용 기능", "학생용 기능", "Google Calendar 연동"])

    with tab1:
        st.header("레지던트용 기능")
        st.write("레지던트용 기능이 여기에 표시됩니다.")
        
        if 'analysis_results' in st.session_state and st.session_state.analysis_results:
            st.subheader("OCS 분석 결과")
            for dept, results in st.session_state.analysis_results.items():
                st.markdown(f"**{dept}**")
                st.write(f" - 오전 환자 수: {results['오전']}명")
                st.write(f" - 오후 환자 수: {results['오후']}명")
        else:
            st.warning("파일을 먼저 업로드하고 분석을 실행해주세요.")

    with tab2:
        st.header("학생용 기능")
        st.write("학생용 기능이 여기에 표시됩니다.")

        if st.session_state.get("uploaded_file"):
            st.info(f"업로드된 파일: {st.session_state.uploaded_file.name}")
            
            # --- 학생용 데이터 추출 ---
            try:
                excel_file, decrypted_file_io = load_excel(st.session_state.uploaded_file, password=st.secrets["excel_password"])
                df_dict = pd.read_excel(excel_file, sheet_name=None)
                
                # '환자명', '진료번호', '등록과' 컬럼 추출
                student_df = pd.DataFrame()
                for sheet_name, df in df_dict.items():
                    if all(col in df.columns for col in ['환자명', '진료번호', '등록과', '예약의사']):
                        df_filtered = df[['환자명', '진료번호', '등록과', '예약의사']].copy()
                        student_df = pd.concat([student_df, df_filtered], ignore_index=True)
                
                if not student_df.empty:
                    student_df = student_df.drop_duplicates(subset=['환자명', '진료번호', '등록과']).reset_index(drop=True)
                    st.subheader("💡 학생용 데이터 미리보기")
                    st.dataframe(student_df)

                    # --- 등록 환자 내원 알림 (이메일) ---
                    st.subheader("📧 등록 환자 내원 알림 (이메일)")
                    
                    user_patients_ref = db.reference(f"users/{st.session_state.current_firebase_key}/patients")
                    existing_patients = user_patients_ref.get() or {}
                    
                    user_pids = {p['진료번호'] for p in existing_patients.values()}
                    df_to_send = student_df[student_df['진료번호'].isin(user_pids)]
                    
                    if not df_to_send.empty:
                        st.dataframe(df_to_send)
                        if st.button("선택된 환자에게 이메일 알림 보내기"):
                            sender_email = st.secrets["email"]["sender_email"]
                            sender_password = st.secrets["email"]["sender_password"]
                            receiver_email = st.session_state.current_user_email
                            
                            send_result = send_email(receiver_email, df_to_send, sender_email, sender_password)
                            if send_result is True:
                                st.success("알림 이메일 전송 완료!")
                            else:
                                st.error(f"알림 이메일 전송 실패: {send_result}")
                    else:
                        st.info("오늘 내원하는 등록된 환자가 없습니다.")
                        
            except ValueError as ve:
                st.error(f"파일 처리 오류: {ve}")
            except Exception as e:
                st.error(f"예상치 못한 오류 발생: {e}")
        else:
            st.warning("파일을 먼저 업로드해주세요.")
    
    with tab3:
        st.header("Google Calendar 연동")
        user_id_safe = sanitize_path(st.session_state.current_user_email)
        service = get_google_calendar_service(user_id_safe)
        
        if service:
            st.success("Google Calendar 연동 준비 완료!")
            if st.session_state.get("uploaded_file"):
                try:
                    excel_file, decrypted_file_io = load_excel(st.session_state.uploaded_file, password=st.secrets["excel_password"])
                    df_dict = pd.read_excel(excel_file, sheet_name=None)
                    
                    patient_list = []
                    for sheet_name, df in df_dict.items():
                        if all(col in df.columns for col in ['환자명', '진료번호', '등록과', '예약일자', '예약시간', '예약의사', '진료내역']):
                            for index, row in df.iterrows():
                                if pd.notna(row['예약일자']) and pd.notna(row['예약시간']):
                                    reservation_date_str = str(row['예약일자']).split(' ')[0]
                                    reservation_time_str = str(row['예약시간']).split(' ')[-1]
                                    
                                    try:
                                        reservation_datetime_obj = datetime.datetime.strptime(f"{reservation_date_str} {reservation_time_str}", "%Y-%m-%d %H:%M:%S")
                                        patient_list.append({
                                            '환자명': row['환자명'],
                                            '진료번호': row['진료번호'],
                                            '등록과': row['등록과'],
                                            '예약일자': row['예약일자'],
                                            '예약시간': row['예약시간'],
                                            '예약의사': row['예약의사'],
                                            '진료내역': row['진료내역'],
                                            'datetime_obj': reservation_datetime_obj
                                        })
                                    except ValueError as ve:
                                        st.warning(f"날짜/시간 변환 오류 발생: {ve} - 데이터 건너뛰기")
                                        continue

                    if patient_list:
                        df_patient_list = pd.DataFrame(patient_list)
                        st.subheader("💡 캘린더에 등록할 환자 목록")
                        st.dataframe(df_patient_list[['환자명', '진료번호', '등록과', '예약일자', '예약시간']])
                        
                        if st.button("캘린더에 이벤트 등록"):
                            with st.spinner('이벤트를 캘린더에 등록하는 중...'):
                                for index, row in df_patient_list.iterrows():
                                    create_calendar_event(
                                        service,
                                        patient_name=row['환자명'],
                                        pid=row['진료번호'],
                                        department=row['등록과'],
                                        reservation_datetime=row['datetime_obj'],
                                        doctor_name=row['예약의사'],
                                        treatment_details=row['진료내역']
                                    )
                                time.sleep(2) # 이벤트 등록 시간 확보
                            st.success("모든 이벤트가 성공적으로 등록되었습니다.")
                            
                except ValueError as ve:
                    st.error(f"파일 처리 오류: {ve}")
                except Exception as e:
                    st.error(f"예상치 못한 오류 발생: {e}")
            else:
                st.warning("파일을 먼저 업로드해주세요.")
    
    st.divider()
    
    # --- 환자 등록 및 관리 기능 ---
    st.header("🏥 내 환자 관리")
    
    with st.expander("➕ 새 환자 등록", expanded=False):
        name = st.text_input("환자명", key="add_name")
        pid = st.text_input("진료번호", key="add_pid")
        selected_department = st.selectbox("등록과", ["외과", "내과", "소아과", "신경과"], key="add_department")

        if st.button("환자 등록"):
            if not name or not pid:
                st.error("환자명과 진료번호를 모두 입력해주세요.")
            else:
                patients_ref_for_user = db.reference(f"users/{st.session_state.current_firebase_key}/patients")
                existing_patient_data = patients_ref_for_user.get() or {}

                is_duplicate = False
                for v in existing_patient_data.values():
                    if (v.get("환자명") == name and 
                        v.get("진료번호") == pid and 
                        v.get("등록과") == selected_department):
                        is_duplicate = True
                        break
                
                if is_duplicate:
                    st.error("이미 등록된 환자입니다.")
                else:
                    patients_ref_for_user.push().set({"환자명": name, "진료번호": pid, "등록과": selected_department})
                    st.success(f"{name} ({pid}) [{selected_department}] 환자 등록 완료")
                    st.rerun()
    
    st.subheader("📋 등록된 환자 목록")
    patients_ref_for_user = db.reference(f"users/{st.session_state.current_firebase_key}/patients")
    existing_patient_data = patients_ref_for_user.get()

    if existing_patient_data:
        patient_list = []
        for key, value in existing_patient_data.items():
            value['key'] = key
            patient_list.append(value)
        
        cols = st.columns([1, 1, 1, 0.2])
        cols[0].write("**환자명**")
        cols[1].write("**진료번호**")
        cols[2].write("**등록과**")
        cols[3].write("")

        for patient in patient_list:
            cols = st.columns([1, 1, 1, 0.2])
            cols[0].write(patient["환자명"])
            cols[1].write(patient["진료번호"])
            cols[2].write(patient["등록과"])
            
            if cols[3].button("❌", key=f"delete_{patient['key']}"):
                patients_ref_for_user.child(patient['key']).delete()
                st.success("환자 정보가 삭제되었습니다.")
                st.rerun()
    else:
        st.info("등록된 환자가 없습니다.")

    # 비밀번호 변경 기능 호출
    change_password_section()

#4. App Entry Point
if st.session_state.auth_status == "unauthenticated":
    st.info("로그인이 필요합니다.")
    login()
