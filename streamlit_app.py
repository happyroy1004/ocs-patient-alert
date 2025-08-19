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
import hashlib # 비밀번호 해싱을 위한 라이브러리 추가

# Google Calendar API 관련 라이브러리 추가
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import base64

# --- 파일 이름 유효성 검사 함수 ---
def is_daily_schedule(file_name):
    """
    파일명이 'ocs_MMDD.xlsx' 또는 'ocs_MMDD.xlsm' 형식인지 확인합니다.
    """
    pattern = r'^ocs_\\d{4}\\.(?:xlsx|xlsm)$'
    return re.match(pattern, file_name, re.IGNORECASE) is not None
    
# --- 이메일 유효성 검사 함수 ---
def is_valid_email(email):
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(email_regex, email) is not None

# --- 비밀번호 해싱 함수 ---
def hash_password(password):
    """입력된 비밀번호를 SHA256으로 해싱합니다."""
    return hashlib.sha256(password.encode()).hexdigest()

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


#2. Excel and Email Processing Functions

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
        file_io = io.BytesIO(uploaded_file.getvalue())
        wb = load_workbook(file_io, data_only=True)
        return wb, file_io
    except Exception as e:
        st.error(f"엑셀 파일 로드 중 오류 발생: {e}")
        return None, None
    
# 데이터 처리 및 스타일링
def process_excel_file_and_style(file_io):
    try:
        raw_df = pd.read_excel(file_io)
        excel_data_dfs = pd.read_excel(file_io, sheet_name=None)
        return excel_data_dfs, raw_df.to_excel(index=False, header=True, engine='xlsxwriter')
    except Exception as e:
        st.error(f"엑셀 데이터 처리 및 스타일링 중 오류 발생: {e}")
        return None, None
    
# 이메일 전송 함수
def send_email(to_email, subject, content):
    st.info("실제 이메일 전송 로직을 여기에 구현하세요.")
    # 실제로는 smtplib 등을 사용하여 이메일을 보냅니다.
    # 예:
    # try:
    #     msg = MIMEMultipart()
    #     msg['From'] = 'your_email@example.com'
    #     msg['To'] = to_email
    #     msg['Subject'] = subject
    #     msg.attach(MIMEText(content, 'plain'))
    #     server = smtplib.SMTP('smtp.example.com', 587)
    #     server.starttls()
    #     server.login('your_email@example.com', 'your_password')
    #     server.send_message(msg)
    #     server.quit()
    #     st.success("이메일 전송 성공!")
    # except Exception as e:
    #     st.error(f"이메일 전송 실패: {e}")


#3. Google Calendar API Functions

SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

def get_google_calendar_service(user_id_safe):
    creds = load_google_creds_from_firebase(user_id_safe)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        save_google_creds_to_firebase(user_id_safe, creds)
    
    try:
        service = build('calendar', 'v3', credentials=creds)
        return service
    except HttpError as error:
        st.error(f"Google Calendar API 연결 오류: {error}")
        return None

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
#4. Excel Processing Constants and Functions
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
        
        non_professors_df['예약시간'] = non_professors_df['예약시간'].astype(str).str.strip()
        non_professors_df = non_professors_df[non_professors_df['예약시간'] != 'nan']
        
        non_professors_df['예약시간'] = non_professors_df['예약시간'].astype(str).str.strip()
        
        morning_patients = non_professors_df[
            (non_professors_df['예약시간'] >= '08:00') & 
            (non_professors_df['예약시간'] <= '12:50')
        ].shape[0]
        
        afternoon_patients = non_professors_df[
            non_professors_df['예약시간'] >= '13:00'
        ].shape[0]

        if afternoon_patients > 0:
            afternoon_patients -= 1
        analysis_results['소치'] = {'오전': morning_patients, '오후': afternoon_patients}

    # 보존과 분석
    if '보존' in mapped_dfs:
        df = mapped_dfs['보존']
        non_professors_df = df[~df['예약의사'].isin(professors_dict.get('보존', []))]
        
        non_professors_df['예약시간'] = non_professors_df['예약시간'].astype(str).str.strip()
        non_professors_df = non_professors_df[non_professors_df['예약시간'] != 'nan']

        non_professors_df['예약시간'] = non_professors_df['예약시간'].astype(str).str.strip()
        
        morning_patients = non_professors_df[
            (non_professors_df['예약시간'] >= '08:00') & 
            (non_professors_df['예약시간'] <= '12:30')
        ].shape[0]
        
        afternoon_patients = non_professors_df[
            non_professors_df['예약시간'] >= '12:50'
        ].shape[0]

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
        bonding_patients_df['예약시간'] = bonding_patients_df['예약시간'].astype(str).str.strip()
        
        morning_bonding = bonding_patients_df[
            (bonding_patients_df['예약시간'] >= '08:00') & 
            (bonding_patients_df['예약시간'] <= '12:50')
        ].shape[0]
        
        afternoon_bonding = bonding_patients_df[
            bonding_patients_df['예약시간'] >= '13:00'
        ].shape[0]
        
        analysis_results['교정'] = {'오전 본딩': morning_bonding, '오후 본딩': afternoon_bonding}

    return analysis_results

# 교수 명단 딕셔너리
professors_dict = {
    '소치': ['소아치과교수1', '소아치과교수2'],
    '보존': ['보존과교수1', '보존과교수2'],
    '교정': ['교정과교수1', '교정과교수2'],
}


#5. Streamlit App Start and Session State
st.set_page_config(layout="wide", page_title="병원 환자 관리 대시보드")

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "current_role" not in st.session_state:
    st.session_state.current_role = "guest"
if "username" not in st.session_state:
    st.session_state.username = ""
if "firebase_key" not in st.session_state:
    st.session_state.firebase_key = ""


#6. User and Admin Login and User Management

def show_login_page():
    st.title("👨‍⚕️ OCS 환자 관리 시스템")
    st.markdown("### 로그인")

    username = st.text_input("사용자 이름", key="login_username")
    password = st.text_input("비밀번호", type="password", key="login_password")

    if st.button("로그인"):
        users_ref = db.reference('users')
        user_data = users_ref.child(username).get()
        
        if user_data and hash_password(password) == user_data.get('password'):
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.current_role = user_data.get('role', '일반사용자')
            st.success(f"로그인 성공! ({st.session_state.current_role} 모드)")
            time.sleep(1)
            st.rerun()
        else:
            st.error("사용자 이름 또는 비밀번호가 올바르지 않습니다.")

def show_user_management():
    if st.session_state.current_role != "admin":
        st.error("이 기능은 관리자만 사용할 수 있습니다.")
        return

    st.header("➕ 사용자 관리 (관리자 전용)")
    st.markdown("새로운 사용자의 계정을 생성하거나 기존 사용자를 관리합니다.")

    new_username = st.text_input("새 사용자 이름", key="new_user")
    new_password = st.text_input("새 비밀번호", type="password", key="new_password")
    role_options = ["admin", "레지던트", "일반사용자"]
    new_role = st.selectbox("역할 선택", role_options)

    if st.button("사용자 계정 생성"):
        if not new_username or not new_password:
            st.error("사용자 이름과 비밀번호를 모두 입력해주세요.")
        else:
            users_ref = db.reference('users')
            if users_ref.child(new_username).get():
                st.error("이미 존재하는 사용자 이름입니다. 다른 이름을 사용해주세요.")
            else:
                try:
                    users_ref.child(new_username).set({
                        'password': hash_password(new_password),
                        'role': new_role
                    })
                    st.success(f"사용자 '{new_username}' ({new_role}) 계정이 성공적으로 생성되었습니다.")
                    st.rerun()
                except Exception as e:
                    st.error(f"사용자 등록 중 오류가 발생했습니다: {e}")

    st.markdown("---")
    st.subheader("등록된 사용자 목록")
    users_ref = db.reference('users')
    users_data = users_ref.get()
    if users_data:
        users_df = pd.DataFrame.from_dict(users_data, orient='index')
        users_df.index.name = "사용자 이름"
        users_df.reset_index(inplace=True)
        st.dataframe(users_df[['사용자 이름', 'role']])


#7. Admin Mode
def show_admin_mode():
    st.sidebar.title("관리자 모드 메뉴")
    st.sidebar.markdown(f"**사용자:** {st.session_state.username}")
    menu = st.sidebar.radio("작업 선택", [
        "환자 명단 보기", "환자 등록/수정", "사용자 관리", "비밀번호 변경", "환자 상태 변경", "엑셀 업로드", "로그아웃"
    ])
    
    st.title("병원 환자 관리 대시보드 (관리자)")
    st.write(f"현재 모드: **{st.session_state.current_role}**")
    
    if menu == "환자 명단 보기":
        st.header("📋 환자 명단")
        patients_ref = db.reference('/patients')
        patient_data = patients_ref.get()
        if patient_data:
            df = pd.DataFrame.from_dict(patient_data, orient='index')
            st.dataframe(df)
        else:
            st.info("등록된 환자 데이터가 없습니다.")

    elif menu == "환자 등록/수정":
        st.header("✍️ 환자 등록 및 수정")
        name = st.text_input("환자명")
        pid = st.text_input("진료번호")
        if st.button("환자 등록"):
            if not name or not pid:
                st.error("환자명과 진료번호를 모두 입력해주세요.")
            else:
                st.success(f"{name} ({pid}) 환자 등록 완료!")

    elif menu == "사용자 관리":
        show_user_management()

    elif menu == "비밀번호 변경":
        st.header("🔑 비밀번호 변경")
        new_password = st.text_input("새 비밀번호", type="password")
        confirm_password = st.text_input("새 비밀번호 확인", type="password")
        if st.button("비밀번호 변경 완료"):
            if new_password == confirm_password and new_password:
                users_ref = db.reference('users')
                users_ref.child(st.session_state.username).update({'password': hash_password(new_password)})
                st.success("비밀번호가 성공적으로 변경되었습니다.")
            else:
                st.error("비밀번호가 일치하지 않거나 비어있습니다.")
                
    elif menu == "환자 상태 변경":
        st.header("🩺 환자 상태 변경")
        st.selectbox("환자 선택", ["환자 A", "환자 B"])
        st.selectbox("상태 변경", ["입원", "퇴원", "전원"])
        if st.button("상태 변경"):
            st.success("환자 상태가 변경되었습니다.")

    elif menu == "엑셀 업로드":
        st.header("📊 OCS 엑셀 파일 업로드")
        uploaded_file = st.file_uploader("OCS 파일을 업로드하세요 (ocs_MMDD.xlsx/xlsm)", type=["xlsx", "xlsm"])
        if uploaded_file:
            if not is_daily_schedule(uploaded_file.name):
                st.error("파일명 형식이 올바르지 않습니다.")
            else:
                try:
                    file_content = uploaded_file.getvalue()
                    if msoffcrypto.OfficeFile(io.BytesIO(file_content)).is_encrypted():
                        password_input = st.text_input("파일 암호를 입력하세요", type="password")
                        if st.button("파일 복호화"):
                            try:
                                with io.BytesIO(file_content) as encrypted_file:
                                    office_file = msoffcrypto.OfficeFile(encrypted_file)
                                    office_file.load_key(password=password_input)
                                    decrypted_file = io.BytesIO()
                                    office_file.decrypt(decrypted_file)
                                    decrypted_file.seek(0)
                                    df = pd.read_excel(decrypted_file)
                                    st.success("파일 복호화 및 업로드 완료!")
                                    st.dataframe(df.head())
                                    st.info("실제 데이터베이스 업로드 로직을 여기에 구현하세요.")
                            except msoffcrypto.exceptions.InvalidKeyError:
                                st.error("잘못된 파일 암호입니다. 다시 시도해주세요.")
                            except Exception as e:
                                st.error(f"파일 복호화 중 예상치 못한 오류가 발생했습니다: {e}")
                    else:
                        df = pd.read_excel(io.BytesIO(file_content))
                        st.success("엑셀 파일 업로드 완료!")
                        st.dataframe(df.head())
                        st.info("실제 데이터베이스 업로드 로직을 여기에 구현하세요.")
                except Exception as e:
                    st.error(f"파일을 처리하는 중 오류가 발생했습니다: {e}")

    elif menu == "로그아웃":
        st.session_state.logged_in = False
        st.session_state.current_role = "guest"
        st.session_state.username = ""
        st.info("로그아웃 되었습니다.")
        time.sleep(1)
        st.rerun()


#8. Regular User Mode
def show_regular_user_mode():
    st.sidebar.title("일반 사용자 모드 메뉴")
    st.sidebar.markdown(f"**사용자:** {st.session_state.username}")
    menu = st.sidebar.radio("작업 선택", [
        "환자 명단 보기", "비밀번호 변경", "로그아웃"
    ])
    
    st.title("병원 환자 관리 대시보드 (일반 사용자)")
    st.write(f"현재 모드: **{st.session_state.current_role}**")
    
    if menu == "환자 명단 보기":
        st.header("📋 환자 명단")
        patients_ref = db.reference('/patients')
        patient_data = patients_ref.get()
        if patient_data:
            df = pd.DataFrame.from_dict(patient_data, orient='index')
            st.dataframe(df)
        else:
            st.info("등록된 환자 데이터가 없습니다.")

    elif menu == "비밀번호 변경":
        st.header("🔑 비밀번호 변경")
        new_password = st.text_input("새 비밀번호", type="password")
        confirm_password = st.text_input("새 비밀번호 확인", type="password")
        if st.button("비밀번호 변경 완료"):
            if new_password == confirm_password and new_password:
                users_ref = db.reference('users')
                users_ref.child(st.session_state.username).update({'password': hash_password(new_password)})
                st.success("비밀번호가 성공적으로 변경되었습니다.")
            else:
                st.error("비밀번호가 일치하지 않거나 비어있습니다.")

    elif menu == "로그아웃":
        st.session_state.logged_in = False
        st.session_state.current_role = "guest"
        st.session_state.username = ""
        st.info("로그아웃 되었습니다.")
        time.sleep(1)
        st.rerun()


#9. Resident User Mode
def show_resident_mode():
    st.sidebar.title("레지던트 모드 메뉴")
    st.sidebar.markdown(f"**사용자:** {st.session_state.username}")
    menu = st.sidebar.radio("작업 선택", [
        "환자 명단 보기", "환자 등록/수정", "비밀번호 변경", "환자 상태 변경", "로그아웃"
    ])
    
    st.title("병원 환자 관리 대시보드 (레지던트)")
    st.write(f"현재 모드: **{st.session_state.current_role}**")
    
    if menu == "환자 명단 보기":
        st.header("📋 환자 명단")
        patients_ref = db.reference('/patients')
        patient_data = patients_ref.get()
        if patient_data:
            df = pd.DataFrame.from_dict(patient_data, orient='index')
            st.dataframe(df)
        else:
            st.info("등록된 환자 데이터가 없습니다.")

    elif menu == "환자 등록/수정":
        st.header("✍️ 환자 등록 및 수정")
        name = st.text_input("환자명")
        pid = st.text_input("진료번호")
        if st.button("환자 등록"):
            if not name or not pid:
                st.error("환자명과 진료번호를 모두 입력해주세요.")
            else:
                st.success(f"{name} ({pid}) 환자 등록 완료!")
    
    elif menu == "비밀번호 변경":
        st.header("🔑 비밀번호 변경")
        new_password = st.text_input("새 비밀번호", type="password")
        confirm_password = st.text_input("새 비밀번호 확인", type="password")
        if st.button("비밀번호 변경 완료"):
            if new_password == confirm_password and new_password:
                users_ref = db.reference('users')
                users_ref.child(st.session_state.username).update({'password': hash_password(new_password)})
                st.success("비밀번호가 성공적으로 변경되었습니다.")
            else:
                st.error("비밀번호가 일치하지 않거나 비어있습니다.")
                
    elif menu == "환자 상태 변경":
        st.header("🩺 환자 상태 변경")
        st.selectbox("환자 선택", ["환자 A", "환자 B"])
        st.selectbox("상태 변경", ["입원", "퇴원", "전원"])
        if st.button("상태 변경"):
            st.success("환자 상태가 변경되었습니다.")

    elif menu == "로그아웃":
        st.session_state.logged_in = False
        st.session_state.current_role = "guest"
        st.session_state.username = ""
        st.info("로그아웃 되었습니다.")
        time.sleep(1)
        st.rerun()

# 10. 메인 실행 로직
if st.session_state.logged_in:
    if st.session_state.current_role == "admin":
        show_admin_mode()
    elif st.session_state.current_role == "레지던트":
        show_resident_mode()
    else:
        show_regular_user_mode()
else:
    show_login_page()
