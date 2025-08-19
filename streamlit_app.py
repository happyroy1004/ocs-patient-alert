# 1. Imports, Validation Functions, and Firebase Initialization
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
import openpyxl  # 추가
import datetime  # 추가
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
    # 'ocs_날짜(4자리).확장자' 패턴을 찾음 (예: ocs_0815.xlsx)
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
            'databaseURL': st.secrets["firebase"]["FIREBASE_DATABASE_URL"]
        })
    except Exception as e:
        st.error("Firebase 초기화 중 오류가 발생했습니다. Streamlit Secrets 설정 파일을 확인해주세요.")
        st.error(f"오류: {e}")

# 초기 사용자 등록 (최초 실행 시에만)
users_ref = db.reference('users')
if users_ref.get() is None:
    st.info("Firebase 데이터베이스에 초기 사용자를 등록합니다.")
    users_ref.child('admin').set({'password': hash_password('admin_password'), 'role': 'admin'})
    users_ref.child('레지던트').set({'password': hash_password('resident_password'), 'role': '레지던트'})
    users_ref.child('일반사용자').set({'password': hash_password('user_password'), 'role': '일반사용자'})
    st.success("초기 사용자 등록 완료! 'admin' / 'admin_password' 로 로그인해보세요.")


# --- 사용자 역할에 따라 UI를 다르게 표시하기 위한 세션 상태 초기화 ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "current_role" not in st.session_state:
    st.session_state.current_role = "guest"
if "username" not in st.session_state:
    st.session_state.username = ""
    

# 2. 로그인 및 사용자 인터페이스
def show_login_page():
    st.title("👨‍⚕️ OCS 환자 관리 시스템")
    st.markdown("### 로그인")

    # 사용자 이름과 비밀번호 입력 필드
    st.session_state.username = st.text_input("사용자 이름", key="login_username")
    password = st.text_input("비밀번호", type="password", key="login_password")

    if st.button("로그인"):
        # Firebase에서 사용자 데이터 가져오기
        users_ref = db.reference('users')
        user_data = users_ref.child(st.session_state.username).get()
        
        # 입력된 사용자 이름과 비밀번호를 확인
        if user_data and hash_password(password) == user_data.get('password'):
            st.session_state.logged_in = True
            st.session_state.current_role = user_data.get('role', '일반사용자') # 역할이 없으면 일반사용자로 설정
            st.success(f"로그인 성공! ({st.session_state.current_role} 모드)")
            time.sleep(1)
            st.rerun()
        else:
            st.error("사용자 이름 또는 비밀번호가 올바르지 않습니다.")

def show_main_page():
    # --- 사이드바 메뉴 ---
    st.sidebar.title("메뉴")
    
    # 역할에 따라 사이드바 메뉴 제목 변경
    if st.session_state.current_role == "admin":
        st.sidebar.subheader("관리자 모드")
        st.sidebar.markdown(f"**사용자:** {st.session_state.username}")
        # 관리자 메뉴에 사용자 등록 추가
        menu = st.sidebar.radio("작업 선택", [
            "환자 명단 보기", "환자 등록/수정", "사용자 등록", "비밀번호 변경", "환자 상태 변경", "로그아웃"
        ])
    elif st.session_state.current_role == "레지던트":
        st.sidebar.subheader("레지던트 모드")
        st.sidebar.markdown(f"**사용자:** {st.session_state.username}")
        menu = st.sidebar.radio("작업 선택", [
            "환자 명단 보기", "환자 등록/수정", "비밀번호 변경", "환자 상태 변경", "로그아웃"
        ])
    else: # 일반 사용자 모드
        st.sidebar.subheader("일반 사용자 모드")
        st.sidebar.markdown(f"**사용자:** {st.session_state.username}")
        menu = st.sidebar.radio("작업 선택", [
            "환자 명단 보기", "비밀번호 변경", "로그아웃"
        ])
    
    st.title("병원 환자 관리 대시보드")
    st.write(f"현재 모드: **{st.session_state.current_role} 모드**")
    
    # 3. 엑셀 파일 업로드 기능 (관리자 모드에서만 보이도록 수정)
    if st.session_state.current_role == "admin":
        st.markdown("---")
        st.header("📊 OCS 엑셀 파일 업로드 (관리자 전용)")
        
        uploaded_file = st.file_uploader("OCS 파일을 업로드하세요 (ocs_MMDD.xlsx/xlsm)", type=["xlsx", "xlsm"])

        if uploaded_file:
            if not is_daily_schedule(uploaded_file.name):
                st.error("파일명 형식이 올바르지 않습니다. 'ocs_MMDD.xlsx' 또는 'ocs_MMDD.xlsm' 형식이어야 합니다.")
            else:
                try:
                    # 파일 내용 읽기
                    file_content = uploaded_file.getvalue()

                    # 암호화된 파일인 경우 복호화
                    if msoffcrypto.OfficeFile(io.BytesIO(file_content)).is_encrypted():
                        # 비밀번호 입력
                        password_input = st.text_input("파일 암호를 입력하세요", type="password")
                        decrypt_button = st.button("파일 복호화")
                        if decrypt_button:
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

                                    # 데이터베이스에 업로드 (더미 코드)
                                    st.info("실제 데이터베이스 업로드 로직을 여기에 구현하세요.")

                            except msoffcrypto.exceptions.InvalidKeyError:
                                st.error("잘못된 파일 암호입니다. 다시 시도해주세요.")
                            except Exception as e:
                                st.error(f"파일 복호화 중 예상치 못한 오류가 발생했습니다: {e}")
                    else:
                        df = pd.read_excel(io.BytesIO(file_content))
                        st.success("엑셀 파일 업로드 완료!")
                        st.dataframe(df.head())

                        # 데이터베이스에 업로드 (더미 코드)
                        st.info("실제 데이터베이스 업로드 로직을 여기에 구현하세요.")
                except Exception as e:
                    st.error(f"파일을 처리하는 중 오류가 발생했습니다: {e}")

    # 4. 기타 기능
    if menu == "환자 명단 보기":
        st.header("📋 환자 명단")
        st.write("환자 명단 데이터를 표시합니다.")
        # Firebase에서 환자 데이터 가져오기 (더미)
        patients_ref = db.reference('/patients')
        patient_data = patients_ref.get()
        if patient_data:
            df = pd.DataFrame.from_dict(patient_data, orient='index')
            st.dataframe(df)
        else:
            st.info("등록된 환자 데이터가 없습니다.")

    elif menu == "환자 등록/수정":
        st.header("✍️ 환자 등록 및 수정")
        st.write("환자 정보를 등록하거나 수정하는 기능입니다.")
        # 환자 등록/수정 UI (더미)
        name = st.text_input("환자명")
        pid = st.text_input("진료번호")
        if st.button("환자 등록"):
            if not name or not pid:
                st.error("환자명과 진료번호를 모두 입력해주세요.")
            else:
                st.success(f"{name} ({pid}) 환자 등록 완료!")

    elif st.session_state.current_role == "admin" and menu == "사용자 등록":
        st.header("➕ 사용자 등록 (관리자 전용)")
        st.markdown("새로운 사용자의 계정을 생성합니다.")
        
        new_username = st.text_input("새 사용자 이름")
        new_password = st.text_input("새 비밀번호", type="password")
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
                        # Firebase에 새 사용자 정보 저장
                        users_ref.child(new_username).set({
                            'password': hash_password(new_password),
                            'role': new_role
                        })
                        st.success(f"사용자 '{new_username}' ({new_role}) 계정이 성공적으로 생성되었습니다.")
                    except Exception as e:
                        st.error(f"사용자 등록 중 오류가 발생했습니다: {e}")


    elif menu == "비밀번호 변경":
        st.header("🔑 비밀번호 변경")
        st.write("사용자 비밀번호를 변경하는 기능입니다.")
        # 비밀번호 변경 UI (더미)
        new_password = st.text_input("새 비밀번호", type="password")
        confirm_password = st.text_input("새 비밀번호 확인", type="password")
        if st.button("비밀번호 변경 완료"):
            if new_password == confirm_password and new_password:
                st.success("비밀번호가 성공적으로 변경되었습니다.")
                # 실제 비밀번호 변경 로직을 여기에 구현해야 합니다.
            else:
                st.error("비밀번호가 일치하지 않거나 비어있습니다.")

    elif menu == "환자 상태 변경":
        st.header("🩺 환자 상태 변경")
        st.write("환자의 입원/퇴원/전원 상태를 변경하는 기능입니다.")
        # 환자 상태 변경 UI (더미)
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

# --- 페이지 렌더링 ---
if st.session_state.logged_in:
    show_main_page()
else:
    show_login_page()
