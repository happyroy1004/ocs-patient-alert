# streamlit_app.py

import streamlit as st
import datetime
import os
import re

# 모듈 임포트
from ui_manager import (
    init_session_state, show_title_and_manual, show_login_and_registration, 
    show_admin_mode_ui, show_user_mode_ui, show_doctor_mode_ui
)
from firebase_utils import get_db_refs, sanitize_path

# --- 1. 초기 설정 및 Firebase 초기화 ---
st.set_page_config(layout="wide")
init_firebase() # Firebase 초기화
# DB 레퍼런스 초기 로드 (ui_manager에서도 사용)
users_ref, doctor_users_ref, db_ref_func = get_db_refs()
init_session_state()

# --- 2. 초기화 및 리다이렉션 처리 ---
# Query Params를 이용한 상태 클리어 처리
if "clear" in st.query_params and st.query_params["clear"] == "true":
    st.session_state.clear()
    st.query_params["clear"] = "false"
    st.rerun()

# --- 3. UI 제목 및 사용 설명서 ---
show_title_and_manual()

# --- 4. 메인 앱 흐름 제어 ---
login_mode = st.session_state.get('login_mode')
firebase_key = st.session_state.get('current_firebase_key', "")
user_name = st.session_state.get('current_user_name', "")
user_id_final = st.session_state.get('found_user_email', "")
user_role = st.session_state.get('current_user_role', 'user')


# A. 로그인/등록/비관리자 모드
if login_mode in ['not_logged_in', 'new_user_registration', 'new_doctor_registration']:
    # 로그인 폼이 아직 필요하거나, 등록 정보 입력 단계인 경우
    show_login_and_registration()
    
# B. 관리자 모드
elif login_mode == 'admin_mode':
    show_admin_mode_ui()

# C. 일반 사용자 또는 치과의사 모드
elif login_mode in ['user_mode', 'doctor_mode']:
    
    # 이메일 주소 변경 기능 (모든 사용자 공통)
    if firebase_key and st.session_state.get('email_change_mode') is not True:
        st.divider()
        st.text_input("아이디 (등록된 이메일)", value=user_id_final, disabled=True)
        if st.button("이메일 주소 변경"):
            st.session_state.email_change_mode = True; st.rerun()
    
    if st.session_state.get('email_change_mode'):
        # ... (이메일 변경 UI/로직 - ui_manager.py에 포함하거나 여기에 간단히 작성)
        st.divider(); st.subheader("이메일 주소 변경")
        new_email_input = st.text_input("새 이메일 주소를 입력하세요", value=user_id_final)
        
        if st.button("변경 완료"):
            if is_valid_email(new_email_input):
                 new_firebase_key = sanitize_path(new_email_input)
                 # DB 업데이트 로직 (firebase_utils에 함수화 필요)
                 # ... (DB 업데이트 로직 생략)
                 st.session_state.email_change_mode = False; st.success("이메일 주소가 변경되었습니다."); st.rerun()
            else: st.error("올바른 이메일 주소 형식이 아닙니다.")

    else:
        # 최종적으로 사용자 모드별 UI 표시
        if user_role == 'doctor':
            show_doctor_mode_ui(firebase_key, user_name)
        else:
            show_user_mode_ui(firebase_key, user_name)
