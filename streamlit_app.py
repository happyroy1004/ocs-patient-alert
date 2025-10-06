# streamlit_app.py

import streamlit as st
import datetime
import os
import re

# 모듈 임포트: ui_manager가 DB 초기화 및 모든 로컬 모듈을 간접적으로 처리합니다.
from ui_manager import (
    init_session_state, show_title_and_manual, show_login_and_registration, 
    show_admin_mode_ui, show_user_mode_ui, show_doctor_mode_ui
)

# --- 1. 초기 설정 및 상태 클리어 ---
st.set_page_config(layout="wide")

# Query Params를 이용한 상태 클리어 처리 (이전 코드 유지)
if "clear" in st.query_params and st.query_params["clear"] == "true":
    st.session_state.clear()
    st.query_params["clear"] = "false"
    st.rerun()

# --- 2. 메인 실행 흐름 ---

# 세션 상태 초기화
init_session_state() 

show_title_and_manual() # 제목 및 매뉴얼 표시

# 🔑 핵심: 로그인 상태에 따른 UI 분기 (로그인 창 복원)
# 'not_logged_in', 'new_user_registration', 'new_doctor_registration' 상태일 때 로그인/등록 UI를 표시합니다.
if st.session_state.login_mode == 'not_logged_in' or \
   st.session_state.login_mode == 'new_user_registration' or \
   st.session_state.login_mode == 'new_doctor_registration':
    show_login_and_registration()

elif st.session_state.login_mode == 'admin_mode':
    show_admin_mode_ui()

elif st.session_state.login_mode == 'user_mode':
    show_user_mode_ui(st.session_state.current_firebase_key, st.session_state.current_user_name)

elif st.session_state.login_mode == 'doctor_mode':
    show_doctor_mode_ui(st.session_state.current_firebase_key, st.session_state.current_user_name)
