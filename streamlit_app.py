# streamlit_app.py (수정 전체 코드)

import streamlit as st
import datetime
import os
import re

# 모듈 임포트: ui_manager는 기존 앱의 핵심 UI를 처리합니다.
from ui_manager import (
    init_session_state, show_title_and_manual, show_login_and_registration, 
    show_admin_mode_ui, show_user_mode_ui, show_doctor_mode_ui
)

# 💡 [추가] 새로운 교수님 평가 모듈 임포트
from professor_reviews_module import show_professor_review_system 

# --- 1. 초기 설정 및 상태 클리어 ---
st.set_page_config(layout="wide")

# Query Params를 이용한 상태 클리어 처리 (기존 코드 유지)
if "clear" in st.query_params and st.query_params["clear"] == "true":
    st.session_state.clear()
    st.query_params["clear"] = "false"
    st.rerun()

# --- 2. 메인 실행 흐름 ---

# 세션 상태 초기화
init_session_state() 

show_title_and_manual() # 제목 및 매뉴얼 표시

# 🔑 핵심: 메인 시스템과 평가 시스템을 분리하는 탭 추가
main_app_tab, review_tab = st.tabs(["메인 시스템 (로그인 필요)", "🧑‍🏫 교수님 평가 시스템"])


# --- 2-1. 메인 시스템 탭 (로그인 기반) ---
with main_app_tab:
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

# --- 2-2. 교수님 평가 시스템 탭 (로그인 불필요) ---
with review_tab:
    # 💡 [추가] 별도 모듈의 UI 함수 호출
    show_professor_review_system()
