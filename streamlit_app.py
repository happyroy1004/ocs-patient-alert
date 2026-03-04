# streamlit_app.py (수정 전체 코드)

import streamlit as st
import datetime
import os
import re

# 모듈 임포트: ui_manager가 DB 초기화 및 모든 로컬 모듈을 간접적으로 처리합니다.
from ui_manager import (
    init_session_state, show_title_and_manual, show_login_and_registration, 
    show_admin_mode_ui, show_user_mode_ui, show_doctor_mode_ui
    # 💡 [변경]: show_professor_review_system은 이제 ui_manager 내에서 호출됩니다.
)

# 💡 [변경]: professor_reviews_module은 더 이상 여기서 직접 임포트하지 않습니다.
# from professor_reviews_module import show_professor_review_system 

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

# 🔑 핵심: 로그인 상태에 따른 UI 분기 (이전 구조로 복귀)

# 'not_logged_in', 'new_user_registration', 'new_doctor_registration' 상태일 때 로그인/등록 UI를 표시합니다.
if st.session_state.login_mode == 'not_logged_in' or \
   st.session_state.login_mode == 'new_user_registration' or \
   st.session_state.login_mode == 'new_doctor_registration':
    show_login_and_registration()

elif st.session_state.login_mode == 'admin_mode':
    show_admin_mode_ui()

elif st.session_state.login_mode == 'user_mode':
    # 🔑 [변경]: 학생 모드(user_mode)에서 교수님 평가 시스템이 통합됩니다.
    show_user_mode_ui(st.session_state.current_firebase_key, st.session_state.current_user_name)

elif st.session_state.login_mode == 'doctor_mode':
    show_doctor_mode_ui(st.session_state.current_firebase_key, st.session_state.current_user_name)

# 💡 최상단 탭 구조 제거: 로그인 후 모드별로 탭이 나오도록 분기를 단순화했습니다.
