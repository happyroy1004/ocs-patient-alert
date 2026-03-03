# streamlit_app.py
import streamlit as st
from ui_manager import (
    init_session_state, show_title_and_manual, show_login_and_registration, 
    show_admin_mode_ui, show_user_mode_ui, show_doctor_mode_ui
) [cite: 1]

st.set_page_config(layout="wide") [cite: 1]
init_session_state() [cite: 1]
show_title_and_manual() [cite: 1]

mode = st.session_state.login_mode [cite: 1]
if mode in ['not_logged_in', 'new_user_registration', 'new_doctor_registration']:
    show_login_and_registration() [cite: 1]
elif mode == 'admin_mode': show_admin_mode_ui() [cite: 1]
elif mode == 'user_mode': show_user_mode_ui(st.session_state.current_firebase_key, st.session_state.current_user_name) [cite: 1]
elif mode == 'doctor_mode': show_doctor_mode_ui(st.session_state.current_firebase_key, st.session_state.current_user_name) [cite: 1]
