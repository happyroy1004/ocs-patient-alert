import streamlit as st
import pandas as pd
import bcrypt

# local imports
from config import (
    DEFAULT_PASSWORD, DEPARTMENTS_FOR_REGISTRATION, PATIENT_DEPT_FLAGS
)
from firebase_utils import (
    get_db_refs, sanitize_path, recover_email, 
    get_google_calendar_service, save_google_creds_to_firebase, 
    load_google_creds_from_firebase
)
import excel_utils
from notification_utils import (
    is_valid_email, send_email, create_calendar_event, 
    get_matching_data, run_auto_notifications
)

# [핵심] firebase_utils의 get_db_refs() 반환값 3개를 정확히 수신
users_ref, doctor_users_ref, db_ref_func = get_db_refs()

# --- 비밀번호 보안 관련 ---
def hash_password(password):
    salt = bcrypt.gensalt()
    return bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')

def check_password(password, hashed_password):
    if not hashed_password or not isinstance(hashed_password, str):
        return False
    try:
        return bcrypt.checkpw(password.encode('utf-8'), hashed_password.encode('utf-8'))
    except Exception:
        return False

# --- 1. 세션 상태 초기화 ---
def init_session_state():
    if 'login_mode' not in st.session_state: st.session_state.login_mode = 'not_logged_in'
    if 'current_firebase_key' not in st.session_state: st.session_state.current_firebase_key = ""
    if 'current_user_name' not in st.session_state: st.session_state.current_user_name = ""
    if 'admin_password_correct' not in st.session_state: st.session_state.admin_password_correct = False

def show_title_and_manual():
    st.markdown("<h1>환자 내원 확인 시스템</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: grey;'>directed by HSY</p>", unsafe_allow_html=True)

# --- 2. 로그인 및 등록 로직 (안정성 강화) ---
def _handle_login(user_name, password_input, role="student"):
    """role에 따라 학생/의사 DB를 명확히 분리하여 검색합니다."""
    clean_name = user_name.strip() # 공백으로 인한 로그인 실패 방지
    if not clean_name: 
        st.error("이름을 입력해주세요.")
        return
    
    target_ref = users_ref if role == "student" else doctor_users_ref
    
    try:
        db_data = target_ref.get()
    except Exception as e:
        st.error(f"데이터베이스 연결 오류: {e}")
        return

    # 데이터베이스에 대상이 있는 경우 탐색
    if db_data and isinstance(db_data, dict):
        for safe_key, info in db_data.items():
            if info.get("name") == clean_name:
                # 비밀번호 확인 로직
                is_valid = False
                if role == "student":
                    is_valid = check_password(password_input, info.get("password"))
                else: # doctor
                    # 의사는 기본 암호(DEFAULT_PASSWORD)를 허용하거나, 해시된 비밀번호가 있다면 확인
                    if password_input == DEFAULT_PASSWORD or check_password(password_input, info.get("password")):
                        is_valid = True
                
                if is_valid:
                    st.session_state.update({
                        'current_firebase_key': safe_key, 
                        'current_user_name': clean_name, 
                        'login_mode': 'user_mode' if role == "student" else 'doctor_mode'
                    })
                    st.rerun()
                else: 
                    st.error("비밀번호가 일치하지 않습니다.")
                return # 이름이 일치하면 성공/실패 여부와 관계없이 함수 종료

    # 검색이 끝났는데도 return되지 않았다면 정보가 없는 것
    if role == "student":
        st.warning(f"'{clean_name}' 학생 정보가 없습니다. 신규 등록을 진행합니다.")
        st.session_state.current_user_name = clean_name
        st.session_state.login_mode = 'new_user_registration'
        st.rerun()
    else:
        st.error(f"등록된 치과의사 '{clean_name}' 정보를 찾을 수 없습니다. 관리자에게 문의하세요.")

# --- 3. 로그인 및 등록 UI (탭 분리) ---
def show_login_and_registration():
    if st.session_state.login_mode == 'not_logged_in':
        # [관리자 로그인] 사이드바를 통해 문열기
        with st.sidebar:
            st.subheader("💻 시스템 관리")
            admin_pw = st.text_input("관리자 암호", type="password")
            if st.button("관리자 모드 진입"):
                try:
                    admin_secret_pw = st.secrets["admin"]["password"]
                except:
                    admin_secret_pw = "1243"

                if admin_pw == admin_secret_pw:
                    st.session_state.login_mode = 'admin_mode'
                    st.session_state.admin_password_correct = True
                    st.rerun()
                else:
                    st.error("암호가 올바르지 않습니다.")

        # [일반 사용자 로그인 UI - 탭 분리]
        st.subheader("시스템 로그인")
        tab_student, tab_doctor = st.tabs(["🎓 학생", "👨‍⚕️ 치과의사"])
        
        with tab_student:
            st.markdown("##### 학생 로그인")
            s_name = st.text_input("성함", key="s_name_input")
            s_pw = st.text_input("비밀번호", type="password", key="s_pw_input")
            if st.button("학생 로그인", use_container_width=True):
                _handle_login(s_name, s_pw, role="student")
                
        with tab_doctor:
            st.markdown("##### 치과의사 로그인")
            d_name = st.text_input("성함", key="d_name_input")
            d_pw = st.text_input("비밀번호", type="password", key="d_pw_input")
            if st.button("치과의사 로그인", use_container_width=True):
                _handle_login(d_name, d_pw, role="doctor")
    
    elif st.session_state.login_mode == 'new_user_registration':
        st.subheader("🎓 신규 학생 등록")
        st.info(f"환영합니다, **{st.session_state.current_user_name}**님! 처음 접속하셨군요. 계정을 생성해주세요.")
        
        email = st.text_input("이메일 주소 (ID 및 알림 수신용)")
        pw = st.text_input("사용할 비밀번호", type="password")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("등록 완료", use_container_width=True):
                if not is_valid_email(email):
                    st.error("올바른 이메일 형식을 입력해주세요.")
                elif not pw:
                    st.error("비밀번호를 입력해주세요.")
                else:
                    safe_key = sanitize_path(email)
                    users_ref.child(safe_key).set({
                        "name": st.session_state.current_user_name, 
                        "email": email, 
                        "password": hash_password(pw)
                    })
                    st.session_state.update({'current_firebase_key': safe_key, 'login_mode': 'user_mode'})
                    st.success("등록되었습니다!")
                    st.rerun()
        with col2:
            if st.button("취소 (돌아가기)", use_container_width=True):
                st.session_state.login_mode = 'not_logged_in'
                st.rerun()

# --- 4. 관리자 모드 UI ---
def show_admin_mode_ui():
    st.title("🛡️ 관리자 대시보드")
    
    if st.button("← 로그아웃"):
        st.session_state.login_mode = 'not_logged_in'
        st.rerun()

    uploaded_file = st.file_uploader("OCS 엑셀 파일을 업로드하세요", type=["xlsx", "xlsm"])
    
    if uploaded_file:
        is_daily = excel_utils.is_daily_schedule(uploaded_file.name)
        try:
            xl_data, styled_file = excel_utils.process_excel_file_and_style(uploaded_file)
            st.success(f"파일 분석 완료: {uploaded_file.name}")
            
            if st.button("🚀 전체 자동 알림 전송 시작"):
                with st.spinner("데이터 매칭 및 전송 중..."):
                    all_patients = db_ref_func("patients").get()
                    all_users = users_ref.get()
                    all_doctors = doctor_users_ref.get()
                    
                    matched_users, matched_docs = get_matching_data(
                        xl_data, all_users, all_patients, all_doctors
                    )
                    run_auto_notifications(
                        matched_users, matched_docs, xl_data, 
                        uploaded_file.name, is_daily, db_ref_func
                    )
                st.balloons()
                st.success("모든 알림 전송 프로세스가 완료되었습니다.")
        except Exception as e:
            st.error(f"엑셀 처리 중 오류 발생: {e}")

# --- 5. 일반 사용자(학생) 모드 UI ---
def show_user_mode_ui(firebase_key, user_name):
    patients_ref = db_ref_func(f"patients/{firebase_key}")
    
    st.subheader(f"🎓 {user_name} 학생님")
    
    # 구글 캘린더 연동 상태 체크
    get_google_calendar_service(firebase_key)
    
    tab_reg, tab_list = st.tabs(["🆕 환자 등록", "📋 목록 관리"])
    
    with tab_reg:
        with st.form("reg_form", clear_on_submit=True):
            p_name = st.text_input("환자 이름")
            p_id = st.text_input("진료 번호 (8자리)")
            depts = st.multiselect("담당 진료과", DEPARTMENTS_FOR_REGISTRATION)
            
            if st.form_submit_button("환자 등록"):
                if not p_name or not p_id or not depts:
                    st.error("모든 항목을 입력해주세요.")
                else:
                    p_id_clean = p_id.strip().zfill(8)
                    p_data = {"환자이름": p_name, "진료번호": p_id_clean}
                    for d in PATIENT_DEPT_FLAGS: 
                        p_data[d.lower()] = (d in depts)
                    patients_ref.child(p_id_clean).set(p_data)
                    st.success(f"{p_name} 환자가 등록되었습니다.")

    with tab_list:
        data = patients_ref.get()
        if data:
            for pid, val in data.items():
                col1, col2 = st.columns([4, 1])
                col1.info(f"**{val.get('환자이름')}** ({pid})")
                if col2.button("삭제", key=f"del_{pid}"):
                    patients_ref.child(pid).delete()
                    st.rerun()
        else:
            st.write("등록된 환자가 없습니다.")

    st.markdown("---")
    if st.button("로그아웃"):
        st.session_state.login_mode = 'not_logged_in'
        st.rerun()

# --- 6. 치과의사 모드 UI ---
def show_doctor_mode_ui(firebase_key, doctor_name):
    st.subheader(f"👨‍⚕️ {doctor_name} 의사님")
    
    # 구글 캘린더 연동 확인
    get_google_calendar_service(firebase_key)
    
    st.info("의사님께 배정된 환자 내원 정보는 시스템에 의해 자동으로 구글 캘린더에 동기화됩니다.")
    
    doc_info = doctor_users_ref.child(firebase_key).get()
    if doc_info:
        st.write(f"📧 연동 이메일: {doc_info.get('email', '미지정')}")
        st.write(f"🏥 소속 과: {doc_info.get('department', '미지정')}")

    st.markdown("---")
    if st.button("로그아웃"):
        st.session_state.login_mode = 'not_logged_in'
        st.rerun()
