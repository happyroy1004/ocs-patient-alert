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
    if 'registration_role' not in st.session_state: st.session_state.registration_role = 'student'

def show_title_and_manual():
    st.markdown("<h1>환자 내원 확인 시스템</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: grey;'>directed by HSY</p>", unsafe_allow_html=True)

# --- 2. 로그인 및 등록 로직 ---
def _handle_login(user_name, password_input, role="student"):
    clean_name = user_name.strip()
    if not clean_name: 
        st.error("이름을 입력해주세요.")
        return
    
    target_ref = users_ref if role == "student" else doctor_users_ref
    
    try:
        db_data = target_ref.get()
    except Exception as e:
        st.error(f"데이터베이스 연결 오류: {e}")
        return

    if db_data and isinstance(db_data, dict):
        for safe_key, info in db_data.items():
            if info.get("name") == clean_name:
                is_valid = False
                if role == "student":
                    is_valid = check_password(password_input, info.get("password"))
                else: 
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
                return 

    st.warning(f"'{clean_name}' 님의 등록된 정보가 없습니다. 신규 가입 화면으로 이동합니다.")
    st.session_state.current_user_name = clean_name
    st.session_state.registration_role = role
    st.session_state.login_mode = 'new_user_registration'
    st.rerun()

# --- 3. 로그인 및 등록 UI ---
def show_login_and_registration():
    if st.session_state.login_mode == 'not_logged_in':
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

        st.subheader("시스템 로그인")
        tab_student, tab_doctor = st.tabs(["🎓 학생", "👨‍⚕️ 치과의사"])
        
        with tab_student:
            s_name = st.text_input("학생 이름", key="s_name_input")
            s_pw = st.text_input("비밀번호", type="password", key="s_pw_input")
            # 경고 메시지 해결: use_container_width=True 대신 width='stretch' 사용
            if st.button("학생 로그인", width="stretch"):
                _handle_login(s_name, s_pw, role="student")
                
        with tab_doctor:
            d_name = st.text_input("치과의사 이름", key="d_name_input")
            d_pw = st.text_input("비밀번호", type="password", key="d_pw_input")
            if st.button("치과의사 로그인", width="stretch"):
                _handle_login(d_name, d_pw, role="doctor")
    
    elif st.session_state.login_mode == 'new_user_registration':
        role = st.session_state.get('registration_role', 'student')
        role_kr = "치과의사" if role == 'doctor' else "학생"
        
        st.subheader(f"✨ 신규 {role_kr} 계정 등록")
        st.info(f"환영합니다, **{st.session_state.current_user_name}**님! 사용할 계정 정보를 입력해주세요.")
        
        email = st.text_input("이메일 주소 (ID 및 알림 수신용)")
        pw = st.text_input("사용할 비밀번호", type="password")
        
        dept = None
        if role == 'doctor':
            dept = st.selectbox("소속 진료과", DEPARTMENTS_FOR_REGISTRATION)
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("등록 완료", width="stretch"):
                if not is_valid_email(email):
                    st.error("올바른 이메일 형식을 입력해주세요.")
                elif not pw:
                    st.error("비밀번호를 입력해주세요.")
                else:
                    safe_key = sanitize_path(email)
                    user_data = {
                        "name": st.session_state.current_user_name, 
                        "email": email, 
                        "password": hash_password(pw)
                    }
                    if role == 'doctor':
                        user_data["department"] = dept
                        doctor_users_ref.child(safe_key).set(user_data)
                    else:
                        users_ref.child(safe_key).set(user_data)
                        
                    st.session_state.update({
                        'current_firebase_key': safe_key, 
                        'login_mode': 'doctor_mode' if role == 'doctor' else 'user_mode'
                    })
                    st.success("성공적으로 등록되었습니다!")
                    st.rerun()
        with col2:
            if st.button("취소 (돌아가기)", width="stretch"):
                st.session_state.login_mode = 'not_logged_in'
                st.rerun()

# --- 4. 관리자 모드 UI ---
def show_admin_mode_ui():
    st.title("🛡️ 관리자 대시보드")
    if st.button("← 로그아웃"):
        st.session_state.login_mode = 'not_logged_in'
        st.rerun()

    tab_auto, tab_manage = st.tabs(["🚀 자동 알림 전송", "👥 사용자 관리"])

    # [탭 1: 자동 알림 전송]
    with tab_auto:
        st.subheader("OCS 엑셀 데이터 매칭 및 전송")
        
        # 🚨 추가 완료 🚨: 엑셀 파일 비밀번호를 입력받습니다.
        excel_pw = st.text_input("🔒 엑셀 파일 암호 (필수)", type="password", placeholder="파일 암호를 입력하세요")
        
        uploaded_file = st.file_uploader("OCS 엑셀 파일을 업로드하세요", type=["xlsx", "xlsm"])
        if uploaded_file:
            if not excel_pw:
                st.warning("⚠️ 엑셀 파일의 암호를 먼저 입력해주세요.")
            else:
                is_daily = excel_utils.is_daily_schedule(uploaded_file.name)
                try:
                    # 🚨 수정 완료 🚨: excel_pw를 파라미터로 넘깁니다.
                    xl_data, styled_file = excel_utils.process_excel_file_and_style(
                        uploaded_file, 
                        db_ref_func, 
                        excel_password=excel_pw
                    )
                    st.success(f"파일 분석 완료: {uploaded_file.name}")
                    
                    if st.button("🚀 전체 자동 알림 전송 시작", type="primary"):
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

    # [탭 2: 사용자 관리]
    with tab_manage:
        st.subheader("등록된 사용자 전체 목록")
        
        all_students = users_ref.get() or {}
        all_doctors = doctor_users_ref.get() or {}
        
        user_list = []
        user_map = {}
        table_data = []

        for safe_key, info in all_students.items():
            name = info.get('name', '이름없음')
            email = info.get('email', '이메일없음')
            display_str = f"🎓 [학생] {name} ({email})"
            user_list.append(display_str)
            user_map[display_str] = {'safe_key': safe_key, 'role': 'student', 'email': email, 'name': name}
            table_data.append({"역할": "🎓 학생", "이름": name, "이메일": email, "진료과": "-"})

        for safe_key, info in all_doctors.items():
            name = info.get('name', '이름없음')
            email = info.get('email', '이메일없음')
            dept = info.get('department', '미지정')
            display_str = f"👨‍⚕️ [의사] {name} ({email})"
            user_list.append(display_str)
            user_map[display_str] = {'safe_key': safe_key, 'role': 'doctor', 'email': email, 'name': name}
            table_data.append({"역할": "👨‍⚕️ 의사", "이름": name, "이메일": email, "진료과": dept})

        if table_data:
            st.dataframe(pd.DataFrame(table_data), width="stretch")
            
            st.markdown("---")
            st.markdown("#### ⚙️ 일괄 작업 수행 (삭제 / 메일 발송)")
            
            select_all = st.checkbox("목록 모두 선택하기")
            if select_all:
                selected_users = st.multiselect("작업할 사용자 선택", options=user_list, default=user_list)
            else:
                selected_users = st.multiselect("작업할 사용자 선택", options=user_list)

            col1, col2 = st.columns([1, 1])
            
            with col1:
                st.markdown("**🗑️ 계정 삭제**")
                if st.button("선택한 사용자 영구 삭제"):
                    if not selected_users:
                        st.warning("삭제할 사용자를 선택해주세요.")
                    else:
                        for u_disp in selected_users:
                            meta = user_map[u_disp]
                            if meta['role'] == 'student':
                                users_ref.child(meta['safe_key']).delete()
                                db_ref_func(f"patients/{meta['safe_key']}").delete() 
                            else:
                                doctor_users_ref.child(meta['safe_key']).delete()
                        st.success(f"{len(selected_users)}명의 사용자가 삭제되었습니다.")
                        st.rerun()

            with col2:
                st.markdown("**📧 단체 메일 발송**")
                mail_subject = st.text_input("메일 제목", placeholder="예: 시스템 공지사항")
                mail_body = st.text_area("메일 내용 (HTML 태그 사용 가능)")
                if st.button("선택한 사용자에게 메일 발송"):
                    if not selected_users:
                        st.warning("메일을 보낼 사용자를 선택해주세요.")
                    elif not mail_body:
                        st.warning("메일 내용을 입력해주세요.")
                    else:
                        try:
                            sender = st.secrets["gmail"]["sender"]
                            sender_pw = st.secrets["gmail"]["app_password"]
                            success_count = 0
                            
                            final_body = f"<h3>{mail_subject}</h3><br>{mail_body}"
                            
                            with st.spinner("메일 발송 중..."):
                                for u_disp in selected_users:
                                    target_email = user_map[u_disp]['email']
                                    if target_email and is_valid_email(target_email):
                                        res = send_email(target_email, None, sender, sender_pw, custom_message=final_body)
                                        if res is True: success_count += 1
                                        
                            st.success(f"총 {success_count}명에게 메일이 발송되었습니다!")
                        except Exception as e:
                            st.error(f"메일 발송 오류: Secrets 설정(gmail)을 확인하세요. ({e})")
        else:
            st.info("현재 등록된 사용자가 없습니다.")

# --- 5. 일반 사용자(학생) 모드 UI ---
def show_user_mode_ui(firebase_key, user_name):
    patients_ref = db_ref_func(f"patients/{firebase_key}")
    
    st.subheader(f"🎓 {user_name} 학생님")
    
    service = get_google_calendar_service(firebase_key)
    if not service:
        raw_creds = db_ref_func(f"google_calendar_creds/{firebase_key}").get()
        if raw_creds:
            st.warning("💡 **안내:** 과거 연동 기록이 있으나, 자동 갱신 권한(refresh_token)이 누락되어 연결이 끊어졌습니다. **버튼을 눌러 1회 재연동 하시면 이후부터는 영구적으로 자동 갱신됩니다.**")

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
    
    service = get_google_calendar_service(firebase_key)
    if not service:
        raw_creds = db_ref_func(f"google_calendar_creds/{firebase_key}").get()
        if raw_creds:
            st.warning("💡 **안내:** 과거 연동 기록이 있으나, 자동 갱신 권한(refresh_token)이 누락되어 연결이 끊어졌습니다. **버튼을 눌러 1회 재연동 하시면 이후부터는 영구적으로 자동 갱신됩니다.**")
    
    st.info("의사님께 배정된 환자 내원 정보는 시스템에 의해 자동으로 구글 캘린더에 동기화됩니다.")
    
    doc_info = doctor_users_ref.child(firebase_key).get()
    if doc_info:
        st.write(f"📧 연동 이메일: {doc_info.get('email', '미지정')}")
        st.write(f"🏥 소속 과: {doc_info.get('department', '미지정')}")

    st.markdown("---")
    if st.button("로그아웃"):
        st.session_state.login_mode = 'not_logged_in'
        st.rerun()
