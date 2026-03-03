# notification_utils.py
import re
import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import datetime
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
# firebase_utils에서 통합된 로딩 및 복구 함수 임포트
from firebase_utils import load_google_creds_from_firebase, recover_email
from config import PATIENT_DEPT_FLAGS, PATIENT_DEPT_TO_SHEET_MAP, SHEET_KEYWORD_TO_DEPARTMENT_MAP

# --- 1. 유효성 검사 유틸리티 ---
def is_valid_email(email):
    """이메일 주소 형식을 확인합니다."""
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(email_regex, email) is not None

# --- 2. 이메일 전송 핵심 로직 ---
def send_email(receiver, rows, sender, password, date_str=None, custom_message=None):
    """이메일을 전송하는 범용 함수입니다."""
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver

        if custom_message:
            msg['Subject'] = "단체 메일 알림" if date_str is None else f"[치과 내원 알림] {date_str} 예약 내역"
            body = custom_message
        else:
            subject_prefix = f"{date_str}일에 내원하는 " if date_str else ""
            msg['Subject'] = f"{subject_prefix}등록 환자 내원 알림"
            
            if rows:
                rows_df = pd.DataFrame(rows)
                html_table = rows_df.to_html(index=False, escape=False)
                style = """
                <style>
                table { width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 14px; }
                th, td { border: 1px solid #dddddd; text-align: left; padding: 8px; }
                th { background-color: #f2f2f2; font-weight: bold; }
                </style>
                """
                body = f"다음 환자가 내일 내원예정입니다:<br><br>{style}{html_table}"
            else:
                 body = "내원 환자 정보가 없습니다."

        msg.attach(MIMEText(body, 'html'))
        
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender, password)
            server.send_message(msg)
        return True
    except Exception as e:
        return str(e)

# --- 3. Google Calendar 이벤트 생성 ---
def create_calendar_event(service, patient_name, pid, department, reservation_datetime, doctor_name, treatment_details, is_daily):
    """Google Calendar에 단일 이벤트를 생성합니다."""
    seoul_tz = datetime.timezone(datetime.timedelta(hours=9))
    event_start = reservation_datetime.replace(tzinfo=seoul_tz)
    event_end = event_start + datetime.timedelta(minutes=30)
    
    event_prefix = "✨ 내원 : " if is_daily else "❓내원 : "
    summary_text = f'{event_prefix}{patient_name} ({department}, {doctor_name})' 
    
    event = {
        'summary': summary_text,
        'location': pid,
        'description': f"환자명 : {patient_name}\n진료번호 : {pid}\n진료내역 : {treatment_details}",
        'start': {'dateTime': event_start.isoformat(), 'timeZone': 'Asia/Seoul'},
        'end': {'dateTime': event_end.isoformat(), 'timeZone': 'Asia/Seoul'},
    }

    try:
        service.events().insert(calendarId='primary', body=event).execute()
        return True
    except Exception as e:
        st.error(f"캘린더 이벤트 생성 오류: {e}")
        return False
        
# --- 4. 데이터 매칭 로직 ---

def standardize_df_for_matching(df):
    """Excel 데이터를 매칭을 위해 표준화합니다."""
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    
    required_cols = ['진료번호', '환자명', '예약의사']
    if not all(col in df.columns for col in required_cols):
         return pd.DataFrame(columns=required_cols) 

    df = df.fillna("").astype(str)
    df['진료번호'] = df['진료번호'].str.strip().str.zfill(8)
    df['환자명'] = df['환자명'].str.strip()
    if '예약의사' in df.columns:
        df['예약의사'] = df['예약의사'].str.replace(" 교수님", "", regex=False).str.strip()
    
    return df

def get_matching_data(excel_data_dfs, all_users_meta, all_patients_data, all_doctors_meta):
    """Excel 데이터와 Firebase 데이터를 매칭합니다."""
    matched_users = []; matched_doctors_data = []
    standardized_dfs = {k: standardize_df_for_matching(v) for k, v in excel_data_dfs.items()}

    # 1. 학생(일반 사용자) 매칭
    if all_patients_data:
        for uid_safe, patients in all_patients_data.items():
            user_name = all_users_meta.get(uid_safe, {}).get("name", recover_email(uid_safe))
            user_email = all_users_meta.get(uid_safe, {}).get("email", recover_email(uid_safe))
            
            rows = []
            if patients:
                for pid, val in patients.items():
                    target_pid = pid.zfill(8)
                    for _, df in standardized_dfs.items():
                        match = df[df['진료번호'] == target_pid]
                        if not match.empty:
                            row_copy = match.iloc[0].copy()
                            row_copy['등록과'] = "매칭됨"
                            rows.append(row_copy)
            
            if rows:
                matched_users.append({"email": user_email, "name": user_name, "data": pd.DataFrame(rows), "safe_key": uid_safe})

    # 2. 치과의사 매칭
    if all_doctors_meta:
        for safe_key, info in all_doctors_meta.items():
            doc_name = info.get("name")
            rows = []
            for _, df in standardized_dfs.items():
                doc_rows = df[df['예약의사'] == doc_name]
                if not doc_rows.empty:
                    rows.append(doc_rows)
            
            if rows:
                matched_doctors_data.append({
                    "safe_key": safe_key, 
                    "name": doc_name, 
                    "email": info.get("email"), 
                    "department": info.get("department", "N/A"),
                    "data": pd.concat(rows)
                })

    return matched_users, matched_doctors_data

# --- 5. 자동 알림 실행 (학생 및 의사 포함 전체 로직) ---
def run_auto_notifications(matched_users, matched_doctors, excel_data_dfs, file_name, is_daily, db_ref):
    """자동으로 학생과 의사에게 메일 및 캘린더 일정을 전송합니다."""
    sender = st.secrets["gmail"]["sender"]
    sender_pw = st.secrets["gmail"]["app_password"]
    
    # [A] 학생(일반 사용자) 전송
    st.markdown("### 📚 학생 자동 전송 결과")
    for user in matched_users:
        # 1. 메일 전송
        df_html = user['data'].to_html(index=False)
        email_body = f"<h4>{user['name']}님, 환자 내원 알림입니다.</h4>{df_html}"
        send_result = send_email(user['email'], None, sender, sender_pw, custom_message=email_body, date_str=file_name)
        
        if send_result is True:
            st.write(f"✔️ **메일:** {user['name']}님에게 전송 완료.")
        else:
            st.error(f"❌ **메일:** {user['name']}님 전송 실패: {send_result}")
        
        # 2. 캘린더 전송 (통합 경로 로딩 적용)
        creds = load_google_creds_from_firebase(user['safe_key'])
        if creds and creds.valid:
            try:
                service = build('calendar', 'v3', credentials=creds)
                for _, row in user['data'].iterrows():
                    dt_str = f"{row['예약일시']} {row['예약시간']}".strip()
                    dt_obj = datetime.datetime.strptime(dt_str, '%Y/%m/%d %H:%M')
                    create_calendar_event(service, row['환자명'], row['진료번호'], "치과", dt_obj, row['예약의사'], row.get('진료내역',''), is_daily)
                st.write(f"✔️ **캘린더:** {user['name']}님 일정 동기화 완료.")
            except Exception as e:
                st.warning(f"⚠️ **캘린더:** {user['name']}님 동기화 중 오류 발생.")
        else:
            st.warning(f"⚠️ **캘린더:** {user['name']}님 계정이 연동되지 않았습니다.")

    # [B] 치과의사 전송
    st.markdown("### 🧑‍⚕️ 치과의사 자동 전송 결과")
    for doc in matched_doctors:
        # 1. 메일 전송
        df_html = doc['data'].to_html(index=False)
        email_body = f"<h4>{doc['name']} 의사님, 예약 환자 알림입니다.</h4>{df_html}"
        send_result = send_email(doc['email'], None, sender, sender_pw, custom_message=email_body, date_str=file_name)
        
        if send_result is True:
            st.write(f"✔️ **메일:** Dr. {doc['name']}에게 전송 완료.")
        else:
            st.error(f"❌ **메일:** Dr. {doc['name']} 전송 실패: {send_result}")

        # 2. 캘린더 전송 (통합 경로 로딩 적용)
        creds = load_google_creds_from_firebase(doc['safe_key'])
        if creds and creds.valid:
            try:
                service = build('calendar', 'v3', credentials=creds)
                for _, row in doc['data'].iterrows():
                    dt_str = f"{row['예약일시']} {row['예약시간']}".strip()
                    dt_obj = datetime.datetime.strptime(dt_str, '%Y/%m/%d %H:%M')
                    create_calendar_event(service, row['환자명'], row['진료번호'], doc['department'], dt_obj, doc['name'], row.get('진료내역',''), is_daily)
                st.write(f"✔️ **캘린더:** Dr. {doc['name']} 일정 동기화 완료.")
            except Exception as e:
                st.error(f"❌ **캘린더:** Dr. {doc['name']} 동기화 실패.")
        else:
            st.warning(f"⚠️ **캘린더:** Dr. {doc['name']}님 계정이 연동되지 않았습니다.")
