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
from firebase_utils import load_google_creds_from_firebase, recover_email
from config import PATIENT_DEPT_FLAGS, PATIENT_DEPT_TO_SHEET_MAP, SHEET_KEYWORD_TO_DEPARTMENT_MAP

# --- 유효성 검사 ---
def is_valid_email(email):
    """이메일 주소 형식을 확인합니다."""
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(email_regex, email) is not None

# --- 이메일 전송 ---
def send_email(receiver, rows, sender, password, date_str=None, custom_message=None):
    """
    이메일을 전송하는 범용 함수입니다.
    custom_message가 있으면 그것을 본문으로 사용합니다 (표 + 텍스트).
    """
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver

        if custom_message:
            # run_auto_notifications에서 만들어준 완벽한 본문(표+텍스트)을 그대로 사용
            msg['Subject'] = "단체 메일 알림" if date_str is None else f"[치과 내원 알림] {date_str} 예약 내역"
            body = custom_message
        else:
            # custom_message가 없을 때의 기본 동작 (단순 표만 전송)
            subject_prefix = ""
            if date_str:
                subject_prefix = f"{date_str}일에 내원하는 "
            msg['Subject'] = f"{subject_prefix}등록 환자 내원 알림"
            
            if rows is not None and isinstance(rows, list):
                rows_df = pd.DataFrame(rows)
                html_table = rows_df.to_html(index=False, escape=False)
                style = """<style>table {width: 100%; border-collapse: collapse;} th, td {border: 1px solid #ddd; padding: 8px;}</style>"""
                body = f"다음 환자가 내일 내원예정입니다:<br><br>{style}{html_table}"
            else:
                 body = "내원 환자 정보가 없습니다."

        msg.attach(MIMEText(body, 'html'))
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        return str(e)

# --- Google Calendar 이벤트 생성 ---
def create_calendar_event(service, patient_name, pid, department, reservation_datetime, doctor_name, treatment_details, is_daily):
    """
    Google Calendar에 단일 이벤트를 생성합니다.
    """
    seoul_tz = datetime.timezone(datetime.timedelta(hours=9))
    event_start = reservation_datetime.replace(tzinfo=seoul_tz)
    event_end = event_start + datetime.timedelta(minutes=30)
    
    # 1. 제목 포맷팅: friendly_time 제거, 단순 시간(HHMM)만 표시
    time_hhmm = event_start.strftime("%H%M")
    
    event_prefix = "✨ 내원 : " if is_daily else "❓내원 : "
    # 예: ✨ 내원 : 0900 홍길동 (보존과, 김의사)
    summary_text = f'{event_prefix}{time_hhmm} {patient_name} ({department}, {doctor_name})' 
    
    # 2. 설명(description) 포맷팅: 맨 윗줄에 데이터 헤더 추가
    # 형식: 진료의사,날짜(MMDD),시간(HHMM),환자이름,환자번호,
    date_mmdd = event_start.strftime("%m%d")
    header_info = f"{doctor_name},{date_mmdd},{time_hhmm},{patient_name},{pid},"
    
    description_text = f"{header_info}\n\n환자명 : {patient_name}\n진료번호 : {pid}\n진료내역 : {treatment_details}"

    event = {
        'summary': summary_text,
        'location': pid,
        'description': description_text,
        'start': {
            'dateTime': event_start.isoformat(),
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': event_end.isoformat(),
            'timeZone': 'Asia/Seoul',
        },
    }

    try:
        service.events().insert(calendarId='primary', body=event).execute()
        return True
    except HttpError as error:
        st.error(f"캘린더 이벤트 생성 중 오류 발생: {error}")
        return False
    except Exception as e:
        st.error(f"알 수 없는 오류 발생: {e}")
        return False
        
# --- 매칭 로직 ---

def standardize_df_for_matching(df):
    """Excel DataFrame의 핵심 컬럼을 매칭을 위해 표준화합니다."""
    df = df.copy()
    
    df.columns = [str(col).strip() for col in df.columns]
    current_cols = df.columns
    
    required_cols = ['진료번호', '환자명', '예약의사']
    
    if not all(col in current_cols for col in required_cols) and not df.empty and len(df) > 0:
        new_header = df.iloc[0] 
        new_header = [str(h).strip() for h in new_header] 
        
        if all(isinstance(col, (int)) for col in current_cols):
             rename_map = {0: '예약일시', 1: '예약시간', 2: '진료번호', 3: '환자명', 5: '예약의사'}
             df.rename(columns=rename_map, inplace=True)
             df.columns = [str(col).strip() for col in df.columns]

        df.columns = new_header
        df = df[1:].reset_index(drop=True)
        current_cols = df.columns
        df.columns = [str(col).strip() for col in df.columns]
    
    if not all(col in current_cols for col in required_cols):
         return pd.DataFrame(columns=required_cols) 

    df = df.fillna("").astype(str)
    df['진료번호'] = df['진료번호'].str.strip().str.zfill(8)
    df['환자명'] = df['환자명'].str.strip()
        
    if '예약의사' in df.columns:
        df['예약의사'] = df['예약의사'].str.strip().str.replace(" 교수님", "", regex=False)
        df['예약의사'] = df['예약의사'].str.replace("'", "", regex=False).str.replace("‘", "", regex=False).str.replace("’", "", regex=False).str.strip()

    df = df[df['진료번호'] != '']
    final_cols = list(set(df.columns) & set(['예약일시', '예약시간', '진료번호', '환자명', '예약의사', '진료내역', '등록과']))
    return df[[col for col in final_cols if col in df.columns]].reset_index(drop=True)


def get_matching_data(excel_data_dfs, all_users_meta, all_patients_data, all_doctors_meta):
    """Excel 데이터와 Firebase 사용자/환자/의사 데이터를 매칭합니다."""
    
    matched_users = []; matched_doctors_data = []

    standardized_dfs = {
        sheet_name: standardize_df_for_matching(df)
        for sheet_name, df in excel_data_dfs.items()
    }

    # 1. 학생 매칭
    if all_patients_data:
        for uid_safe, registered_patients_for_this_user in all_patients_data.items():
            user_email = recover_email(uid_safe); user_display_name = user_email
            if all_users_meta and uid_safe in all_users_meta and "name" in all_users_meta[uid_safe]:
                user_display_name = all_users_meta[uid_safe]["name"]; user_email = all_users_meta[uid_safe]["email"]
            
            registered_patients_data = []
            if registered_patients_for_this_user:
                for pid_key, val in registered_patients_for_this_user.items(): 
                    registered_depts = [
                        dept.capitalize() for dept in PATIENT_DEPT_FLAGS + ['치주'] 
                        if val.get(dept.lower()) is True or val.get(dept.lower()) == 'True' or val.get(dept.lower()) == 'true'
                    ]
                    registered_patients_data.append({"환자명": val.get("환자이름", "").strip(), "진료번호": pid_key.strip().zfill(8), "등록과_리스트": registered_depts})
            
            matched_rows_for_user = []
            for registered_patient in registered_patients_data:
                registered_depts = registered_patient["등록과_리스트"]; sheets_to_search = set()
                for dept in registered_depts: sheets_to_search.update(PATIENT_DEPT_TO_SHEET_MAP.get(dept, [dept]))

                for sheet_name_excel_raw, df_sheet in standardized_dfs.items(): 
                    excel_sheet_department = None
                    for keyword, department_name in SHEET_KEYWORD_TO_DEPARTMENT_MAP.items():
                        if keyword.lower() in sheet_name_excel_raw.strip().lower(): excel_sheet_department = department_name; break
                    
                    if excel_sheet_department in sheets_to_search:
                        for _, excel_row in df_sheet.iterrows():
                            if (registered_patient["환자명"] == excel_row.get("환자명", "") and registered_patient["진료번호"] == excel_row.get("진료번호", "")):
                                matched_row_copy = excel_row.copy(); matched_row_copy["시트"] = sheet_name_excel_raw
                                matched_row_copy["등록과"] = ", ".join(registered_depts); matched_rows_for_user.append(matched_row_copy); break
            
            if matched_rows_for_user:
                combined_matched_df = pd.DataFrame(matched_rows_for_user)
                matched_users.append({"email": user_email, "name": user_display_name, "data": combined_matched_df, "safe_key": uid_safe})

    # 2. 치과의사 매칭
    doctors = []
    if all_doctors_meta:
        for safe_key, user_info in all_doctors_meta.items():
            if user_info: doctors.append({"safe_key": safe_key, "name": user_info.get("name", "이름 없음"), "email": user_info.get("email", "이메일 없음"), "department": user_info.get("department", "미지정")})
    
    if doctors and standardized_dfs:
        for res in doctors:
            doctor_dept = res['department']; sheets_to_search = PATIENT_DEPT_TO_SHEET_MAP.get(doctor_dept, [doctor_dept])
            matched_rows_for_doctor = [] 
            
            for sheet_name_excel_raw, df_sheet in standardized_dfs.items(): 
                excel_sheet_department = None
                for keyword, department_name in SHEET_KEYWORD_TO_DEPARTMENT_MAP.items():
                    if keyword.lower() in sheet_name_excel_raw.strip().lower(): excel_sheet_department = department_name; break
                
                if excel_sheet_department in sheets_to_search:
                    for _, excel_row in df_sheet.iterrows():
                        if excel_row.get('예약의사', '') == res['name']:
                            matched_rows_for_doctor.append(excel_row.copy())
            
            if matched_rows_for_doctor:
                 res['data'] = pd.DataFrame(matched_rows_for_doctor) 
                 matched_doctors_data.append(res)
                 
    return matched_users, matched_doctors_data

# --- 자동 알림 실행 ---
def run_auto_notifications(matched_users, matched_doctors, excel_data_dfs, file_name, is_daily, db_ref):
    """자동으로 모든 매칭 사용자에게 메일 및 캘린더 일정을 전송하는 핵심 로직"""
    sender = st.secrets["gmail"]["sender"]; sender_pw = st.secrets["gmail"]["app_password"]
    
    # --- 텍스트 생성 헬퍼 함수 ---
    def generate_email_body_with_text(user_name, df_matched, file_name):
        # 1. HTML Table 생성
        email_cols = ['환자명', '진료번호', '예약의사', '진료내역', '예약일시', '예약시간', '등록과']
        df_for_mail = df_matched[[col for col in email_cols if col in df_matched.columns]]
        
        # 스타일이 적용된 HTML Table
        table_style = """
        <style>
        table {width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 14px;}
        th, td {border: 1px solid #dddddd; text-align: left; padding: 8px;}
        th {background-color: #f2f2f2; font-weight: bold;}
        </style>
        """
        html_table = df_for_mail.to_html(index=False, escape=False)
        
        # 2. 텍스트 데이터 생성 (요청한 포맷)
        text_lines = []
        for _, row in df_matched.iterrows():
            try:
                # 날짜/시간 정제
                r_date = str(row.get('예약일시', '')).strip().replace('-', '/').replace('.', '/')
                r_time = str(row.get('예약시간', '')).strip()
                
                mmdd = "0000"; hhmm = "0000"
                if r_date and r_time:
                    try:
                        # 날짜 파싱 시도
                        full_datetime_str = f"{r_date} {r_time}"
                        dt = datetime.datetime.strptime(full_datetime_str, '%Y/%m/%d %H:%M')
                        mmdd = dt.strftime("%m%d")
                        hhmm = dt.strftime("%H%M")
                    except ValueError:
                        # 파싱 실패시 날짜에서 숫자만 추출 시도 (예: 2024/11/28 -> 1128)
                        parts = r_date.split('/')
                        if len(parts) >= 3: mmdd = f"{parts[1]}{parts[2]}"
                        hhmm = r_time.replace(':', '')

                # 포맷: 진료의사,날짜,시간,환자이름,환자번호,
                line = f"{row.get('예약의사','')},{mmdd},{hhmm},{row.get('환자명','')},{row.get('진료번호','')},"
                text_lines.append(line)
            except: continue
            
        formatted_text_html = "<br>".join(text_lines)
        
        # 3. 최종 본문 결합 (Table + Text)
        full_body = f"""
        <p>안녕하세요, {user_name}님.</p>
        <p>{file_name} 분석 결과, 내원 예정인 환자 진료 정보입니다.</p>
        <div class='table-container'>{table_style}{html_table}</div>
        <br><br>
        <hr>
        <p><strong>[복사 붙여넣기용 데이터]</strong></p>
        <div style='background-color: #f4f4f4; padding: 10px; border: 1px solid #ccc; font-family: monospace; font-size: 14px; line-height: 1.6;'>
        {formatted_text_html}
        </div>
        <p>확인 부탁드립니다.</p>
        """
        return full_body, df_for_mail.to_dict('records')

    # 1. 학생(일반 사용자) 자동 전송
    st.markdown("### 📚 학생(일반 사용자) 자동 전송 결과")
    if matched_users:
        for user_match_info in matched_users:
            real_email = user_match_info['email']; df_matched = user_match_info['data']
            user_name = user_match_info['name']; user_safe_key = user_match_info['safe_key']
            
            # 본문 생성 (Table + Text 포함)
            email_body, rows_as_dict = generate_email_body_with_text(user_name, df_matched, file_name)
            
            try:
                # custom_message에 email_body를 전달하여 send_email이 이것을 그대로 쓰게 함
                send_email(receiver=real_email, rows=rows_as_dict, sender=sender, password=sender_pw, custom_message=email_body, date_str=file_name) 
                st.write(f"✔️ **메일:** {user_name} ({real_email})에게 전송 완료.")
            except Exception as e: st.error(f"❌ **메일:** {user_name} ({real_email})에게 전송 실패: {e}")

            # 캘린더 등록
            creds = load_google_creds_from_firebase(user_safe_key)
            if creds and creds.valid and not creds.expired:
                try:
                    service = build('calendar', 'v3', credentials=creds)
                    for _, row in df_matched.iterrows():
                        reservation_date_raw = str(row.get('예약일시', '')).strip().replace('-', '/').replace('.', '/')
                        reservation_time_raw = str(row.get('예약시간', '')).strip()
                        if reservation_date_raw and reservation_time_raw:
                            try:
                                full_datetime_str = f"{reservation_date_raw} {reservation_time_raw}"
                                reservation_datetime = datetime.datetime.strptime(full_datetime_str, '%Y/%m/%d %H:%M')
                                create_calendar_event(
                                    service, row.get('환자명', 'N/A'), row.get('진료번호', ''), row.get('등록과', ''), 
                                    reservation_datetime, row.get('예약의사', ''), row.get('진료내역', ''), is_daily
                                )
                            except: pass
                    st.write(f"✔️ **캘린더:** {user_name}에게 일정 추가 완료.")
                except Exception as e: st.warning(f"⚠️ **캘린더:** {user_name} 일정 추가 중 오류: {e}")
            else: st.warning(f"⚠️ **캘린더:** {user_name}님은 Google Calendar 계정이 연동되지 않았습니다.")
    else: st.info("매칭된 학생(사용자)이 없습니다.")

    # 2. 치과의사 자동 전송
    st.markdown("### 🧑‍⚕️ 치과의사 자동 전송 결과")
    if matched_doctors:
        for res in matched_doctors:
            df_matched = res['data']
            
            # 본문 생성 (Table + Text 포함)
            email_body, rows_as_dict = generate_email_body_with_text(res['name'], df_matched, file_name)
            
            try:
                send_email(receiver=res['email'], rows=rows_as_dict, sender=sender, password=sender_pw, custom_message=email_body, date_str=file_name)
                st.write(f"✔️ **메일:** Dr. {res['name']}에게 전송 완료!")
            except Exception as e: st.error(f"❌ **메일:** Dr. {res['name']}에게 전송 실패: {e}")

            creds = load_google_creds_from_firebase(res['safe_key'])
            if creds and creds.valid and not creds.expired:
                try:
                    service = build('calendar', 'v3', credentials=creds)
                    for _, row in df_matched.iterrows():
                        reservation_date_raw = str(row.get('예약일시', '')).strip().replace('-', '/').replace('.', '/')
                        reservation_time_raw = str(row.get('예약시간', '')).strip()
                        if reservation_date_raw and reservation_time_raw:
                            try:
                                full_datetime_str = f"{reservation_date_raw} {reservation_time_raw}"
                                reservation_datetime = datetime.datetime.strptime(full_datetime_str, '%Y/%m/%d %H:%M')
                                create_calendar_event(
                                    service, row.get('환자명', 'N/A'), row.get('진료번호', ''), res.get('department', 'N/A'), 
                                    reservation_datetime, row.get('예약의사', ''), row.get('진료내역', ''), is_daily
                                )
                            except: pass
                    st.write(f"✔️ **캘린더:** Dr. {res['name']}에게 일정 추가 완료.")
                except Exception as e: st.warning(f"⚠️ **캘린더:** Dr. {res['name']} 일정 추가 중 오류: {e}")
            else: st.warning(f"⚠️ **캘린더:** Dr. {res['name']}님은 Google Calendar 계정이 연동되지 않았습니다.")
    else: st.info("매칭된 치과의사 계정이 없습니다.")
