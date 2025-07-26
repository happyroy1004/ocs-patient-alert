# streamlit_app.py
import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, db
import io
import msoffcrypto
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from openpyxl import Workbook

# 🔐 Firebase 초기화
if not firebase_admin._apps:
    cred = credentials.Certificate(st.secrets["firebase_credentials"])
    firebase_admin.initialize_app(cred, {
        'databaseURL': st.secrets["database_url"]
    })

# 📌 Firebase-safe 경로로 변환
def sanitize_path(s):
    import re
    return re.sub(r'[.$#[\]/]', '_', s)

# 🧾 암호화 여부 확인
def is_encrypted_excel(file):
    try:
        file.seek(0)
        office_file = msoffcrypto.OfficeFile(file)
        return office_file.is_encrypted()
    except Exception:
        return False

# 🧾 엑셀 파일 로드 함수
def load_excel(file, password=None):
    try:
        file.seek(0)
        office_file = msoffcrypto.OfficeFile(file)
        if office_file.is_encrypted():
            if not password:
                raise ValueError("암호화된 파일입니다. 암호를 입력해주세요.")
            decrypted = io.BytesIO()
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            return pd.ExcelFile(decrypted)
        else:
            file.seek(0)
            return pd.ExcelFile(file)
    except Exception as e:
        raise ValueError(f"엑셀 처리 실패: {e}")

# 📤 Gmail 발송 함수
def send_email(to_email, subject, body):
    gmail_user = st.secrets["email"]["gmail_user"]
    gmail_pass = st.secrets["email"]["gmail_pass"]

    msg = MIMEMultipart()
    msg['From'] = gmail_user
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(gmail_user, gmail_pass)
        server.send_message(msg)

# 엑셀을 바이너리로 변환
def convert_df_to_excel_bytes(df):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    for r_idx, row in enumerate(df.values.tolist(), start=2):
        for c_idx, val in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=val)
    for c_idx, col in enumerate(df.columns, start=1):
        ws.cell(row=1, column=c_idx, value=col)
    wb.save(output)
    return output.getvalue()

# 등록 환자 전체 매칭 및 사용자별 메일 발송
def match_and_email_to_users(sheet_dict):
    root_ref = db.reference("patients")
    all_users = root_ref.get()
    count = 0
    for user_id, patients in all_users.items():
        target_set = set((p['환자명'], p['진료번호']) for p in patients.values())
        matched_rows = []
        for sheet_name, df in sheet_dict.items():
            temp_df = df.astype(str)
            matched = temp_df[temp_df.apply(lambda row: (row.get("환자명"), row.get("진료번호")) in target_set, axis=1)]
            if not matched.empty:
                matched_rows.append(matched)
        if matched_rows:
            final_df = pd.concat(matched_rows)
            # Firebase에서 ID = 이메일로 가정
            send_email(user_id, "[내원 안내] 등록된 환자가 발견되었습니다.", final_df.to_string(index=False))
            count += 1
    return count

# 🔓 앱 실행 시작
st.title("🔐 토탈환자 내원 확인 시스템")

# 로그인 ID 입력
google_id = st.text_input("로그인 ID를 입력하세요")
if not google_id:
    st.stop()

is_admin = google_id.strip().lower() == "admin"
firebase_key = sanitize_path(google_id)
ref = db.reference(f"patients/{firebase_key}")
existing_data = ref.get()

# 일반 사용자 인터페이스
if not is_admin:
    st.subheader("📄 등록된 환자 목록")
    if existing_data:
        for key, val in existing_data.items():
            name = val.get("환자명", "없음")
            number = val.get("진료번호", "없음")
            col1, col2 = st.columns([0.85, 0.15])
            with col1:
                st.markdown(f"**👤 {name} | 🆔 {number}**")
            with col2:
                if st.button("삭제", key=f"del_{key}"):
                    db.reference(f"patients/{firebase_key}/{key}").delete()
                    st.success("삭제 완료")
                    st.rerun()
    else:
        st.info("등록된 환자가 없습니다.")

    with st.form("register_form"):
        st.subheader("➕ 환자 등록")
        name = st.text_input("환자명")
        number = st.text_input("진료번호")
        submitted = st.form_submit_button("등록")
        if submitted:
            if not name or not number:
                st.warning("모든 항목을 입력해주세요.")
            elif existing_data and any(v['환자명'] == name and v['진료번호'] == number for v in existing_data.values()):
                st.error("이미 등록된 환자입니다.")
            else:
                ref.push().set({"환자명": name, "진료번호": number})
                st.success("등록 완료")
                st.rerun()

# 관리자 전용 인터페이스
else:
    st.subheader("📂 엑셀 업로드 및 분석")
    uploaded_file = st.file_uploader("엑셀 파일(.xlsx/.xlsm)", type=["xlsx", "xlsm"])
    if uploaded_file:
        password = None
        if is_encrypted_excel(uploaded_file):
            password = st.text_input("🔑 암호 입력", type="password")
        try:
            xl = load_excel(uploaded_file, password)
            processed = {}
            for sheet in xl.sheet_names:
                df = xl.parse(sheet, header=1)
                if "환자명" in df.columns and "진료번호" in df.columns:
                    processed[sheet] = df
            st.success("엑셀 처리 완료")

            with st.expander("⬇️ 처리된 엑셀 다운로드"):
                for sheet, df in processed.items():
                    st.download_button(f"{sheet}.xlsx", convert_df_to_excel_bytes(df), f"{sheet}_processed.xlsx")

            if st.button("📧 등록 사용자에게 내원 안내 메일 발송"):
                total = match_and_email_to_users(processed)
                st.success(f"총 {total}명의 사용자에게 메일 발송 완료")

        except Exception as e:
            st.error(f"❌ 오류: {e}")
