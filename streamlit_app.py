import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, db
import io
import msoffcrypto
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from openpyxl import load_workbook
from openpyxl.styles import Font

# 🔐 Firebase 초기화
if not firebase_admin._apps:
    cred = credentials.Certificate(st.secrets["firebase_credentials"])
    firebase_admin.initialize_app(cred, {
        'databaseURL': st.secrets["firebase"]["database_url"]
    })

# 📌 Firebase-safe 경로 변환
def sanitize_path(email):
    return email.replace(".", "_dot_").replace("@", "_at_")

# 📩 이메일 주소 복원
def recover_email(safe_id: str) -> str:
    if safe_id.endswith("_com"):
        safe_id = safe_id[:-4] + ".com"
    return safe_id.replace("_at_", "@").replace("_dot_", ".")

# 🔒 암호화된 엑셀 여부 확인
def is_encrypted_excel(file):
    try:
        file.seek(0)
        return msoffcrypto.OfficeFile(file).is_encrypted()
    except Exception:
        return False

# 📂 엑셀 로드
def load_excel(file, password=None):
    try:
        file.seek(0)
        office_file = msoffcrypto.OfficeFile(file)
        if office_file.is_encrypted():
            if not password:
                raise ValueError("암호화된 파일입니다.")
            decrypted = io.BytesIO()
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            return pd.ExcelFile(decrypted), decrypted
        else:
            return pd.ExcelFile(file), file
    except Exception as e:
        raise ValueError(f"엑셀 처리 실패: {e}")

# 📧 이메일 전송
def send_email(receiver, rows, sender, password):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver
        msg['Subject'] = "📌 등록 환자 내원 알림"
        html_table = rows.to_html(index=False, escape=False)
        body = f"다음 등록 환자가 내원했습니다:<br><br>{html_table}"
        msg.attach(MIMEText(body, 'html'))

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        return str(e)

# 📑 엑셀 시트 파싱 및 정제
def process_excel(filelike):
    wb = load_workbook(filelike)
    processed = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        values = list(ws.values)
        while values and (values[0] is None or all(v is None for v in values[0])):
            values.pop(0)
        if len(values) < 2:
            continue
        df = pd.DataFrame(values)
        df.columns = df.iloc[0]
        df = df.drop([0]).reset_index(drop=True)
        df = df.fillna("").astype(str)
        df['환자명'] = df['환자명'].str.strip()
        df['진료번호'] = df['진료번호'].str.strip().str.zfill(8)
        processed[sheet_name] = df
    return processed

# 🌐 Streamlit 시작
st.title("🩺 환자 내원 확인 시스템")
user_id = st.text_input("아이디를 입력하세요")
if not user_id:
    st.stop()

firebase_key = sanitize_path(user_id)

# 👤 사용자 모드
if user_id != "admin":
    st.subheader("📝 내 환자 등록")
    ref = db.reference(f"patients/{firebase_key}")
    existing_data = ref.get()

    if existing_data:
        for key, val in existing_data.items():
            with st.container():
                col1, col2 = st.columns([0.85, 0.15])
                with col1:
                    st.markdown(f"👤 {val['환자명']} / 🆔 {val['진료번호']}")
                with col2:
                    if st.button("❌ 삭제", key=key):
                        db.reference(f"patients/{firebase_key}/{key}").delete()
                        st.success("삭제 완료")
                        st.rerun()
    else:
        st.info("등록된 환자가 없습니다.")

    with st.form("register_form"):
        name = st.text_input("환자명")
        pid = st.text_input("진료번호")
        submitted = st.form_submit_button("등록")
        if submitted:
            if not name or not pid:
                st.warning("모든 항목을 입력해주세요.")
            elif existing_data and any(
                v["환자명"] == name and v["진료번호"] == pid for v in existing_data.values()):
                st.error("이미 등록된 환자입니다.")
            else:
                ref.push().set({"환자명": name, "진료번호": pid})
                st.success(f"{name} ({pid}) 등록 완료")
                st.rerun()

# 🔑 관리자 모드
else:
    st.subheader("📂 엑셀 업로드 및 사용자 일치 검사")
    uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])
    if uploaded_file:
        password = None
        if is_encrypted_excel(uploaded_file):
            password = st.text_input("🔑 엑셀 파일 비밀번호 입력", type="password")
            if not password:
                st.stop()

        try:
            xl, raw_file = load_excel(uploaded_file, password)
            sender = st.secrets["gmail"]["sender"]
            sender_pw = st.secrets["gmail"]["app_password"]

            users_ref = db.reference("patients")
            all_users = users_ref.get()
            if not all_users:
                st.warning("❗ 등록된 사용자가 없습니다.")
                st.stop()

            excel_data = process_excel(raw_file)
            matched_users = []

            for uid, plist in all_users.items():
                registered_set = set(
                    (v["환자명"].strip(), v["진료번호"].strip().zfill(8)) for v in plist.values())
                matched_rows = []

                for sheet, df in excel_data.items():
                    matched = df[df.apply(lambda row: (row["환자명"], row["진료번호"]) in registered_set, axis=1)]
                    if not matched.empty:
                        matched["시트"] = sheet
                        matched_rows.append(matched)

                if matched_rows:
                    combined = pd.concat(matched_rows, ignore_index=True)
                    matched_users.append((uid, combined))

            if matched_users:
                st.success(f"🔍 {len(matched_users)}명의 사용자와 일치하는 환자 발견됨.")

                for uid, df in matched_users:
                    st.markdown(f"### 📧 {recover_email(uid)}")
                    st.dataframe(df)

                if st.button("📤 메일 보내기"):
                    for uid, df in matched_users:
                        real_email = recover_email(uid)
                        result = send_email(real_email, df, sender, sender_pw)
                        if result is True:
                            st.success(f"✅ {real_email} 전송 완료")
                        else:
                            st.error(f"❌ {real_email} 전송 실패: {result}")

                # 📥 엑셀 다운로드 버튼
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    for sheet, df in excel_data.items():
                        df.to_excel(writer, sheet_name=sheet, index=False)
                output_buffer.seek(0)
                output_filename = uploaded_file.name.replace(".xlsx", "_processed.xlsx").replace(".xlsm", "_processed.xlsx")
                st.download_button("📥 처리된 엑셀 다운로드", output_buffer, file_name=output_filename)
            else:
                st.info("📭 매칭된 사용자 없음")

        except Exception as e:
            st.error(f"❌ 파일 처리 실패: {e}")
