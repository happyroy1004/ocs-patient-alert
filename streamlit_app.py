import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, db
import io
import msoffcrypto
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import re

# 🔐 Firebase 초기화
if not firebase_admin._apps:
    cred = credentials.Certificate(st.secrets["firebase_credentials"])
    firebase_admin.initialize_app(cred, {
        'databaseURL': st.secrets["database_url"]
    })

# 📌 Firebase-safe 경로 변환
def sanitize_path(s):
    return re.sub(r'[.$#[\]/]', '_', s)

# 🔒 암호화 확인
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
                raise ValueError("암호화된 파일입니다. 암호를 입력해주세요.")
            decrypted = io.BytesIO()
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            return pd.ExcelFile(decrypted)
        else:
            return pd.ExcelFile(file)
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

# ✅ Streamlit 시작
st.title("🩺 환자 내원 확인 시스템")
user_id = st.text_input("아이디를 입력하세요")
if not user_id:
    st.stop()

firebase_key = sanitize_path(user_id)

# 🔑 관리자 모드
if user_id == "admin":
    st.subheader("📂 엑셀 업로드 및 전체 사용자 비교")

    uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])
    if uploaded_file:
        password = None
        if is_encrypted_excel(uploaded_file):
            password = st.text_input("🔑 엑셀 파일 비밀번호 입력", type="password")
            if not password:
                st.stop()

        try:
            xl = load_excel(uploaded_file, password)
            sender = st.secrets["gmail"]["sender"]
            sender_pw = st.secrets["gmail"]["app_password"]

            matched_users = []

            users_ref = db.reference("patients")
            all_users = users_ref.get()

            if not all_users:
                st.warning("❗ Firebase에 등록된 사용자가 없습니다.")
                st.stop()

            # 엑셀 파싱 + 사용자별 비교
            for uid, plist in all_users.items():
                registered_set = set((v["환자명"], v["진료번호"]) for v in plist.values())
                matched_rows = []

                for sheet_name in xl.sheet_names:
                    try:
                        df = xl.parse(sheet_name, header=1).astype(str)
                        if "환자명" not in df.columns or "진료번호" not in df.columns:
                            continue
                        df["환자명"] = df["환자명"].str.strip()
                        df["진료번호"] = df["진료번호"].str.strip()
                        matched = df[df.apply(lambda row: (row["환자명"], row["진료번호"]) in registered_set, axis=1)]
                        if not matched.empty:
                            matched["시트명"] = sheet_name
                            matched_rows.append(matched)
                    except Exception:
                        continue

                if matched_rows:
                    result_df = pd.concat(matched_rows, ignore_index=True)
                    matched_users.append((uid, result_df))

            if matched_users:
                send = st.radio("✉️ 이메일을 전송하시겠습니까?", ["예", "아니오"])
                if send == "예":
                    for uid, df_matched in matched_users:
                        recipient_email = uid
                        result = send_email(recipient_email, df_matched, sender, sender_pw)
                        if result == True:
                            st.success(f"✅ {recipient_email} 전송 완료")
                        else:
                            st.error(f"❌ {recipient_email} 전송 실패: {result}")
                else:
                    for uid, df in matched_users:
                        st.markdown(f"### 📧 {uid}")
                        st.dataframe(df)
            else:
                st.info("📭 매칭된 사용자 없음")

        except Exception as e:
            st.error(f"❌ 파일 처리 실패: {e}")

# 👥 일반 사용자
else:
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
