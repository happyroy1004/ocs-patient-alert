import streamlit as st
import pandas as pd
import msoffcrypto
import io
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import firebase_admin
from firebase_admin import credentials, firestore

# 초기화: firebase-admin
if not firebase_admin._apps:
    cred = credentials.Certificate({
        "type": st.secrets["firebase"]["type"],
        "project_id": st.secrets["firebase"]["project_id"],
        "private_key_id": st.secrets["firebase"]["private_key_id"],
        "private_key": st.secrets["firebase"]["private_key"].replace('\\n', '\n'),
        "client_email": st.secrets["firebase"]["client_email"],
        "client_id": st.secrets["firebase"]["client_id"],
        "auth_uri": st.secrets["firebase"]["auth_uri"],
        "token_uri": st.secrets["firebase"]["token_uri"],
        "auth_provider_x509_cert_url": st.secrets["firebase"]["auth_provider_x509_cert_url"],
        "client_x509_cert_url": st.secrets["firebase"]["client_x509_cert_url"]
    })
    firebase_admin.initialize_app(cred)

db = firestore.client()

st.title("📂 관리자 Excel 업로드 및 환자 이메일 알림")

user_id = st.text_input("아이디를 입력하세요")

if user_id == "admin":
    st.success("🔐 관리자 모드입니다.")

    uploaded_file = st.file_uploader("🔒 암호화된 Excel 파일을 업로드하세요", type=["xlsx"])

    if uploaded_file:
        password = st.text_input("Excel 파일의 암호를 입력하세요", type="password")

        if password:
            try:
                decrypted = io.BytesIO()
                file = msoffcrypto.OfficeFile(uploaded_file)
                file.load_key(password=password)
                file.decrypt(decrypted)

                df = pd.read_excel(decrypted, engine="openpyxl")
                st.success("✅ 파일이 성공적으로 복호화되었습니다.")
                st.dataframe(df)

                # Excel 다운로드
                processed_file = io.BytesIO()
                df.to_excel(processed_file, index=False, engine='openpyxl')
                processed_file.seek(0)
                st.download_button("📥 처리된 파일 다운로드", processed_file, file_name="processed.xlsx")

                # 메일 발송 여부
                send_email = st.radio("📧 사용자에게 환자 내원 이메일을 보내시겠습니까?", ["예", "아니오"])

                if send_email == "예":
                    sender_email = st.secrets["gmail"]["sender"]
                    sender_pw = st.secrets["gmail"]["app_password"]

                    users_ref = db.collection("users")
                    docs = users_ref.stream()

                    for doc in docs:
                        user_email = doc.id
                        patient_list = doc.to_dict().get("patients", [])
                        matched = []

                        for entry in patient_list:
                            name = entry.get("name")
                            number = str(entry.get("number"))
                            if ((df['이름'] == name) & (df['환자번호'].astype(str) == number)).any():
                                matched.append(f"{name} ({number})")

                        if matched:
                            try:
                                msg = MIMEMultipart()
                                msg['From'] = sender_email
                                msg['To'] = user_email
                                msg['Subject'] = "📌 등록 환자 내원 알림"
                                body = "다음 환자가 방문했습니다:\n" + "\n".join(matched)
                                msg.attach(MIMEText(body, 'plain'))

                                server = smtplib.SMTP('smtp.gmail.com', 587)
                                server.starttls()
                                server.login(sender_email, sender_pw)
                                server.send_message(msg)
                                server.quit()

                                st.success(f"✅ {user_email}에게 메일 전송 완료")
                            except Exception as e:
                                st.error(f"❌ {user_email} 전송 실패: {e}")
                        else:
                            st.info(f"ℹ️ {user_email}: 등록 환자 중 내원자 없음")

            except Exception as e:
                st.error(f"❌ 복호화 또는 처리 실패: {e}")
else:
    st.info("관리자 계정으로만 접근 가능합니다.")
