import streamlit as st
import pandas as pd
import io
import base64
import msoffcrypto
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import firebase_admin
from firebase_admin import credentials, firestore

# Firebase 인증
cred = credentials.Certificate(st.secrets["firebase_credentials"])
firebase_admin.initialize_app(cred)
db = firestore.client()

# Gmail 정보
EMAIL_ADDRESS = st.secrets["gmail"]["address"]
EMAIL_PASSWORD = st.secrets["gmail"]["password"]

# 로그인
st.title("OCS Patient Alert System")
login_type = st.selectbox("로그인 유형을 선택하세요", ["사용자", "관리자"])
user_id = st.text_input("아이디를 입력하세요")

if user_id:
    user_ref = db.collection("users").document(user_id)

    # 1) 일반 사용자 로그인
    if login_type == "사용자":
        st.subheader(f"👩‍⚕️ {user_id}님 환자 목록")
        doc = user_ref.get()
        patient_list = doc.to_dict().get("patients", []) if doc.exists else []

        # 등록
        name = st.text_input("환자 이름")
        number = st.text_input("환자 진료번호")
        if st.button("등록"):
            new_patient = {"name": name, "number": number}
            if new_patient not in patient_list:
                patient_list.append(new_patient)
                user_ref.set({"patients": patient_list}, merge=True)
                st.success("✅ 환자 등록 완료")
            else:
                st.warning("⚠️ 이미 등록된 환자입니다.")

        # 삭제
        if patient_list:
            delete_index = st.selectbox("삭제할 환자 선택", range(len(patient_list)), format_func=lambda i: f"{patient_list[i]['name']} / {patient_list[i]['number']}")
            if st.button("삭제"):
                del patient_list[delete_index]
                user_ref.set({"patients": patient_list}, merge=True)
                st.success("🗑️ 삭제 완료")

        # 목록 표시
        st.write(pd.DataFrame(patient_list))

    # 2) 관리자 기능
    elif login_type == "관리자":
        st.subheader("📁 엑셀 업로드 및 복호화")
        uploaded_file = st.file_uploader("🔐 암호화된 엑셀 파일 업로드", type=["xls", "xlsx"])
        password = st.text_input("엑셀 암호 입력", type="password")

        if uploaded_file and password:
            decrypted = io.BytesIO()
            office_file = msoffcrypto.OfficeFile(uploaded_file)
            try:
                office_file.load_key(password=password)
                office_file.decrypt(decrypted)
                df = pd.read_excel(decrypted)

                st.success("🔓 복호화 성공!")
                st.dataframe(df)

                # 다운로드 버튼
                towrite = io.BytesIO()
                df.to_excel(towrite, index=False, engine='openpyxl')
                towrite.seek(0)
                b64 = base64.b64encode(towrite.read()).decode()
                href = f'<a href="data:application/octet-stream;base64,{b64}" download="처리된_엑셀.xlsx">📥 처리된 파일 다운로드</a>'
                st.markdown(href, unsafe_allow_html=True)

                # 사용자들에게 이메일 전송
                if st.button("📧 등록된 사용자에게 내원 환자 이메일 알림 보내기"):
                    users = db.collection("users").stream()
                    for user_doc in users:
                        uid = user_doc.id
                        user_data = user_doc.to_dict()
                        email = user_data.get("email")  # 필요시 DB에 미리 등록되어 있어야 함
                        patients = user_data.get("patients", [])
                        matches = []
                        for patient in patients:
                            name, number = patient["name"], str(patient["number"])
                            match_rows = df[(df["환자명"] == name) & (df["환자번호"].astype(str) == number)]
                            if not match_rows.empty:
                                matches.append(match_rows)

                        if matches and email:
                            combined = pd.concat(matches)
                            send_email(uid, email, combined)

                    st.success("📤 이메일 전송 완료!")

            except Exception as e:
                st.error(f"❌ 복호화 실패: {str(e)}")


# 이메일 발송 함수
def send_email(user_id, to_email, matched_df):
    msg = MIMEMultipart()
    msg['Subject'] = f"[환자 내원 알림] {user_id}님 등록 환자 내원"
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_email

    body = f"{user_id}님,\n아래는 내원한 환자 정보입니다:\n\n{matched_df.to_string(index=False)}"
    msg.attach(MIMEText(body, "plain"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)
