import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, db
import msoffcrypto
import io

st.set_page_config(page_title="🔐 환자 등록기", layout="centered")

st.title("🩺 환자 중복 등록 검사")

# ✅ Firebase 인증 초기화
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
        "client_x509_cert_url": st.secrets["firebase"]["client_x509_cert_url"],
        "universe_domain": st.secrets["firebase"]["universe_domain"]
    })
    db_url = st.secrets["database_url"]
    firebase_admin.initialize_app(cred, {"databaseURL": db_url})

# ✅ 사용자 인증 (Google 계정으로 로그인한 ID 입력)
google_id = st.text_input("📧 본인의 Google 계정 ID를 입력해주세요 (예: abc123@gmail.com)", key="google_id")
safe_google_id = google_id.replace(".", "(dot)")

# ✅ 암호화된 엑셀 파일 업로드
uploaded_file = st.file_uploader("🔐 암호화된 Excel 파일 업로드", type=["xls", "xlsx"])
password = st.text_input("🔑 Excel 파일 암호를 입력하세요", type="password")

if uploaded_file and password:
    # 복호화 처리
    office_file = msoffcrypto.OfficeFile(uploaded_file)
    try:
        office_file.load_key(password=password)
        decrypted = io.BytesIO()
        office_file.decrypt(decrypted)

        # 엑셀 읽기
        decrypted.seek(0)
        df = pd.read_excel(decrypted, engine="openpyxl")

        # OCS 기준 열명
        name_col = "환자명"
        number_col = "진료번호"

        if not {name_col, number_col}.issubset(df.columns):
            st.error(f"❌ '{name_col}'과 '{number_col}' 열이 포함된 엑셀 파일을 업로드해주세요.")
        else:
            # Firebase에서 기존 등록된 환자 불러오기
            ref = db.reference(f"patients/{safe_google_id}")
            existing_patients = ref.get() or {}

            duplicates = []
            for _, row in df.iterrows():
                patient_name = str(row[name_col]).strip()
                patient_number = str(row[number_col]).strip()

                for patient in existing_patients.values():
                    if (patient.get("name") == patient_name and
                        str(patient.get("number")) == patient_number):
                        duplicates.append(f"{patient_name} ({patient_number})")
                        break

            # 중복 결과 출력
            if duplicates:
                st.warning("⚠️ 이미 등록된 환자입니다:")
                for d in duplicates:
                    st.write(f"• {d}")
            else:
                st.success("✅ 중복 환자 없음! 모두 신규 등록 가능합니다.")

                # 등록 버튼
                if st.button("✅ Firebase에 신규 환자 등록하기"):
                    for _, row in df.iterrows():
                        patient_name = str(row[name_col]).strip()
                        patient_number = str(row[number_col]).strip()

                        # 중복 확인 후 등록
                        already_exists = False
                        for patient in existing_patients.values():
                            if (patient.get("name") == patient_name and
                                str(patient.get("number")) == patient_number):
                                already_exists = True
                                break
                        if not already_exists:
                            ref.push({
                                "name": patient_name,
                                "number": patient_number
                            })

                    st.success("🎉 환자 정보가 성공적으로 등록되었습니다!")

    except Exception as e:
        st.error(f"❌ 오류 발생: {e}")
