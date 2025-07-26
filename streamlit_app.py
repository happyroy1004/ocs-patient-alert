import streamlit as st
import pandas as pd
import json
import firebase_admin
from firebase_admin import credentials, db
import msoffcrypto
import io

# ✅ Firebase 초기화
if not firebase_admin._apps:
    firebase_config = st.secrets["firebase"]
    database_url = st.secrets["database_url"]
    cred = credentials.Certificate(firebase_config)
    firebase_admin.initialize_app(cred, {
        'databaseURL': database_url
    })

# ✅ 사용자 입력
st.title("🩺 환자 등록 및 중복 검사 시스템")
google_id = st.text_input("👤 구글 ID를 입력하세요", key="google_id")

# ✅ 환자 등록 폼
st.header("1️⃣ 환자 등록")
with st.form("register_patient_form"):
    name = st.text_input("환자 이름")
    patient_number = st.text_input("환자 번호")
    submitted = st.form_submit_button("등록하기")
    if submitted and google_id and name and patient_number:
        ref = db.reference(f"patients/{google_id}")
        ref.push({"name": name, "number": patient_number})
        st.success("✅ 환자 정보가 등록되었습니다.")
    elif submitted:
        st.error("⚠️ 모든 필드를 입력해주세요.")

# ✅ 파일 업로드 및 중복 검사
st.header("2️⃣ Excel 파일 업로드 및 중복 검사")
uploaded_file = st.file_uploader("🔒 비밀번호로 보호된 .xlsx 파일 업로드", type=["xlsx"])
excel_password = st.text_input("📎 엑셀 파일 비밀번호", type="password")

if uploaded_file and excel_password:
    try:
        # ✅ 엑셀 파일 복호화
        decrypted = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(uploaded_file)
        office_file.load_key(password=excel_password)
        office_file.decrypt(decrypted)
        decrypted.seek(0)

        # ✅ 엑셀 읽기
        df = pd.read_excel(decrypted, engine="openpyxl")
        if not {"이름", "번호"}.issubset(df.columns):
            st.error("❌ '이름'과 '번호' 열이 포함된 엑셀 파일을 업로드해주세요.")
        else:
            ref = db.reference(f"patients/{google_id}")
            existing_patients = ref.get() or {}

            duplicates = []
            for _, row in df.iterrows():
                for patient in existing_patients.values():
                    if row["이름"] == patient["name"] and str(row["번호"]) == str(patient["number"]):
                        duplicates.append(f"{row['이름']} ({row['번호']})")
                        break

            if duplicates:
                st.error("❗ 중복 환자 발견:\n" + "\n".join(duplicates))
            else:
                st.success("✅ 중복된 환자 정보가 없습니다.")
    except Exception as e:
        st.error(f"❌ 오류 발생: {e}")

# ✅ 등록된 환자 리스트 출력
st.header("📋 내 등록 환자 목록")
if google_id:
    try:
        ref = db.reference(f"patients/{google_id}")
        patient_data = ref.get()
        if patient_data:
            for patient in patient_data.values():
                st.markdown(f"- {patient['name']} ({patient['number']})")
        else:
            st.write("등록된 환자가 없습니다.")
    except Exception as e:
        st.error(f"❌ 환자 목록 불러오기 실패: {e}")
