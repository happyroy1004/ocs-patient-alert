import streamlit as st
import pandas as pd
import json
import firebase_admin
from firebase_admin import credentials, db
import msoffcrypto
import io
import re

# 🔑 Firebase 초기화
if not firebase_admin._apps:
    cred = credentials.Certificate(st.secrets["firebase"])
    firebase_admin.initialize_app(cred, {
        'databaseURL': st.secrets["database_url"]
    })

# 🔒 사용자 Google ID 입력
st.title("🔐 환자 등록 & 엑셀 중복 검사 시스템")
google_id_input = st.text_input("Google 계정 ID를 입력하세요", key="google_id")

# 🔐 Firebase 경로용으로 Google ID 정제
def sanitize_key(key: str) -> str:
    return re.sub(r'[.#$/\[\]]', '_', key)

safe_google_id = sanitize_key(google_id_input) if google_id_input else None

# ✅ 사용자 환자 등록
st.header("1️⃣ 환자 정보 등록")
with st.form("register_patient_form"):
    name = st.text_input("환자 이름")
    patient_number = st.text_input("환자 번호")
    submitted = st.form_submit_button("등록하기")
    if submitted and safe_google_id and name and patient_number:
        ref = db.reference(f"patients/{safe_google_id}")
        ref.push({"name": name, "number": patient_number})
        st.success("✅ 환자 정보가 성공적으로 등록되었습니다.")
    elif submitted:
        st.error("⚠️ 모든 필드를 입력해주세요.")

# 📄 업로드된 암호화된 엑셀 파일 처리
st.header("2️⃣ 엑셀 파일 업로드 및 중복 검사")
uploaded_file = st.file_uploader("비밀번호로 보호된 .xlsx 파일 업로드", type=["xlsx"])
excel_password = st.text_input("엑셀 파일 비밀번호", type="password")

if uploaded_file and excel_password and safe_google_id:
    try:
        # 🔓 암호화된 파일 복호화
        decrypted = io.BytesIO()
        file = msoffcrypto.OfficeFile(uploaded_file)
        file.load_key(password=excel_password)
        file.decrypt(decrypted)

        # 📖 엑셀 내용 읽기
        decrypted.seek(0)
        df = pd.read_excel(decrypted, engine="openpyxl")

        if not {'이름', '번호'}.issubset(df.columns):
            st.error("❌ '이름'과 '번호' 열이 포함된 엑셀 파일을 업로드해주세요.")
        else:
            # 🔍 Firebase에서 현재 사용자 환자 목록 가져오기
            ref = db.reference(f"patients/{safe_google_id}")
            existing_patients = ref.get() or {}

            duplicates = []
            for _, row in df.iterrows():
                for patient in existing_patients.values():
                    if row['이름'] == patient['name'] and str(row['번호']) == str(patient['number']):
                        duplicates.append(f"{row['이름']} ({row['번호']})")
                        break

            if duplicates:
                st.error(f"❗ 중복 환자 발견:\n" + "\n".join(duplicates))
            else:
                st.success("✅ 중복된 환자 정보가 없습니다.")
    except Exception as e:
        st.error(f"❌ 오류 발생: {str(e)}")

# 👀 사용자 등록 환자 리스트 출력
st.header("📋 내 등록 환자 목록")
if safe_google_id:
    try:
        ref = db.reference(f"patients/{safe_google_id}")
        patient_data = ref.get()
        if patient_data:
            for patient in patient_data.values():
                st.markdown(f"- {patient['name']} ({patient['number']})")
        else:
            st.write("등록된 환자가 없습니다.")
    except Exception as e:
        st.error(f"❌ 환자 목록 불러오기 실패: {str(e)}")
