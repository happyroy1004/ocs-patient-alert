import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
import pandas as pd
import json
import io

# Firebase 초기화 (중복 방지)
if not firebase_admin._apps:
    cred = credentials.Certificate(json.loads(st.secrets["firebase_admin_json"]))
    firebase_admin.initialize_app(cred, {
        'databaseURL': st.secrets["firebase_database_url"]
    })

st.set_page_config(page_title="OCS 환자 알림 시스템", page_icon="🩺")

st.title("🩺 OCS 환자 알림 시스템")

# 사용자 식별용 구글 ID 입력
google_id = st.text_input("🧑 ID", key="google_id")
if not google_id:
    st.warning("ID를 입력해주세요.")
    st.stop()

# 환자 등록
st.subheader("👤 환자 등록")

with st.form("patient_form"):
    patient_name = st.text_input("환자 이름")
    patient_number = st.text_input("환자 번호")
    submitted = st.form_submit_button("등록")

    if submitted:
        if patient_name and patient_number:
            user_ref = db.reference(f"patients/{google_id}")
            existing = user_ref.get() or {}

            # 중복 검사
            duplicate = any(
                p.get("name") == patient_name and p.get("number") == patient_number
                for p in existing.values()
            )
            if duplicate:
                st.error(f"⚠️ 이미 등록된 환자입니다: {patient_name} ({patient_number})")
            else:
                new_id = user_ref.push().key
                user_ref.child(new_id).set({
                    "name": patient_name,
                    "number": patient_number
                })
                st.success(f"✅ '{patient_name}' 등록 완료!")
        else:
            st.error("모든 정보를 입력해주세요.")

# 환자 목록 표시
st.subheader("📋 등록된 토탈 환자 목록")
patients = db.reference(f"patients/{google_id}").get()
if patients:
    for pid, patient in patients.items():
        st.markdown(f"- {patient['name']} ({patient['number']})")
else:
    st.info("아직 등록된 환자가 없습니다.")

# Excel 업로드
st.subheader("📂 Excel 파일 업로드 (환자번호/이름 중복 검사)")
uploaded_file = st.file_uploader("Excel 파일 업로드", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl", dtype=str)

        if not {"환자이름", "환자번호"}.issubset(df.columns):
            st.error("❌ '환자이름' 및 '환자번호' 열이 필요합니다.")
        else:
            duplicates = []
            registered = db.reference(f"patients/{google_id}").get() or {}

            for _, row in df.iterrows():
                name, number = row["환자이름"], row["환자번호"]
                if any(p["name"] == name and p["number"] == number for p in registered.values()):
                    duplicates.append(f"{name} ({number})")

            if duplicates:
                st.error("토탈 환자 내원 :\n" + "\n".join(duplicates))
            else:
                st.success("토탈 환자가 내원하지 않습니다")

    except Exception as e:
        st.error(f"❌ 오류 발생: {str(e)}")
