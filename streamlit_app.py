import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore

# ✅ Firebase 초기화
if not firebase_admin._apps:
    firebase_config = st.secrets["firebase"]
    cred = credentials.Certificate(dict(firebase_config))
    firebase_admin.initialize_app(cred)

db = firestore.client()

st.title("📋 환자 등록 및 파일 업로드")

# ✅ 사용자 ID 입력
google_id = st.text_input("당신의 Google ID를 입력하세요 (예: user@gmail.com)")

if google_id:
    # Firestore에서 사용자 문서 참조
    user_ref = db.collection("users").document(google_id)
    user_doc = user_ref.get()

    if not user_doc.exists:
        user_ref.set({"patients": []})  # 초기화

    # ✅ 환자 등록
    with st.form("register_patient"):
        st.subheader("👤 새 환자 등록")
        patient_name = st.text_input("환자 이름")
        patient_number = st.text_input("환자 번호")
        submitted = st.form_submit_button("등록")

        if submitted and patient_name and patient_number:
            existing_patients = user_ref.get().to_dict().get("patients", [])
            duplicate = any(
                p["name"] == patient_name and p["number"] == patient_number for p in existing_patients
            )
            if duplicate:
                st.warning("이미 등록된 환자입니다.")
            else:
                new_entry = {"name": patient_name, "number": patient_number}
                user_ref.update({"patients": firestore.ArrayUnion([new_entry])})
                st.success("✅ 환자가 등록되었습니다!")

    # ✅ 현재 등록된 환자 목록
    st.subheader("📑 내 등록 환자 목록")
    user_data = user_ref.get().to_dict()
    for p in user_data.get("patients", []):
        st.markdown(f"- {p['name']} ({p['number']})")

    # ✅ 엑셀 업로드 및 중복 확인
    st.subheader("📁 Excel 파일 업로드 (환자번호/이름 중복 검사)")
    uploaded_file = st.file_uploader("Excel 파일 업로드", type=["xlsx"])

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file, dtype=str)
            st.write("📄 업로드된 데이터 미리보기:")
            st.dataframe(df)

            # '이름', '환자번호' 열 존재 여부 확인
            if not all(col in df.columns for col in ["이름", "환자번호"]):
                st.error("❌ Excel 파일에 '이름'과 '환자번호' 열이 있어야 합니다.")
            else:
                registered = user_ref.get().to_dict().get("patients", [])
                duplicates = []

                for _, row in df.iterrows():
                    for p in registered:
                        if row["이름"] == p["name"] and row["환자번호"] == p["number"]:
                            duplicates.append(p)

                if duplicates:
                    st.warning("토탈 환자가 내원합니다:")
                    for d in duplicates:
                        st.markdown(f"- {d['name']} ({d['number']})")
                else:
                    st.success("토탈 환자가 아무도 내원하지 않습니다!")
        except Exception as e:
            st.error(f"❌ 오류 발생: {e}")
else:
    st.info("👤 먼저 Google ID를 입력하세요.")
