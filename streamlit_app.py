import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd
import io

# Firebase 인증 및 초기화
firebase_config = st.secrets["firebase"]
cred = credentials.Certificate(dict(firebase_config))
firebase_admin.initialize_app(cred)
db = firestore.client()

# 🔑 사용자 로그인 (구글 ID 입력 기반)
st.title("🔔 OCS 환자 알림 시스템")
user_google_id = st.text_input("📧 구글 아이디를 입력하세요", key="google_id")

if not user_google_id:
    st.warning("구글 아이디를 먼저 입력해주세요.")
    st.stop()

st.markdown("---")

# 📝 환자 등록
st.subheader("👤 환자 등록")
with st.form("register_form"):
    name = st.text_input("환자 이름")
    patient_number = st.text_input("환자 번호")
    submitted = st.form_submit_button("환자 등록")

    if submitted:
        if not name or not patient_number:
            st.error("이름과 환자 번호를 모두 입력하세요.")
        else:
            doc_ref = db.collection("users").document(user_google_id).collection("patients").document(f"{name}_{patient_number}")
            doc_ref.set({"name": name, "number": patient_number})
            st.success(f"'{name}' (번호: {patient_number}) 등록 완료!")

# 📄 사용자별 등록된 환자 목록 보기
st.subheader("📋 등록된 환자 목록")
patients_ref = db.collection("users").document(user_google_id).collection("patients")
patients = patients_ref.stream()
for p in patients:
    info = p.to_dict()
    st.markdown(f"- {info['name']} (번호: {info['number']})")

st.markdown("---")

# 📤 Excel 파일 업로드
st.subheader("📑 Excel 파일 업로드")
uploaded_file = st.file_uploader("Excel 파일을 업로드하세요 (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
        st.write("📄 업로드된 데이터", df)

        if not {"이름", "번호"}.issubset(df.columns):
            st.error("엑셀 파일에 '이름'과 '번호' 열이 모두 포함되어야 합니다.")
            st.stop()

        duplicates = []
        for row in df.itertuples(index=False):
            name_excel = getattr(row, '이름')
            number_excel = getattr(row, '번호')
            doc = patients_ref.document(f"{name_excel}_{number_excel}").get()
            if doc.exists:
                duplicates.append(f"{name_excel} (번호: {number_excel})")

        if duplicates:
            st.warning("❗ 등록된 환자와 일치하는 항목이 파일에 포함되어 있습니다:")
            for d in duplicates:
                st.markdown(f"- {d}")
        else:
            st.success("✅ 중복된 환자 없음!")

    except Exception as e:
        st.error(f"파일 처리 중 오류 발생: {e}")
