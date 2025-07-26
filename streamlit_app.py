import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd
import json
import io

# --- Firebase 초기화 ---
if not firebase_admin._apps:
    cred = credentials.Certificate(json.loads(json.dumps(st.secrets["firebase"])))
    firebase_admin.initialize_app(cred)

db = firestore.client()

# --- 제목 ---
st.title("🩺 OCS 환자 알림 시스템")

# --- 환자 등록 폼 ---
st.subheader("환자 이름을 입력하세요")

with st.form("patient_form"):
    patient_name = st.text_input("환자 이름")
    google_id = st.text_input("구글 아이디 (예: example@gmail.com)")
    uploaded_file = st.file_uploader("📄 Excel 파일 업로드 (.xlsx)", type=["xlsx"])
    submitted = st.form_submit_button("환자 등록")

    if submitted:
        if patient_name.strip() == "" or google_id.strip() == "":
            st.error("⚠️ 이름과 구글 아이디를 모두 입력해주세요.")
        else:
            # Firestore에 저장
            db.collection("patients").add({
                "name": patient_name,
                "google_id": google_id
            })
            st.success(f"✅ '{patient_name}' 등록 완료!")

            # 엑셀 파일 처리
            if uploaded_file:
                try:
                    df = pd.read_excel(uploaded_file)
                    st.subheader("📊 업로드된 Excel 내용")
                    st.dataframe(df)
                except Exception as e:
                    st.error(f"엑셀 파일을 읽는 중 오류 발생: {e}")

# --- 등록된 환자 목록 출력 ---
st.subheader("📋 등록된 환자 목록")

try:
    patients = db.collection("patients").stream()
    for doc in patients:
        patient = doc.to_dict()
        name = patient.get("name", "이름 없음")
        google_id = patient.get("google_id", "이메일 없음")
        st.markdown(f"- {name} ({google_id})")
except Exception as e:
    st.error(f"데이터를 불러오는 중 오류 발생: {e}")
