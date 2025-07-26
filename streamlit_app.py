import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore
import json

st.set_page_config(page_title="OCS 환자 알림 시스템", page_icon="🩺")

st.title("🩺 OCS 환자 알림 시스템")
st.write("환자 이름을 입력하세요")

# 🔐 Firebase credentials 가져오기
firebase_config = json.loads(st.secrets["firebase"].to_json())

# 🔑 Firebase Admin SDK 초기화
if not firebase_admin._apps:
    cred = credentials.Certificate(firebase_config)
    firebase_admin.initialize_app(cred)

# 🔥 Firestore 클라이언트 연결
db = firestore.client()
patients_ref = db.collection("patients")

# 📝 입력 폼
patient_name = st.text_input("")

if st.button("환자 등록"):
    if patient_name.strip() != "":
        # Firestore에 데이터 추가
        patients_ref.add({"name": patient_name})
        st.success(f"'{patient_name}' 등록 완료!")
    else:
        st.warning("환자 이름을 입력하세요.")

# 📋 등록된 환자 목록 출력
st.markdown("## 📋 등록된 환자 목록")

docs = patients_ref.stream()
for doc in docs:
    patient = doc.to_dict()
    st.write(f"- {patient.get('name', '이름 없음')}")
