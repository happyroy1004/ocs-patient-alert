import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore

# 🔐 Firebase 서비스 계정 키 불러오기 (secrets.toml에서)
firebase_config = st.secrets["firebase"]


# 🔐 Firebase Admin SDK 초기화
if not firebase_admin._apps:
    cred = credentials.Certificate(firebase_config)
    firebase_admin.initialize_app(cred)

# 🧠 Firestore DB 연결
db = firestore.client()

# 🔧 테스트용 UI
st.title("🩺 OCS 환자 알림 시스템")

# 환자 이름 입력
patient_name = st.text_input("환자 이름을 입력하세요")

if st.button("환자 등록"):
    if patient_name.strip():
        doc_ref = db.collection("patients").document()
        doc_ref.set({"name": patient_name})
        st.success(f"환자 {patient_name} 등록 완료!")
    else:
        st.warning("환자 이름을 입력해주세요.")

# 등록된 환자 목록 보기
st.subheader("📋 등록된 환자 목록")
patients = db.collection("patients").stream()
for p in patients:
    st.write(f"- {p.to_dict().get('name')}")
s
