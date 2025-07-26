import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore

st.set_page_config(page_title="OCS 환자 알림 시스템", page_icon="🩺")

st.title("🩺 OCS 환자 알림 시스템")
st.write("환자 이름을 입력하세요")

# ✅ Firebase secrets에서 인증 정보 불러오기
firebase_config = dict(st.secrets["firebase"])

# ✅ Firebase Admin SDK 초기화
if not firebase_admin._apps:
    cred = credentials.Certificate(firebase_config)
    firebase_admin.initialize_app(cred)

# ✅ Firestore 클라이언트 생성
db = firestore.client()
patients_ref = db.collection("patients")

# ✅ 환자 입력 필드
patient_name = st.text_input("환자 이름")

# ✅ 등록 버튼 클릭 시 Firestore에 저장
if st.button("환자 등록"):
    if patient_name.strip():
        patients_ref.add({"name": patient_name})
        st.success(f"'{patient_name}' 등록 완료!")
    else:
        st.warning("환자 이름을 입력하세요.")

# ✅ 등록된 환자 목록 보여주기
st.markdown("## 📋 등록된 환자 목록")

docs = patients_ref.stream()
for doc in docs:
    data = doc.to_dict()
    st.write(f"- {data.get('name', '(이름 없음)')}")
