import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore

# 🔐 secrets.toml에서 firebase key 로드
firebase_config = st.secrets["FIREBASE_KEY"]

# ✅ Firebase Admin 초기화
if not firebase_admin._apps:
    cred = credentials.Certificate(firebase_config)
    firebase_admin.initialize_app(cred)

# ✅ Firestore 인스턴스
db = firestore.client()

# 테스트 UI
st.title("Firebase 연결 테스트")

name = st.text_input("이름을 입력하세요")
if st.button("저장"):
    db.collection("users").document(name).set({"name": name})
    st.success("저장 완료!")

if st.button("불러오기"):
    docs = db.collection("users").stream()
    for doc in docs:
        st.write(doc.id, doc.to_dict())
