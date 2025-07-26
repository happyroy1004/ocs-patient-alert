import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd
import io

# Firebase secrets 불러오기
firebase_config = dict(st.secrets["firebase"])

# Firebase 초기화
if not firebase_admin._apps:
    cred = credentials.Certificate(firebase_config)
    firebase_admin.initialize_app(cred)

db = firestore.client()

# 앱 제목
st.title("🩺 OCS 환자 알림 시스템")

# Google ID 입력
google_id = st.text_input("📧 담당자의 Google 이메일을 입력하세요", key="google_id")

# 엑셀 파일 업로드
uploaded_file = st.file_uploader("📂 환자 명단이 담긴 Excel 파일을 업로드하세요", type=["xlsx", "xls"])

if uploaded_file and google_id:
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
        st.success("✅ 엑셀 파일 업로드 완료")

        # DataFrame 미리보기
        st.subheader("📋 업로드된 데이터 미리보기")
        st.dataframe(df)

        # '환자 이름'이라는 컬럼이 있다면 Firestore에 저장
        if '환자 이름' in df.columns:
            for name in df['환자 이름'].dropna():
                doc_ref = db.collection("patients").document(name)
                doc_ref.set({"name": name, "google_id": google_id})
            st.success("🎉 모든 환자 정보를 Firebase에 저장했습니다.")
        else:
            st.warning("❗ '환자 이름'이라는 컬럼이 Excel에 포함되어 있어야 합니다.")

    except Exception as e:
        st.error(f"❌ 오류 발생: {e}")

# Firestore에서 전체 목록 출력
st.subheader("📜 등록된 환자 목록")

patients_ref = db.collection("patients").stream()
for doc in patients_ref:
    patient = doc.to_dict()
    st.markdown(f"- {patient['name']} ({patient['google_id']})")
