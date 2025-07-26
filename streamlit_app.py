import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
import pandas as pd
import io

# --- Firebase 초기화 ---
firebase_config = st.secrets["firebase"]
database_url = firebase_config["database_url"]

cred = credentials.Certificate({
    "type": firebase_config["type"],
    "project_id": firebase_config["project_id"],
    "private_key_id": firebase_config["private_key_id"],
    "private_key": firebase_config["private_key"],
    "client_email": firebase_config["client_email"],
    "client_id": firebase_config["client_id"],
    "auth_uri": firebase_config["auth_uri"],
    "token_uri": firebase_config["token_uri"],
    "auth_provider_x509_cert_url": firebase_config["auth_provider_x509_cert_url"],
    "client_x509_cert_url": firebase_config["client_x509_cert_url"],
    "universe_domain": firebase_config["universe_domain"]
})

if not firebase_admin._apps:
    firebase_admin.initialize_app(cred, {
        'databaseURL': database_url
    })

# --- 사용자 Google ID 입력 ---
st.title("📋 환자 중복 등록 확인")
google_id = st.text_input("Google 계정 ID를 입력하세요:")

if google_id:
    ref = db.reference(f"patients/{google_id}")

    uploaded_file = st.file_uploader("📎 엑셀 파일을 업로드하세요", type=["xlsx"])

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            if not {'이름', '차트번호'}.issubset(df.columns):
                st.error("❌ 엑셀 파일에 '이름'과 '차트번호' 열이 존재해야 합니다.")
            else:
                data = df[['이름', '차트번호']].astype(str)
                new_patients = []

                existing = ref.get() or {}

                for _, row in data.iterrows():
                    key = f"{row['이름']}_{row['차트번호']}"
                    if key in existing:
                        st.warning(f"⚠️ 이미 존재하는 환자: {key}")
                    else:
                        new_patients.append((key, row.to_dict()))

                if new_patients:
                    st.success(f"✅ 새로 등록될 환자 수: {len(new_patients)}")
                    if st.button("💾 등록"):
                        for key, patient_data in new_patients:
                            ref.child(key).set(patient_data)
                        st.success("저장 완료!")

        except Exception as e:
            st.error(f"❌ 오류 발생: {e}")

    # 📤 전체 환자 목록 다운로드
    if st.button("⬇️ 전체 환자 목록 다운로드"):
        try:
            data = ref.get()
            if data:
                df_all = pd.DataFrame(data.values())
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_all.to_excel(writer, index=False)
                st.download_button(
                    label="📥 다운로드 (xlsx)",
                    data=buffer.getvalue(),
                    file_name="patients.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("📭 등록된 환자가 없습니다.")
        except Exception as e:
            st.error(f"❌ 환자 목록 불러오기 실패: {e}")
