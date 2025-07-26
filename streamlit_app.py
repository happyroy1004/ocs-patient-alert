import streamlit as st
import pandas as pd
import msoffcrypto
import io
import firebase_admin
from firebase_admin import credentials, db

# 🔐 Firebase 인증 및 초기화
if not firebase_admin._apps:
    cred = credentials.Certificate({
        "type": st.secrets["firebase"]["type"],
        "project_id": st.secrets["firebase"]["project_id"],
        "private_key_id": st.secrets["firebase"]["private_key_id"],
        "private_key": st.secrets["firebase"]["private_key"].replace("\\n", "\n"),
        "client_email": st.secrets["firebase"]["client_email"],
        "client_id": st.secrets["firebase"]["client_id"],
        "auth_uri": st.secrets["firebase"]["auth_uri"],
        "token_uri": st.secrets["firebase"]["token_uri"],
        "auth_provider_x509_cert_url": st.secrets["firebase"]["auth_provider_x509_cert_url"],
        "client_x509_cert_url": st.secrets["firebase"]["client_x509_cert_url"]
    })
    firebase_admin.initialize_app(cred, {
        'databaseURL': st.secrets["database_url"]
    })

st.title("📁 환자 등록 및 조회 시스템")

# 1️⃣ Google ID 입력
google_id = st.text_input("🔑 Google ID를 입력하세요")

if google_id:
    # 2️⃣ 파일 업로드 및 복호화
    uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xls", "xlsx"])

    if uploaded_file:
        password = st.text_input("🔐 파일 암호를 입력하세요", type="password")
        if password:
            try:
                decrypted = io.BytesIO()
                office_file = msoffcrypto.OfficeFile(uploaded_file)
                office_file.load_key(password=password)
                office_file.decrypt(decrypted)

                df = pd.read_excel(decrypted, sheet_name=None)  # 모든 시트 불러오기
                st.success("✅ 파일 복호화 및 로딩 완료")

                for sheet_name, sheet_df in df.items():
                    st.subheader(f"📋 시트: {sheet_name}")
                    sheet_df = sheet_df.dropna(how="all")  # 전체 빈 행 제거
                    if sheet_df.empty:
                        st.info("⚠️ 데이터가 없습니다.")
                        continue

                    # '환자명' 또는 '이름', '진료번호' 또는 '번호' 열 자동 감지
                    name_col = next((col for col in sheet_df.columns if '환자명' in col or '이름' in col), None)
                    number_col = next((col for col in sheet_df.columns if '진료번호' in col or '번호' in col), None)

                    if not name_col or not number_col:
                        st.warning("❌ '환자명' 또는 '진료번호' 열을 찾을 수 없습니다.")
                        continue

                    sheet_df = sheet_df[[name_col, number_col]].dropna()

                    ref = db.reference(f"patients/{google_id}")
                    existing_data = ref.get() or {}

                    new_entries = []
                    for _, row in sheet_df.iterrows():
                        name = str(row[name_col]).strip()
                        number = str(row[number_col]).strip()
                        key = f"{name}_{number}"

                        if key not in existing_data:
                            new_entries.append({"name": name, "number": number})
                            ref.child(key).set({"name": name, "number": number})

                    st.success(f"✅ 새로운 환자 {len(new_entries)}명 등록 완료")
                    st.dataframe(pd.DataFrame(new_entries) if new_entries else pd.DataFrame(existing_data.values()))

            except Exception as e:
                st.error(f"❌ 오류 발생: {e}")
        else:
            st.info("🔑 파일 암호를 입력해주세요.")
else:
    st.warning("👤 먼저 Google ID를 입력해주세요.")
