import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, db
import io
import msoffcrypto

# 🔐 Firebase 초기화
if not firebase_admin._apps:
    cred = credentials.Certificate(st.secrets["firebase_credentials"])
    firebase_admin.initialize_app(cred, {
        'databaseURL': st.secrets["database_url"]
    })

# 📌 Firebase-safe 경로로 변환
def sanitize_path(s):
    import re
    return re.sub(r'[.$#[\]/]', '_', s)

# 🧾 엑셀 파일 복호화
def decrypt_excel(file, password):
    decrypted = io.BytesIO()
    office_file = msoffcrypto.OfficeFile(file)
    office_file.load_key(password=password)
    office_file.decrypt(decrypted)
    decrypted.seek(0)
    return decrypted

# 📁 Streamlit 앱
st.title("🔒 암호화된 OCS 환자 파일 분석기")

# 1️⃣ 구글 아이디 입력
google_id = st.text_input("Google ID를 입력하세요 (예: your_email@gmail.com)")
if not google_id:
    st.stop()
firebase_key = sanitize_path(google_id)

# 2️⃣ 기존 등록된 환자 목록 표시
ref = db.reference(f"patients/{firebase_key}")
existing_data = ref.get()

if existing_data:
    st.subheader("📄 등록된 환자 목록")
    existing_df = pd.DataFrame(existing_data.values())  # ← 🔄 키 없이 값만 가져옴
    st.dataframe(existing_df[["name", "number"]])
else:
    st.info("아직 등록된 환자가 없습니다.")

# 3️⃣ 신규 환자 등록
with st.form("register_patient"):
    st.subheader("➕ 신규 환자 등록")
    new_name = st.text_input("환자명")
    new_number = st.text_input("진료번호")
    submitted = st.form_submit_button("등록")

    if submitted:
        if not new_name or not new_number:
            st.warning("환자명과 진료번호를 모두 입력해주세요.")
        else:
            if existing_data and any(v.get("name") == new_name and v.get("number") == new_number for v in existing_data.values()):
                st.error("이미 등록된 환자입니다.")
            else:
                new_ref = ref.push()
                new_ref.set({"name": new_name, "number": new_number})
                st.success(f"환자 {new_name} ({new_number})가 등록되었습니다.")
                st.rerun()

# 4️⃣ 엑셀 업로드 및 분석
st.subheader("🔐 OCS 엑셀 업로드 및 분석")
uploaded_file = st.file_uploader("암호화된 Excel(.xlsx/.xlsm) 파일 업로드", type=["xlsx", "xlsm"])
password = st.text_input("Excel 파일 암호 입력", type="password")

if uploaded_file and password:
    try:
        decrypted = decrypt_excel(uploaded_file, password)
        xl = pd.ExcelFile(decrypted)

        # 🔍 등록 환자 목록 준비
        registered_set = set()
        if existing_data:
            registered_set = {(v.get("name"), str(v.get("number"))) for v in existing_data.values()}

        for sheet_name in xl.sheet_names:
            try:
                df = xl.parse(sheet_name, header=1)
                if "환자명" not in df.columns or "진료번호" not in df.columns:
                    continue  # name, number 없는 시트는 건너뜀

                df = df.rename(columns={"환자명": "name", "진료번호": "number"})
                df["number"] = df["number"].astype(str)
                df["name"] = df["name"].astype(str)

                matched_df = df[df.apply(lambda row: (row["name"], row["number"]) in registered_set, axis=1)]

                if not matched_df.empty:
                    st.markdown(f"### 📋 시트: {sheet_name}")
                    st.success("✅ 등록된 환자 정보가 있는 행만 표시합니다.")
                    st.dataframe(matched_df.reset_index(drop=True))

            except Exception as e:
                st.error(f"❌ 시트 '{sheet_name}' 처리 중 오류 발생: {e}")

    except Exception as e:
        st.error(f"❌ 복호화 실패: {e}")
