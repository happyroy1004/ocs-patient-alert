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

# 🧾 엑셀 파일 복호화 or 일반 처리
def load_excel(file, password=None):
    try:
        file.seek(0)
        office_file = msoffcrypto.OfficeFile(file)
        if office_file.is_encrypted():
            if not password:
                raise ValueError("암호화된 파일입니다. 암호를 입력해주세요.")
            decrypted = io.BytesIO()
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            return pd.ExcelFile(decrypted)
        else:
            file.seek(0)
            return pd.ExcelFile(file)
    except Exception as e:
        raise ValueError(f"엑셀 처리 실패: {e}")

# 📁 Streamlit 앱 시작
st.title("🔒 토탈환자 내원확인")

# 1️⃣ 구글 아이디 입력
google_id = st.text_input("Google ID를 입력하세요 (예: your_email@gmail.com)")
if not google_id:
    st.stop()
firebase_key = sanitize_path(google_id)

# 2️⃣ 등록된 환자 목록 조회
ref = db.reference(f"patients/{firebase_key}")
existing_data = ref.get()

if existing_data:
    st.subheader("📄 등록된 환자 목록")
    existing_df = pd.DataFrame(existing_data.values())
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
        elif existing_data and any(v.get("name") == new_name and v.get("number") == new_number for v in existing_data.values()):
            st.error("이미 등록된 환자입니다.")
        else:
            new_ref = ref.push()
            new_ref.set({"name": new_name, "number": new_number})
            st.success(f"환자 {new_name} ({new_number})가 등록되었습니다.")
            st.rerun()

# 4️⃣ 엑셀 업로드 및 분석
st.subheader("🔐 OCS 엑셀 업로드 및 분석")
uploaded_file = st.file_uploader("암호화된 또는 일반 Excel(.xlsx/.xlsm) 파일 업로드", type=["xlsx", "xlsm"])
password = st.text_input("Excel 파일 암호 입력 (암호 없을 경우 비워두세요)", type="password")

if uploaded_file:
    try:
        xl = load_excel(uploaded_file, password)

        registered_set = set((d["name"], d["number"]) for d in existing_data.values()) if existing_data else set()
        found_any = False

        for sheet_name in xl.sheet_names:
            try:
                df = xl.parse(sheet_name, header=1)
                if "name" not in df.columns or "number" not in df.columns:
                    continue

                all_patients = df[["name", "number"]].dropna().astype(str)
                matched = all_patients[all_patients.apply(lambda row: (row["name"], row["number"]) in registered_set, axis=1)]

                if not matched.empty:
                    found_any = True
                    st.markdown(f"### 📋 시트: {sheet_name}")
                    st.markdown("🗂️ 전체 환자 목록")
                    st.dataframe(all_patients)

                    if st.checkbox("✅ 등록된 환자만 필터링", value=True, key=f"filter_{sheet_name}"):
                        st.dataframe(matched)

            except Exception as e:
                st.error(f"❌ 시트 '{sheet_name}' 처리 중 오류 발생: {e}")

        if not found_any:
            st.warning("🔎 토탈 환자 내원 예정 없습니다.")

    except Exception as e:
        st.error(f"❌ 파일 처리 실패: {e}")
