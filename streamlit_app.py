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
st.title("🔒 토탈환자 내원확인")

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
        # 누적 등록된 환자 명단
registered_set = set((d["name"], d["number"]) for d in existing_data.values()) if existing_data else set()

found_any = False  # ✅ 등록된 환자가 적어도 한 명이라도 발견되었는지 추적

for sheet_name in xl.sheet_names:
    df = xl.parse(sheet_name, header=1)

    if "name" not in df.columns or "number" not in df.columns:
        continue

    all_patients = df[["name", "number"]].dropna()
    all_patients = all_patients.astype(str)

    matched = all_patients[all_patients.apply(lambda row: (row["name"], row["number"]) in registered_set, axis=1)]

    if not matched.empty:
        found_any = True
        st.markdown(f"### 📋 시트: {sheet_name}")
        st.markdown("🗂️ 전체 환자 목록")
        st.dataframe(all_patients)
        st.checkbox("✅ 등록된 환자만 필터링", value=True, key=f"filter_{sheet_name}")
        st.dataframe(matched)

# ✅ 아무 시트에도 등록된 환자가 없으면 안내 메시지 출력
if not found_any:
    st.warning("🔎 토탈 환자 내원 예정 없습니다.")

    except Exception as e:
        st.error(f"❌ 복호화 실패: {e}")
