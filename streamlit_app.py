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

# 🧾 엑셀 파일 암호화 여부 확인
def is_encrypted_excel(file):
    try:
        file.seek(0)
        office_file = msoffcrypto.OfficeFile(file)
        return office_file.is_encrypted()
    except Exception:
        return False

# 🧾 엑셀 파일 로드
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
st.title("📁 토탈환자 내원확인")

# 1️⃣ 구글 아이디 입력
google_id = st.text_input("아이디를 입력하세요")
if not google_id:
    st.stop()
firebase_key = sanitize_path(google_id)

# 2️⃣ 등록된 환자 목록 조회
ref = db.reference(f"patients/{firebase_key}")
existing_data = ref.get()

st.subheader("📄 등록된 토탈환자 목록")
if existing_data:
    for key, val in existing_data.items():
        col1, col2, col3 = st.columns([4, 4, 2])
        with col1:
            st.write(f"👤 이름: {val.get('환자명', '없음')}")
        with col2:
            st.write(f"🆔 번호: {val.get('진료번호', '없음')}")
        with col3:
            if st.button("❌ 삭제", key=f"delete_{key}"):
                db.reference(f"patients/{firebase_key}/{key}").delete()
                st.success("삭제되었습니다.")
                st.rerun()
else:
    st.info("아직 등록된 환자가 없습니다.")

# 3️⃣ 신규 환자 등록
with st.form("register_patient"):
    st.subheader("➕ 신규 토탈환자 등록")
    new_name = st.text_input("환자명")
    new_number = st.text_input("진료번호")
    submitted = st.form_submit_button("등록")

    if submitted:
        if not new_name or not new_number:
            st.warning("환자명과 진료번호를 모두 입력해주세요.")
        elif existing_data and any(v.get("환자명") == new_name and v.get("진료번호") == new_number for v in existing_data.values()):
            st.error("이미 등록된 환자입니다.")
        else:
            new_ref = ref.push()
            new_ref.set({"환자명": new_name, "진료번호": new_number})
            st.success(f"환자 {new_name} ({new_number})가 등록되었습니다.")
            st.rerun()

# 4️⃣ 엑셀 업로드 및 분석
st.subheader("📂 OCS 엑셀 업로드")
uploaded_file = st.file_uploader("Excel(.xlsx/.xlsm) 파일 업로드", type=["xlsx", "xlsm"])

password = None
if uploaded_file:
    encrypted = is_encrypted_excel(uploaded_file)

    if encrypted:
        password = st.text_input("🔑 암호화된 파일입니다. 암호를 입력하세요", type="password")
        if not password:
            st.stop()

    try:
        xl = load_excel(uploaded_file, password=password if encrypted else None)
        registered_set = set((d["환자명"], d["진료번호"]) for d in existing_data.values()) if existing_data else set()
        found_any = False

        for sheet_name in xl.sheet_names:
            try:
                df = xl.parse(sheet_name, header=1)
                if "환자명" not in df.columns or "진료번호" not in df.columns:
                    continue
                df = df.astype(str)
                matched_df = df[df.apply(lambda row: (row["환자명"], row["진료번호"]) in registered_set, axis=1)]

                if not matched_df.empty:
                    found_any = True
                    st.markdown(f"### 📋 시트: {sheet_name}")
                    st.dataframe(matched_df)

            except Exception as e:
                st.error(f"❌ 시트 '{sheet_name}' 처리 오류: {e}")

        if not found_any:
            st.warning("🔎 토탈 환자 내원 예정 없습니다.")

    except Exception as e:
        st.error(f"❌ 파일 처리 실패: {e}")
