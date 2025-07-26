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

# 기존 환자 목록 표시
ref = db.reference(f"patients/{firebase_key}")
existing_data = ref.get()
if existing_data:
    st.subheader("📄 기존 등록된 환자 목록")
    existing_df = pd.DataFrame(existing_data).T

    # '이름'과 '번호' 컬럼이 있는지 확인
    if "이름" in existing_df.columns and "번호" in existing_df.columns:
        st.dataframe(existing_df[["이름", "번호"]])
    else:
        st.dataframe(existing_df)  # 전체 컬럼 보여주기
        st.warning("❗ '이름' 또는 '번호' 컬럼이 없어 전체 데이터를 출력했습니다.")
else:
    st.info("아직 등록된 환자가 없습니다.")

# 3️⃣ 새로운 환자 등록
with st.form("register_patient"):
    st.subheader("➕ 신규 환자 등록")
    new_name = st.text_input("환자명")
    new_number = st.text_input("진료번호")
    submitted = st.form_submit_button("등록")

    if submitted:
        if not new_name or not new_number:
            st.warning("환자명과 진료번호를 모두 입력해주세요.")
        else:
            # 중복 확인
            if existing_data and any(v["이름"] == new_name and v["번호"] == new_number for v in existing_data.values()):
                st.error("이미 등록된 환자입니다.")
            else:
                new_ref = ref.push()
                new_ref.set({"이름": new_name, "번호": new_number})
                st.success(f"환자 {new_name} ({new_number})가 등록되었습니다.")
                st.experimental_rerun()

# 4️⃣ 엑셀 파일 업로드 + 복호화
st.subheader("🔐 OCS 엑셀 업로드 및 분석")
uploaded_file = st.file_uploader("암호화된 Excel(.xlsx/.xlsm) 파일 업로드", type=["xlsx", "xlsm"])
password = st.text_input("Excel 파일 암호 입력", type="password")

if uploaded_file and password:
    try:
        decrypted = decrypt_excel(uploaded_file, password)
        xl = pd.ExcelFile(decrypted)
        for sheet_name in xl.sheet_names:
            try:
                df = xl.parse(sheet_name, header=1)
                if "환자명" not in df.columns or "진료번호" not in df.columns:
                    st.warning(f"❌ 시트 '{sheet_name}'에서 '환자명' 또는 '진료번호' 열을 찾을 수 없습니다.")
                    continue

                patients_in_sheet = df[["환자명", "진료번호"]].dropna()
                patients_in_sheet.columns = ["이름", "번호"]

                st.markdown(f"### 📋 시트: {sheet_name}")
                st.dataframe(patients_in_sheet)

            except Exception as e:
                st.error(f"❌ 시트 '{sheet_name}' 처리 중 오류 발생: {e}")

    except Exception as e:
        st.error(f"❌ 복호화 실패: {e}")
