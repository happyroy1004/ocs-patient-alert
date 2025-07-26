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
def load_excel(file):
    try:
        file.seek(0)
        office_file = msoffcrypto.OfficeFile(file)
        if office_file.is_encrypted():
            return True, office_file
        else:
            file.seek(0)
            return False, pd.ExcelFile(file)
    except Exception as e:
        raise ValueError(f"엑셀 처리 실패: {e}")

# 📁 Streamlit 앱 시작
st.title("📁 토탈환자 내원확인")

# 1️⃣ 구글 아이디 입력
google_id = st.text_input("Google ID를 입력하세요 (예: your_email@gmail.com)")
if not google_id:
    st.stop()
firebase_key = sanitize_path(google_id)

# 2️⃣ 등록된 환자 목록 조회
ref = db.reference(f"patients/{firebase_key}")
existing_data = ref.get()

st.subheader("📄 등록된 토탈환자 목록")
if existing_data:
    for key, val in existing_data.items():
        st.write(f"👤 이름: {val.get('환자명', '없음')}  ")
        st.write(f"🆔 번호: {val.get('진료번호', '없음')}")
        if st.button("❌ 삭제", key=key):
            ref.child(key).delete()
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

# 4️⃣ 엑셀 업로드
st.subheader("🔐 OCS 엑셀 업로드 및 분석")
uploaded_file = st.file_uploader("암호화된 또는 일반 Excel(.xlsx/.xlsm) 파일 업로드", type=["xlsx", "xlsm"])

if uploaded_file:
    is_encrypted, result = load_excel(uploaded_file)

    password = None
    if is_encrypted:
        password = st.text_input("🔑 파일이 암호화되어 있습니다. 암호를 입력해주세요", type="password")
        if password:
            try:
                decrypted = io.BytesIO()
                result.load_key(password=password)
                result.decrypt(decrypted)
                decrypted.seek(0)
                xl = pd.ExcelFile(decrypted)
            except Exception as e:
                st.error(f"❌ 암호 해제 실패: {e}")
                st.stop()
        else:
            st.stop()
    else:
        xl = result

    # 🔍 등록된 환자 세트
    registered_set = set((d["환자명"], d["진료번호"]) for d in existing_data.values()) if existing_data else set()
    found_any = False

    for sheet_name in xl.sheet_names:
        try:
            df = xl.parse(sheet_name, header=1)
            if "환자명" not in df.columns or "진료번호" not in df.columns:
                continue

            all_patients = df[["환자명", "진료번호"]].dropna().astype(str)
            matched = all_patients[all_patients.apply(lambda row: (row["환자명"], row["진료번호"]) in registered_set, axis=1)]

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
