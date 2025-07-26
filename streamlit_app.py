import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, db
import msoffcrypto
import io

st.set_page_config(page_title="환자 등록 확인기", page_icon="🦷", layout="wide")
st.title("🦷 환자 등록 확인기")

# 🔑 Firebase 연결
if "firebase_initialized" not in st.session_state:
    try:
        cred = credentials.Certificate("firebase_key.json")  # 반드시 이 json 파일이 같이 있어야 함
        firebase_admin.initialize_app(cred, {
            'databaseURL': st.secrets["database_url"]
        })
        st.session_state.firebase_initialized = True
    except Exception as e:
        st.error("Firebase 초기화 실패: " + str(e))

# 🔐 암호화된 Excel 업로드
st.header("🔓 암호화된 Excel 파일 업로드")
encrypted_file = st.file_uploader("🔒 암호화된 Excel 파일 (.xlsx)", type=["xlsx"])
password = st.text_input("📌 파일 암호를 입력하세요", type="password")

# 🔑 Google ID 입력
st.header("👤 사용자 정보 입력")
google_id = st.text_input("Google ID를 입력하세요 (예: your_email@gmail.com)")

# Firebase 참조 경로 (이메일 특수문자 제거)
def sanitize_id(raw_id: str) -> str:
    return raw_id.replace("@", "_at_").replace(".", "_dot_")

firebase_key = sanitize_id(google_id) if google_id else None

# ✅ 기존 등록된 환자 불러오기
if firebase_key:
    ref = db.reference(f"patients/{firebase_key}")
    existing_data = ref.get()
    existing_set = set()
    if existing_data:
        for item in existing_data.values():
            name = str(item.get("name")).strip()
            number = str(item.get("number")).strip()
            existing_set.add((name, number))

    st.subheader("📄 기존 등록된 환자 목록")
    if existing_data:
        existing_df = pd.DataFrame.from_dict(existing_data, orient="index")
        if {"name", "number"}.issubset(existing_df.columns):
            st.dataframe(existing_df[["name", "number"]])
        else:
            st.dataframe(existing_df)
            st.warning("⚠️ 'name' 또는 'number' 컬럼이 없어 전체 데이터를 출력했습니다.")
    else:
        st.info("ℹ️ 등록된 환자가 없습니다.")

# ✅ 엑셀 복호화 및 판별
if encrypted_file and password and firebase_key:
    try:
        decrypted = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(encrypted_file)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)

        xls = pd.ExcelFile(decrypted)
        sheet_names = xls.sheet_names

        for sheet_name in sheet_names:
            st.subheader(f"📑 시트: {sheet_name}")
            df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)

            if "환자명" not in df.columns or "진료번호" not in df.columns:
                st.warning("❌ '환자명' 또는 '진료번호' 열이 없습니다.")
                continue

            results = []
            for _, row in df.iterrows():
                name = str(row["환자명"]).strip()
                number = str(row["진료번호"]).strip()
                exists = (name, number) in existing_set
                results.append({
                    "환자명": name,
                    "진료번호": number,
                    "등록 여부": "✅ 등록됨" if exists else "➕ 미등록"
                })

            result_df = pd.DataFrame(results)
            st.dataframe(result_df)

    except Exception as e:
        st.error(f"❌ 파일 복호화 또는 처리 중 오류 발생: {e}")
