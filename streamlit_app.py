import streamlit as st
import pandas as pd
import msoffcrypto
import io
import firebase_admin
from firebase_admin import credentials, db

# Firebase 초기화
if not firebase_admin._apps:
    cred = credentials.Certificate({
        "type": st.secrets["firebase"]["type"],
        "project_id": st.secrets["firebase"]["project_id"],
        "private_key_id": st.secrets["firebase"]["private_key_id"],
        "private_key": st.secrets["firebase"]["private_key"],
        "client_email": st.secrets["firebase"]["client_email"],
        "client_id": st.secrets["firebase"]["client_id"],
        "auth_uri": st.secrets["firebase"]["auth_uri"],
        "token_uri": st.secrets["firebase"]["token_uri"],
        "auth_provider_x509_cert_url": st.secrets["firebase"]["auth_provider_x509_cert_url"],
        "client_x509_cert_url": st.secrets["firebase"]["client_x509_cert_url"],
        "universe_domain": st.secrets["firebase"]["universe_domain"],
    })
    firebase_admin.initialize_app(cred, {
        "databaseURL": st.secrets["database_url"]
    })

st.title("🔐 환자 등록 확인 시스템")
st.markdown("Google ID로 등록된 환자만 확인 가능하며, 암호화된 Excel 파일을 사용합니다.")

# 1. 사용자로부터 Google ID 입력받기
google_id = st.text_input("👤 Google ID를 입력하세요 (예: example@gmail.com)")

# 이메일에 불가능한 문자가 있는지 체크
def is_valid_path_string(s):
    return all(c not in s for c in ".#$[]")

# 2. 파일 업로드 및 복호화
uploaded_file = st.file_uploader("🔒 암호화된 Excel 파일 업로드", type=["xlsx", "xlsm"])
password = st.text_input("암호를 입력하세요", type="password")

df = None
if uploaded_file and password:
    try:
        office_file = msoffcrypto.OfficeFile(uploaded_file)
        office_file.load_key(password=password)
        decrypted = io.BytesIO()
        office_file.decrypt(decrypted)
        df = pd.read_excel(decrypted, sheet_name=None)

        st.success("✅ 복호화 및 업로드 성공!")

    except Exception as e:
        st.error(f"❌ 복호화 실패: {e}")

# 3. Firebase에서 해당 Google ID로 등록된 환자 불러오기
if google_id:
    if not is_valid_path_string(google_id):
        st.error("❌ Google ID에는 '.', '#', '$', '[', ']' 문자를 포함할 수 없습니다.")
    else:
        try:
            ref = db.reference(f"patients/{google_id.replace('.', '_')}")
            existing_data = ref.get()
            if existing_data:
                existing_df = pd.DataFrame(existing_data.values())
                st.success("✅ 등록된 환자 목록")
                st.dataframe(existing_df[["name", "number"]])
            else:
                st.info("ℹ️ 등록된 환자가 없습니다.")
        except Exception as e:
            st.error(f"❌ 환자 목록 불러오기 실패: {e}")

# 4. 업로드된 엑셀에서 환자명 + 진료번호 체크
if df and google_id and is_valid_path_string(google_id):
    try:
        ref = db.reference(f"patients/{google_id.replace('.', '_')}")
        existing_data = ref.get()
        existing_set = set()
        if existing_data:
            for record in existing_data.values():
                existing_set.add((record["name"], record["number"]))

        for sheet_name, sheet_df in df.items():
            try:
                sheet_df.columns = sheet_df.iloc[0]
                sheet_df = sheet_df.drop(sheet_df.index[0])
                sheet_df = sheet_df.rename(columns=lambda x: str(x).strip())

                # '성명'과 '진료번호'를 기준으로 환자 식별
                if "성명" not in sheet_df.columns or "진료번호" not in sheet_df.columns:
                    st.warning(f"⚠️ 시트 '{sheet_name}'에 '성명' 또는 '진료번호' 열이 없습니다.")
                    continue

                sheet_df = sheet_df[["성명", "진료번호"]].dropna()
                sheet_df.columns = ["name", "number"]

                sheet_df["등록여부"] = sheet_df.apply(
                    lambda row: "✅ 등록됨" if (row["name"], str(row["number"])) in existing_set else "❌ 미등록",
                    axis=1
                )
                st.subheader(f"📄 시트: {sheet_name}")
                st.dataframe(sheet_df)

            except Exception as e:
                st.error(f"❌ 시트 '{sheet_name}' 처리 중 오류 발생: {e}")
    except Exception as e:
        st.error(f"❌ 전체 처리 중 오류: {e}")
