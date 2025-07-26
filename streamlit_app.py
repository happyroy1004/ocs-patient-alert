import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, db
import msoffcrypto
import io

# Firebase Realtime Database 연결
if not firebase_admin._apps:
    cred = credentials.Certificate("firebase_key.json")
    firebase_admin.initialize_app(cred, {
        "databaseURL": "https://ocs-patientalert-default-rtdb.firebaseio.com"
    })

st.title("📋 환자 등록 및 조회")

# 🔑 Google ID 입력
google_id = st.text_input("Google ID를 입력하세요:")

if not google_id:
    st.warning("Google ID를 먼저 입력해주세요.")
    st.stop()

# 🔐 암호화된 Excel 파일 업로드 및 복호화
uploaded_file = st.file_uploader("🔓 암호화된 Excel 파일 업로드", type=["xls", "xlsx"])
password = st.text_input("엑셀 파일 암호", type="password")

if uploaded_file and password:
    decrypted = io.BytesIO()
    try:
        office_file = msoffcrypto.OfficeFile(uploaded_file)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)
        decrypted.seek(0)

        # 📄 모든 시트 읽기 (두 번째 행을 컬럼명으로 인식)
        xls = pd.ExcelFile(decrypted)
        sheet_names = xls.sheet_names

        st.success("✅ 파일 복호화 성공")

        for sheet_name in sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=1)  # 두 번째 행을 컬럼명으로 지정
                st.subheader(f"📑 시트: {sheet_name}")

                if '환자명' not in df.columns or '진료번호' not in df.columns:
                    st.error("❌ '환자명' 또는 '진료번호' 열을 찾을 수 없습니다.")
                    continue

                df_show = df[['환자명', '진료번호']].dropna()
                st.dataframe(df_show)

                # 🔍 이미 등록된 환자 불러오기
                ref = db.reference(f"patients/{google_id}")
                existing_data = ref.get() or {}

                # 📥 중복 제거 및 새 환자 등록
                new_entries = 0
                for _, row in df_show.iterrows():
                    name = str(row['환자명']).strip()
                    number = str(row['진료번호']).strip()
                    key = f"{name}_{number}"

                    if key not in existing_data:
                        ref.child(key).set({
                            "이름": name,
                            "번호": number
                        })
                        new_entries += 1

                st.success(f"✅ 새로 등록된 환자 수: {new_entries}")

                # 📋 전체 환자 보기
                updated_data = ref.get()
                if updated_data:
                    st.markdown("### 🔎 전체 등록 환자")
                    result_df = pd.DataFrame([
                        {"이름": v["이름"], "번호": v["번호"]}
                        for v in updated_data.values()
                    ])
                    st.dataframe(result_df)
                else:
                    st.info("아직 등록된 환자가 없습니다.")
            except Exception as e:
                st.error(f"❌ 시트 '{sheet_name}' 처리 중 오류 발생: {e}")
    except Exception as e:
        st.error(f"❌ 파일 복호화 실패: {e}")
else:
    st.info("파일과 암호를 모두 입력해야 환자 데이터를 불러올 수 있습니다.")
