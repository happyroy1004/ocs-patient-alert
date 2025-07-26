import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, db
import io

# -------------------- Firebase 초기화 --------------------
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
        "client_x509_cert_url": st.secrets["firebase"]["client_x509_cert_url"],
        "universe_domain": st.secrets["firebase"]["universe_domain"]
    })
    firebase_admin.initialize_app(cred, {
        "databaseURL": st.secrets["database_url"]
    })

# -------------------- 앱 UI --------------------
st.title("📋 OCS 환자 등록 & 조회")

uploaded_file = st.file_uploader("🗂 OCS 엑셀 파일을 업로드하세요", type=["xls", "xlsx"])
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=None)  # 모든 시트 불러오기
        all_data = []

        for sheet_name, sheet_df in df.items():
            if "환자명" in sheet_df.columns and "진료번호" in sheet_df.columns:
                # 빈 행 제거
                clean_df = sheet_df.dropna(subset=["환자명", "진료번호"])
                for _, row in clean_df.iterrows():
                    all_data.append({
                        "이름": str(row["환자명"]).strip(),
                        "번호": str(row["진료번호"]).strip(),
                        "진료과": str(row.get("진료과", "")).strip()
                    })
        extracted_df = pd.DataFrame(all_data)
        st.success(f"✅ {len(extracted_df)}명의 환자 정보를 불러왔습니다.")
        st.dataframe(extracted_df)

        # Firebase에 이미 등록된 환자 불러오기
        ref = db.reference("patients")
        existing_patients = ref.get() or {}
        existing_keys = {f"{v['이름']}_{v['번호']}" for v in existing_patients.values()}

        # 새로 등록할 환자만 필터링
        new_patients = extracted_df[
            ~extracted_df.apply(lambda x: f"{x['이름']}_{x['번호']}", axis=1).isin(existing_keys)
        ]

        st.write("🆕 새로 등록할 환자:")
        st.dataframe(new_patients)

        if st.button("📤 Firebase에 환자 등록"):
            for _, row in new_patients.iterrows():
                key = f"{row['이름']}_{row['번호']}"
                ref.push({
                    "이름": row["이름"],
                    "번호": row["번호"],
                    "진료과": row["진료과"]
                })
            st.success("✅ 새 환자 등록 완료!")

    except Exception as e:
        st.error(f"❌ 처리 중 오류 발생: {e}")

# -------------------- 환자 목록 조회 --------------------
st.header("📖 등록된 환자 목록")
ref = db.reference("patients")
all_patients = ref.get() or {}

if all_patients:
    df_registered = pd.DataFrame(all_patients.values())
    st.dataframe(df_registered)
else:
    st.info("현재 등록된 환자가 없습니다.")
