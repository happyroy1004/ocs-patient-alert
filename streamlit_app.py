import streamlit as st
import firebase_admin
from firebase_admin import credentials, db
import pandas as pd
import io
from google.oauth2 import id_token
from google.auth.transport import requests
from streamlit.runtime.scriptrunner import get_script_run_ctx
import json

# ------------------------------
# 1. Firebase 초기화
# ------------------------------
if not firebase_admin._apps:
    cred = credentials.Certificate(st.secrets["firebase"])
    firebase_admin.initialize_app(cred, {
        "databaseURL": st.secrets["firebase"]["database_url"]
    })

# ------------------------------
# 2. 사용자 인증
# ------------------------------
def get_google_user_id():
    ctx = get_script_run_ctx()
    if ctx is None or ctx.session_id is None:
        return None

    if "user_email" not in st.session_state:
        st.warning("⚠️ 인증된 사용자가 아닙니다.")
        return None

    return st.session_state.user_email.replace(".", "_")  # Firebase 키용

# ------------------------------
# 3. Firebase에 등록된 환자 목록 불러오기
# ------------------------------
def get_registered_patients(user_id):
    ref = db.reference(f"patients/{user_id}")
    data = ref.get()
    if data:
        return pd.DataFrame(data.values())
    return pd.DataFrame(columns=["환자명", "환자번호"])

# ------------------------------
# 4. 중복 검사
# ------------------------------
def is_duplicate(patient_df, name, number):
    return not patient_df[(patient_df["환자명"] == name) & (patient_df["환자번호"] == number)].empty

# ------------------------------
# 5. Streamlit UI
# ------------------------------
st.title("📋 환자 등록 및 중복 검사")

user_id = get_google_user_id()
if user_id is None:
    st.stop()

st.markdown(f"**현재 사용자:** `{user_id.replace('_', '.')}`")

# Firebase에서 기존 환자 불러오기
existing_df = get_registered_patients(user_id)

uploaded_file = st.file_uploader("Excel 파일 업로드", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # 최소한의 컬럼 확인
    if not all(col in df.columns for col in ["환자명", "환자번호"]):
        st.error("❌ '환자명'과 '환자번호' 열이 필요합니다.")
        st.stop()

    new_patients = []
    duplicate_patients = []

    for _, row in df.iterrows():
        name = str(row["환자명"]).strip()
        number = str(row["환자번호"]).strip()

        if is_duplicate(existing_df, name, number):
            duplicate_patients.append({"환자명": name, "환자번호": number})
        else:
            new_patients.append({"환자명": name, "환자번호": number})

    # 결과 표시
    st.success(f"✅ 신규 등록 환자 수: {len(new_patients)}")
    st.warning(f"⚠️ 중복 환자 수: {len(duplicate_patients)}")

    if new_patients:
        with st.expander("📥 신규 환자 등록"):
            for p in new_patients:
                st.write(f"- {p['환자명']} / {p['환자번호']}")
            if st.button("📌 Firebase에 신규 환자 저장"):
                ref = db.reference(f"patients/{user_id}")
                for p in new_patients:
                    ref.push(p)
                st.success("🎉 저장 완료! 새로고침 해보세요.")

    if duplicate_patients:
        with st.expander("❗ 중복 환자 목록"):
            for p in duplicate_patients:
                st.write(f"- {p['환자명']} / {p['환자번호']}")
