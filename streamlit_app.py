import streamlit as st
import firebase_admin
from firebase_admin import credentials

# ✅ firebase_credentials 를 dict 로 처리
firebase_cred_dict = {
    "type": st.secrets["firebase_credentials"]["type"],
    "project_id": st.secrets["firebase_credentials"]["project_id"],
    "private_key_id": st.secrets["firebase_credentials"]["private_key_id"],
    "private_key": st.secrets["firebase_credentials"]["private_key"],
    "client_email": st.secrets["firebase_credentials"]["client_email"],
    "client_id": st.secrets["firebase_credentials"]["client_id"],
    "auth_uri": st.secrets["firebase_credentials"]["auth_uri"],
    "token_uri": st.secrets["firebase_credentials"]["token_uri"],
    "auth_provider_x509_cert_url": st.secrets["firebase_credentials"]["auth_provider_x509_cert_url"],
    "client_x509_cert_url": st.secrets["firebase_credentials"]["client_x509_cert_url"],
    "universe_domain": st.secrets["firebase_credentials"]["universe_domain"]
}

# ✅ dict 전달
cred = credentials.Certificate(firebase_cred_dict)
firebase_admin.initialize_app(cred, {
    'databaseURL': st.secrets["firebase"]["database_url"]
})

# Load Firebase credentials from local file
cred = credentials.Certificate("firebase_key.json")
firebase_admin.initialize_app(cred, {
    'databaseURL': st.secrets["firebase"]["database_url"]
})

st.title("Patient Alert App")

# Dummy Data Display
st.write("Firebase and Gmail credentials loaded successfully.")


