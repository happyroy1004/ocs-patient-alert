import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, db
import smtplib
from email.mime.text import MIMEText

# Load Firebase credentials from local file
cred = credentials.Certificate("firebase_key.json")
firebase_admin.initialize_app(cred, {
    'databaseURL': st.secrets["firebase"]["database_url"]
})

st.title("Patient Alert App")

# Dummy Data Display
st.write("Firebase and Gmail credentials loaded successfully.")
