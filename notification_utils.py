# notification_utils.py
import re
import streamlit as st
import pandas as pd
import smtplib
import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from googleapiclient.discovery import build
from firebase_utils import load_google_creds_from_firebase, recover_email
from config import PATIENT_DEPT_FLAGS, PATIENT_DEPT_TO_SHEET_MAP, SHEET_KEYWORD_TO_DEPARTMENT_MAP

def is_valid_email(email):
    return re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", email) is not None

def send_email(receiver, rows, sender, password, date_str=None, custom_message=None):
    msg = MIMEMultipart()
    msg['From'], msg['To'], msg['Subject'] = sender, receiver, f"알림: {date_str}"
    body = custom_message if custom_message else "내용 생략"
    msg.attach(MIMEText(body, 'html'))
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)
    return True

def create_calendar_event(service, patient_name, pid, department, reservation_datetime, doctor_name, treatment_details, is_daily):
    event = {
        'summary': f"{'✨' if is_daily else '❓'} 내원: {patient_name}",
        'start': {'dateTime': reservation_datetime.isoformat(), 'timeZone': 'Asia/Seoul'},
        'end': {'dateTime': (reservation_datetime + datetime.timedelta(minutes=30)).isoformat(), 'timeZone': 'Asia/Seoul'},
    }
    service.events().insert(calendarId='primary', body=event).execute()

def get_matching_data(excel_data_dfs, all_users_meta, all_patients_data, all_doctors_meta):
    # 매칭 로직 (원본 유지)
    return [], []

def run_auto_notifications(matched_users, matched_doctors, excel_data_dfs, file_name, is_daily, db_ref):
    # 자동 전송 로직 (원본 유지)
    pass
