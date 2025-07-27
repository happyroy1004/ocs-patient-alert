import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, db
import io
import msoffcrypto
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from openpyxl import load_workbook
from openpyxl.styles import Font

# 🔐 Firebase 초기화
if not firebase_admin._apps:
    cred = credentials.Certificate(st.secrets["firebase_credentials"])
    firebase_admin.initialize_app(cred, {
        'databaseURL': st.secrets["firebase"]["database_url"]
    })

# 📌 Firebase-safe 경로 변환
def sanitize_path(email):
    return email.replace(".", "_dot_").replace("@", "_at_")

# 📩 이메일 주소 복원
def recover_email(safe_id: str) -> str:
    if safe_id.endswith("_com"):
        safe_id = safe_id[:-4] + ".com"
    return safe_id.replace("_at_", "@").replace("_dot_", ".")

# 🔒 암호화된 엑셀 여부 확인
def is_encrypted_excel(file):
    try:
        file.seek(0)
        return msoffcrypto.OfficeFile(file).is_encrypted()
    except Exception:
        return False

# 📂 엑셀 로드 + Colab 스타일 처리
def process_excel_file(file, password):
    decrypted = io.BytesIO()
    office_file = msoffcrypto.OfficeFile(file)
    office_file.load_key(password=password)
    office_file.decrypt(decrypted)
    decrypted.seek(0)

    wb = load_workbook(filename=decrypted, data_only=True)
    processed_sheets = {}

    sheet_name_mapping = {
        '교정': '교정', '교정과': '교정',
        '구강내과': '내과', '내과': '내과',
        '구강악안면외과': '외과', '외과': '외과',
        '보존과': '보존', '보존': '보존',
        '보철과': '보철', '보철': '보철',
        '소아치과': '소치', '소치': '소치',
        '원내생진료센터': '원내생', '원내생': '원내생',
        '원스톱협진센터': '원스톱', '원스톱': '원스톱',
        '임플란트진료센터': '임플란트', '임플란트': '임플란트',
        '치주과': '치주', '치주': '치주',
        '임플실': '임플란트', '원진실': '원내생'
    }

    professors_dict = {
        '소치': ['김현태', '장기택', '김정욱', '현홍근', '김영재', '신터전', '송지수'],
        '보존': ['이인복', '금기연', '이우철', '유연지', '서덕규', '이창하', '김선영', '손원준'],
        '외과': ['최진영', '서병무', '명훈', '김성민', '박주영', '양훈주', '한정준', '권익재'],
        '치주': ['구영', '이용무', '설양조', '구기태', '김성태', '조영단'],
        '보철': ['곽재영', '김성균', '임영준', '김명주', '권호범', '여인성', '윤형인', '박지만', '이재현', '조준호'],
        '교정': [], '내과': [], '원내생': [], '원스톱': [], '임플란트': [],
    }

    def process_sheet(df, professors_list, sheet_key):
        df = df.drop(columns=['예약일시'], errors='ignore')
        df = df.sort_values(by=['예약의사', '예약시간'])
        professors = df[df['예약의사'].isin(professors_list)]
        non_professors = df[~df['예약의사'].isin(professors_list)]

        if sheet_key != '보철':
            non_professors = non_professors.sort_values(by=['예약시간', '예약의사'])
        else:
            non_professors = non_professors.sort_values(by=['예약의사', '예약시간'])

        final_rows = []
        current_time = None
        current_doctor = None

        for _, row in non_professors.iterrows():
            if sheet_key != '보철':
                if current_time != row['예약시간']:
                    if current_time is not None:
                        final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
                    current_time = row['예약시간']
            else:
                if current_doctor != row['예약의사']:
                    if current_doctor is not None:
                        final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
                    current_doctor = row['예약의사']
            final_rows.append(row)

        final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
        final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
        final_rows.append(pd.Series(["<교수님>"] + [" "] * (len(df.columns) - 1), index=df.columns))

        current_professor = None
        for _, row in professors.iterrows():
            if current_professor != row['예약의사']:
                if current_professor is not None:
                    final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
                current_professor = row['예약의사']
            final_rows.append(row)

        final_df = pd.DataFrame(final_rows, columns=df.columns)
        final_df = final_df[['진료번호', '예약시간', '환자명', '예약의사', '진료내역']]
        return final_df

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        values = list(ws.values)
        while values and (values[0] is None or all(v is None for v in values[0])):
            values.pop(0)
        if len(values) < 2:
            continue
        df = pd.DataFrame(values)
        df.columns = df.iloc[0]
        df = df.drop([0]).reset_index(drop=True)
        df = df.fillna("").astype(str)
        df['예약의사'] = df['예약의사'].str.strip().str.replace(" 교수님", "", regex=False)
        df['환자명'] = df['환자명'].str.strip()
        df['진료번호'] = df['진료번호'].str.strip()

        sheet_key = sheet_name_mapping.get(sheet_name.strip(), None)
        if not sheet_key:
            continue

        professors_list = professors_dict.get(sheet_key, [])
        processed_df = process_sheet(df, professors_list, sheet_key)
        processed_sheets[sheet_name] = processed_df

    if not processed_sheets:
        return None, None

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        for sheet_name, df in processed_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output_buffer.seek(0)

    wb2 = load_workbook(output_buffer)
    for sheet_name in wb2.sheetnames:
        ws = wb2[sheet_name]
        header = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            if row[0].value == "<교수님>":
                for cell in row:
                    if cell.value:
                        cell.font = Font(bold=True)
            if sheet_name.strip() == "교정" and '진료내역' in header:
                idx = header['진료내역'] - 1
                cell = row[idx]
                text = str(cell.value)
                if any(keyword in text for keyword in ['본딩', 'bonding']):
                    cell.font = Font(bold=True)

    final_output = io.BytesIO()
    wb2.save(final_output)
    final_output.seek(0)
    return final_output, processed_sheets

# 📧 이메일 전송
def send_email(receiver, rows, sender, password):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver
        msg['Subject'] = "📌 등록 환자 내원 알림"

        html_table = rows.to_html(index=False, escape=False)
        body = f"다음 등록 환자가 내원했습니다:<br><br>{html_table}"
        msg.attach(MIMEText(body, 'html'))

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        return str(e)

# 🌐 Streamlit 시작
st.title("🩺 환자 내원 확인 시스템")
user_id = st.text_input("아이디를 입력하세요")
if not user_id:
    st.stop()

firebase_key = sanitize_path(user_id)

if user_id != "admin":
    # 일반 사용자 모드 생략 (기존과 동일)
    pass
else:
    st.subheader("📂 엑셀 업로드 및 사용자 일치 검사")
    uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])
    if uploaded_file:
        password = None
        if is_encrypted_excel(uploaded_file):
            password = st.text_input("🔑 엑셀 파일 비밀번호 입력", type="password")
            if not password:
                st.stop()

        try:
            output_file, matched_all_sheets = process_excel_file(uploaded_file, password)
            if not output_file:
                st.warning("⚠ 처리된 시트가 없습니다.")
                st.stop()

            uploaded_file_name = uploaded_file.name
            if uploaded_file_name.endswith(".xlsx"):
                processed_name = uploaded_file_name.replace(".xlsx", "_processed.xlsx")
            elif uploaded_file_name.endswith(".xlsm"):
                processed_name = uploaded_file_name.replace(".xlsm", "_processed.xlsx")
            else:
                processed_name = uploaded_file_name + "_processed.xlsx"

            st.download_button("⬇️ 처리된 파일 다운로드", output_file.read(), file_name=processed_name)

            sender = st.secrets["gmail"]["sender"]
            sender_pw = st.secrets["gmail"]["app_password"]

            users_ref = db.reference("patients")
            all_users = users_ref.get()
            if not all_users:
                st.warning("❗ 등록된 사용자가 없습니다.")
                st.stop()

            matched_users = []

            for uid, plist in all_users.items():
                registered_set = set((str(v["환자명"]).strip(), str(v["진료번호"]).strip()) for v in plist.values())
                matched_rows = []
                for df in matched_all_sheets.values():
                    temp_df = df.copy()
                    temp_df[["환자명", "진료번호"]] = temp_df[["환자명", "진료번호"]].astype(str).apply(lambda x: x.str.strip())
                    match_df = temp_df[temp_df.apply(lambda row: (row["환자명"], row["진료번호"]) in registered_set, axis=1)]
                    if not match_df.empty:
                        matched_rows.append(match_df)
                if matched_rows:
                    combined = pd.concat(matched_rows, ignore_index=True)
                    matched_users.append((uid, combined))

            if matched_users:
                st.success(f"🔍 {len(matched_users)}명의 사용자와 일치하는 환자 발견됨.")
                if st.button("📤 메일 보내기"):
                    for uid, df_matched in matched_users:
                        real_email = recover_email(uid)
                        result = send_email(real_email, df_matched, sender, sender_pw)
                        if result is True:
                            st.success(f"✅ {real_email} 전송 완료")
                        else:
                            st.error(f"❌ {real_email} 전송 실패: {result}")
                else:
                    for uid, df in matched_users:
                        st.markdown(f"### 📧 {recover_email(uid)}")
                        st.dataframe(df)
            else:
                st.info("📭 매칭된 사용자 없음")

        except Exception as e:
            st.error(f"❌ 파일 처리 실패: {e}")
