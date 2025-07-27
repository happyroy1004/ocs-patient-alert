import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, db
import io
import msoffcrypto
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

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
    # 예: happyhomeisfun@gmail_com → happyhomeisfun@gmail.com
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

# 📂 엑셀 로드
def load_excel(file, password=None):
    try:
        file.seek(0)
        office_file = msoffcrypto.OfficeFile(file)
        if office_file.is_encrypted():
            if not password:
                raise ValueError("암호화된 파일입니다.")
            decrypted = io.BytesIO()
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            return pd.ExcelFile(decrypted)
        else:
            return pd.ExcelFile(file)
    except Exception as e:
        raise ValueError(f"엑셀 처리 실패: {e}")

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

# 👤 일반 사용자 모드
if user_id != "admin":
    st.subheader("📝 내 환자 등록")
    ref = db.reference(f"patients/{firebase_key}")
    existing_data = ref.get()

    if existing_data:
        for key, val in existing_data.items():
            with st.container():
                col1, col2 = st.columns([0.85, 0.15])
                with col1:
                    st.markdown(f"👤 {val['환자명']} / 🆔 {val['진료번호']}")
                with col2:
                    if st.button("❌ 삭제", key=key):
                        db.reference(f"patients/{firebase_key}/{key}").delete()
                        st.success("삭제 완료")
                        st.rerun()
    else:
        st.info("등록된 환자가 없습니다.")

    with st.form("register_form"):
        name = st.text_input("환자명")
        pid = st.text_input("진료번호")
        submitted = st.form_submit_button("등록")
        if submitted:
            if not name or not pid:
                st.warning("모든 항목을 입력해주세요.")
            elif existing_data and any(
                v["환자명"] == name and v["진료번호"] == pid for v in existing_data.values()):
                st.error("이미 등록된 환자입니다.")
            else:
                ref.push().set({"환자명": name, "진료번호": pid})
                st.success(f"{name} ({pid}) 등록 완료")
                st.rerun()

# 🔑 관리자 모드
else:
    st.subheader("📂 Excel 업로드, 처리, 사용자 알림")

    uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])
    if uploaded_file:
        password = None
        if is_encrypted_excel(uploaded_file):
            password = st.text_input("🔑 Excel 암호를 입력하세요", type="password")
            if not password:
                st.stop()

        try:
            # 📌 Colab 기반 처리 함수 정의
            def process_excel_file(file_obj, password):
                import msoffcrypto
                from openpyxl import load_workbook
                from openpyxl.styles import Font

                sheet_name_mapping = {
                    '교정': '교정', '교정과': '교정', '구강내과': '내과', '내과': '내과',
                    '구강악안면외과': '외과', '외과': '외과', '보존과': '보존', '보존': '보존',
                    '보철과': '보철', '보철': '보철', '소아치과': '소치', '소치': '소치',
                    '원내생진료센터': '원내생', '원내생': '원내생', '원스톱협진센터': '원스톱',
                    '원스톱': '원스톱', '임플란트진료센터': '임플란트', '임플란트': '임플란트',
                    '치주과': '치주', '치주': '치주', '임플실': '임플란트', '원진실': '원내생'
                }

                professors_dict = {
                    '소치': ['김현태', '장기택', '김정욱', '현홍근', '김영재', '신터전', '송지수'],
                    '보존': ['이인복', '금기연', '이우철', '유연지', '서덕규', '이창하', '김선영', '손원준'],
                    '외과': ['최진영', '서병무', '명훈', '김성민', '박주영', '양훈주', '한정준', '권익재'],
                    '치주': ['구영', '이용무', '설양조', '구기태', '김성태', '조영단'],
                    '보철': ['곽재영', '김성균', '임영준', '김명주', '권호범', '여인성', '윤형인', '박지만', '이재현', '조준호'],
                    '교정': [], '내과': [], '원내생': [], '원스톱': [], '임플란트': [],
                }

                def process_sheet(df, professors, key):
                    df = df.drop(columns=['예약일시'], errors='ignore')
                    df = df.sort_values(by=['예약의사', '예약시간'])
                    profs = df[df['예약의사'].isin(professors)]
                    others = df[~df['예약의사'].isin(professors)]
                    if key != '보철':
                        others = others.sort_values(by=['예약시간', '예약의사'])
                    else:
                        others = others.sort_values(by=['예약의사', '예약시간'])

                    rows = []
                    current = None
                    for _, row in others.iterrows():
                        if key != '보철' and current != row['예약시간']:
                            if current is not None:
                                rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
                            current = row['예약시간']
                        elif key == '보철' and current != row['예약의사']:
                            if current is not None:
                                rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
                            current = row['예약의사']
                        rows.append(row)

                    rows += [pd.Series([" "] * len(df.columns), index=df.columns)] * 2
                    rows.append(pd.Series(["<교수님>"] + [" "] * (len(df.columns) - 1), index=df.columns))
                    for _, row in profs.iterrows():
                        rows.append(row)

                    final = pd.DataFrame(rows, columns=df.columns)
                    return final[['진료번호', '예약시간', '환자명', '예약의사', '진료내역']]

                # 📂 복호화
                decrypted = io.BytesIO()
                file = msoffcrypto.OfficeFile(uploaded_file)
                file.load_key(password=password)
                file.decrypt(decrypted)
                decrypted.seek(0)

                wb = load_workbook(decrypted, data_only=True)
                output = io.BytesIO()
                writer = pd.ExcelWriter(output, engine="openpyxl")
                all_dfs = []

                for sheet_name in wb.sheetnames:
                    data = list(wb[sheet_name].values)
                    while data and not any(data[0]):
                        data.pop(0)
                    if len(data) < 2:
                        continue
                    df = pd.DataFrame(data)
                    df.columns = df.iloc[0]
                    df = df.drop([0]).reset_index(drop=True).fillna("").astype(str)
                    df['예약의사'] = df['예약의사'].str.strip().str.replace(" 교수님", "")
                    key = sheet_name_mapping.get(sheet_name.strip())
                    if not key:
                        continue
                    processed = process_sheet(df, professors_dict.get(key, []), key)
                    processed.to_excel(writer, sheet_name=sheet_name, index=False)
                    all_dfs.append(processed)

                writer.close()
                output.seek(0)
                return output, pd.concat(all_dfs, ignore_index=True) if all_dfs else None

            # 📌 실제 실행
            processed_file, full_df = process_excel_file(uploaded_file, password)
            if not full_df.empty:
                st.success("✅ Excel 파일이 성공적으로 처리되었습니다.")
                st.download_button("📥 처리된 파일 다운로드", data=processed_file, file_name="processed_output.xlsx")
            else:
                st.warning("⚠️ 처리된 내용이 없습니다.")
                st.stop()

            # 🔍 Firebase 사용자와 매칭
            users_ref = db.reference("patients")
            all_users = users_ref.get()
            matched_users = []

            for uid, plist in all_users.items():
                registered_set = set((v["환자명"], v["진료번호"]) for v in plist.values())
                matched = full_df[full_df.apply(lambda row: (row["환자명"], row["진료번호"]) in registered_set, axis=1)]
                if not matched.empty:
                    matched_users.append((uid, matched))

            if matched_users:
                st.success(f"📌 {len(matched_users)}명 사용자에게 매칭된 환자가 있습니다.")
                if st.button("📤 메일 보내기"):
                    sender = st.secrets["gmail"]["sender"]
                    sender_pw = st.secrets["gmail"]["app_password"]
                    for uid, df_matched in matched_users:
                        email = recover_email(uid)
                        result = send_email(email, df_matched, sender, sender_pw)
                        if result is True:
                            st.success(f"✅ {email} 전송 완료")
                        else:
                            st.error(f"❌ {email} 전송 실패: {result}")
                else:
                    for uid, df in matched_users:
                        st.markdown(f"#### 🔎 {recover_email(uid)}")
                        st.dataframe(df)
            else:
                st.info("🔍 등록된 사용자와 일치하는 환자가 없습니다.")

        except Exception as e:
            st.error(f"❌ 처리 실패: {e}")
