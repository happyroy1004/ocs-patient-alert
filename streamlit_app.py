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
import re
import json

# --- 이메일 유효성 검사 함수 ---
def is_valid_email(email):
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(email_regex, email) is not None

# Firebase 초기화
if not firebase_admin._apps:
    try:
        firebase_credentials_json_str = st.secrets["FIREBASE_SERVICE_ACCOUNT_JSON"]
        firebase_credentials_dict = json.loads(firebase_credentials_json_str)

        cred = credentials.Certificate(firebase_credentials_dict)
        firebase_admin.initialize_app(cred, {
            'databaseURL': st.secrets["FIREBASE_DATABASE_URL"]
        })
    except Exception as e:
        st.error(f"Firebase 초기화 오류: {e}")
        st.info("secrets.toml 파일의 Firebase 설정(FIREBASE_SERVICE_ACCOUNT_JSON 또는 FIREBASE_DATABASE_URL)을 확인해주세요.")
        st.stop()

# Firebase-safe 경로 변환 (이메일을 Firebase 키로 사용하기 위해)
def sanitize_path(email):
    return email.replace(".", "_dot_").replace("@", "_at_")

# 이메일 주소 복원 (Firebase 안전 키에서 원래 이메일로)
def recover_email(safe_id: str) -> str:
    email = safe_id.replace("_at_", "@").replace("_dot_", ".")
    # '.com'이 '_com'으로 변환된 경우를 처리 (필요한 경우에만)
    if not email.endswith(".com") and email.endswith("_com"):
        email = email[:-4] + ".com"
    return email

# 암호화된 엑셀 파일인지 확인
def is_encrypted_excel(file):
    try:
        file.seek(0)
        return msoffcrypto.OfficeFile(file).is_encrypted()
    except Exception:
        return False

# 엑셀 파일 로드 및 복호화
def load_excel(file, password=None):
    try:
        file.seek(0)
        office_file = msoffcrypto.OfficeFile(file)
        if office_file.is_encrypted():
            if not password:
                raise ValueError("암호화된 파일입니다. 비밀번호를 입력해주세요.")
            decrypted = io.BytesIO()
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            return pd.ExcelFile(decrypted), decrypted
        else:
            return pd.ExcelFile(file), file
    except Exception as e:
        raise ValueError(f"엑셀 로드 또는 복호화 실패: {e}")

# 이메일 전송 함수
def send_email(receiver, rows, sender, password, date_str=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver

        subject_prefix = ""
        if date_str:
            subject_prefix = f"{date_str}일에 내원하는 "
        msg['Subject'] = f"{subject_prefix}등록 환자 내원 알림"

        html_table = rows.to_html(index=False, escape=False)

        style = """
        <style>
            table {
                width: 100%;
                max-width: 100%;
                border-collapse: collapse;
                font-family: Arial, sans-serif;
                font-size: 14px;
                table-layout: fixed;
            }
            th, td {
                border: 1px solid #dddddd;
                text-align: left;
                padding: 8px;
                vertical-align: top;
                word-wrap: break-word;
                word-break: break-word;
            }
            th {
                background-color: #f2f2f2;
                font-weight: bold;
                white-space: nowrap;
            }
            tr:nth-child(even) {
                background-color: #f9f9f9;
            }
            .table-container {
                overflow-x: auto;
                -webkit-overflow-scrolling: touch;
            }
        </style>
        """

        body = f"다음 등록 환자가 내원했습니다:<br><br><div class='table-container'>{style}{html_table}</div>"
        msg.attach(MIMEText(body, 'html'))

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        return str(e)

# --- 엑셀 처리 관련 상수 및 함수 ---
# 시트 이름 키워드와 해당 진료과 매핑
sheet_keyword_to_department_map = {
    '치과보철과': '보철', '보철과': '보철', '보철': '보철',
    '치과교정과' : '교정', '교정과': '교정', '교정': '교정',
    '구강 악안면외과' : '외과', '구강악안면외과': '외과', '외과': '외과',
    '구강 내과' : '내과', '구강내과': '내과', '내과': '내과',
    '치과보존과' : '보존', '보존과': '보존', '보존': '보존',
    '소아치과': '소치', '소치': '소치',
    '원내생진료센터': '원내생', '원내생': '원내생',
    '원스톱 협진센터' : '원스톱', '원스톱협진센터': '원스톱', '원스톱': '원스톱',
    '임플란트 진료센터' : '임플란트', '임플란트진료센터': '임플란트', '임플란트': '임플란트',
    '임플' : '임플란트', '치주과': '치주', '치주': '치주',
    '임플실': '임플란트', '원진실': '원내생', '병리': '병리'
}

# 각 진료과별 교수님 명단 (엑셀 시트 정렬에 사용)
professors_dict = {
    '소치': ['김현태', '장기택', '김정욱', '현홍근', '김영재', '신터전', '송지수'],
    '보존': ['이인복', '금기연', '이우철', '유연지', '서덕규', '이창하', '김선영', '손원준'],
    '외과': ['최진영', '서병무', '명훈', '김성민', '박주영', '양훈주', '한정준', '권익재'],
    '치주': ['구영', '이용무', '설양조', '구기태', '김성태', '조영단'],
    '보철': ['곽재영', '김성균', '임영준', '김명주', '권호범', '여인성', '윤형인', '박지만', '이재현', '조준호'],
    '교정': [], '내과': [], '원내생': [], '원스톱': [], '임플란트': [], '병리': []
}

# 엑셀 시트 데이터 처리 (교수님/비교수님, 시간/의사별 정렬)
def process_sheet_v8(df, professors_list, sheet_key):
    # '예약일시' 컬럼이 있으면 삭제
    df = df.drop(columns=['예약일시'], errors='ignore')
    # 필수 컬럼 확인
    if '예약의사' not in df.columns or '예약시간' not in df.columns:
        st.error(f"시트 처리 오류: '예약의사' 또는 '예약시간' 컬럼이 DataFrame에 없습니다.")
        return pd.DataFrame(columns=['진료번호', '예약시간', '환자명', '예약의사', '진료내역'])

    df = df.sort_values(by=['예약의사', '예약시간']) # 기본 정렬
    professors = df[df['예약의사'].isin(professors_list)] # 교수님 데이터 분리
    non_professors = df[~df['예약의사'].isin(professors_list)] # 교수님 아닌 데이터 분리

    # 진료과에 따른 추가 정렬 (보철과만 특이)
    if sheet_key != '보철':
        non_professors = non_professors.sort_values(by=['예약시간', '예약의사'])
    else:
        non_professors = non_professors.sort_values(by=['예약의사', '예약시간'])

    final_rows = []
    current_time = None
    current_doctor = None

    # 교수님 아닌 데이터 처리 (빈 줄 삽입 로직)
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
                current_doctor = row['예 예약의사']
        final_rows.append(row)

    # 교수님 데이터 처리 전 구분선 및 "<교수님>" 표기
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
    # 필요한 컬럼만 선택하여 반환
    required_cols = ['진료번호', '예약시간', '환자명', '예약의사', '진료내역']
    final_df = final_df[[col for col in required_cols if col in final_df.columns]]
    return final_df

# 엑셀 파일 전체 처리 및 스타일 적용
def process_excel_file_and_style(file_bytes_io):
    file_bytes_io.seek(0)

    try:
        wb_raw = load_workbook(filename=file_bytes_io, data_only=True)
    except Exception as e:
        raise ValueError(f"엑셀 워크북 로드 실패: {e}")

    processed_sheets_dfs = {}

    for sheet_name_raw in wb_raw.sheetnames:
        sheet_name_lower = sheet_name_raw.strip().lower()

        sheet_key = None
        # 시트 이름을 기반으로 진료과 매핑
        for keyword, department_name in sorted(sheet_keyword_to_department_map.items(), key=lambda item: len(item[0]), reverse=True):
            if keyword.lower() in sheet_name_lower:
                sheet_key = department_name
                break

        if not sheet_key:
            st.warning(f"시트 '{sheet_name_raw}'을(를) 인식할 수 없습니다. 건너킵니다.")
            continue

        ws = wb_raw[sheet_name_raw]
        values = list(ws.values)
        # 빈 상단 행 제거
        while values and (values[0] is None or all((v is None or str(v).strip() == "") for v in values[0])):
            values.pop(0)
        if len(values) < 2:
            st.warning(f"시트 '{sheet_name_raw}'에 유효한 데이터가 충분하지 않습니다. 건너킵니다.")
            continue

        df = pd.DataFrame(values)
        df.columns = df.iloc[0] # 첫 행을 컬럼명으로
        df = df.drop([0]).reset_index(drop=True) # 첫 행 삭제 및 인덱스 재설정
        df = df.fillna("").astype(str) # NaN 값 채우고 모든 컬럼을 문자열로

        if '예약의사' in df.columns:
            df['예약의사'] = df['예약의사'].str.strip().str.replace(" 교수님", "", regex=False)
        else:
            st.warning(f"시트 '{sheet_name_raw}': '예약의사' 컬럼이 없습니다. 이 시트는 처리되지 않습니다.")
            continue

        professors_list = professors_dict.get(sheet_key, [])
        try:
            processed_df = process_sheet_v8(df, professors_list, sheet_key)
            processed_sheets_dfs[sheet_name_raw] = processed_df
        except KeyError as e:
            st.error(f"시트 '{sheet_name_raw}' 처리 중 컬럼 오류: {e}. 이 시트는 건너킵니다.")
            continue
        except Exception as e:
            st.error(f"시트 '{sheet_name_raw}' 처리 중 알 수 없는 오류: {e}. 이 시트는 건너킵니다.")
            continue

    if not processed_sheets_dfs:
        st.info("처리된 시트가 없습니다.")
        return None, None

    # 스타일 적용을 위해 처리된 데이터를 다시 엑셀로 저장 (메모리 내에서)
    output_buffer_for_styling = io.BytesIO()
    with pd.ExcelWriter(output_buffer_for_styling, engine='openpyxl') as writer:
        for sheet_name_raw, df in processed_sheets_dfs.items():
            df.to_excel(writer, sheet_name=sheet_name_raw, index=False)

    output_buffer_for_styling.seek(0)
    wb_styled = load_workbook(output_buffer_for_styling)

    # 각 시트에 스타일 적용
    for sheet_name in wb_styled.sheetnames:
        ws = wb_styled[sheet_name]
        header = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}

        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
            # 교수님 섹션 글씨 진하게
            if row[0].value == "<교수님>":
                for cell in row:
                    if cell.value:
                        cell.font = Font(bold=True)

            # 교정 시트의 '진료내역'에 특정 키워드 포함 시 글씨 진하게
            if sheet_name.strip() == "교정" and '진료내역' in header:
                idx = header['진료내역'] - 1
                if len(row) > idx:
                    cell = row[idx]
                    text = str(cell.value)
                    if any(keyword in text for keyword in ['본딩', 'bonding']):
                        cell.font = Font(bold=True)

    final_output_bytes = io.BytesIO()
    wb_styled.save(final_output_bytes)
    final_output_bytes.seek(0)

    return processed_sheets_dfs, final_output_bytes

# --- Streamlit 애플리케이션 시작 ---
st.title("환자 내원 확인 시스템") # 기존 제목
st.markdown("---") # 구분선 추가
st.markdown("<p style='text-align: left; color: grey; font-size: small;'>directed by HSY</p>", unsafe_allow_html=True) # 왼쪽 정렬, 작은 글씨

# 사용자 입력 필드
user_name = st.text_input("사용자 이름을 입력하세요 (예시: 김민지)")
user_id = st.text_input("아이디를 입력하세요 (예시: example@gmail.com)")

# Admin 계정 확인 로직 (이름과 아이디 모두 'admin'일 경우)
is_admin_mode = (user_name.strip().lower() == "admin" and user_id.strip().lower() == "admin")

# 입력 유효성 검사 및 초기 안내
if user_id and user_name:
    # Admin 모드가 아닐 경우에만 이메일 형식 검사
    if not is_admin_mode and not is_valid_email(user_id):
        st.error("올바른 이메일 주소 형식이 아닙니다. 'user@example.com'과 같이 입력해주세요.")
        st.stop()
elif not user_id or not user_name:
    st.info("내원 알람 노티를 받을 사용자의 이름과 이메일 주소를 입력해주세요.")
    st.stop()

# Firebase 경로에 사용할 안전한 키 생성 (Admin 계정은 실제 Firebase 키로 사용되지 않음)
firebase_key = sanitize_path(user_id)

# Firebase 데이터베이스 참조 설정
users_ref = db.reference("users") # 사용자 이름 등 메타 정보 저장용
# Admin 모드가 아닐 경우에만 해당 사용자의 환자 정보 참조
if not is_admin_mode:
    patients_ref_for_user = db.reference(f"patients/{firebase_key}")

# 사용자 정보 (이름, 이메일) Firebase 'users' 노드에 저장 또는 업데이트
# Admin 계정일 때는 이 과정 건너뛰기
if not is_admin_mode:
    current_user_meta_data = users_ref.child(firebase_key).get()
    # 사용자 정보가 없거나, 현재 입력된 이름/이메일과 다르면 업데이트
    if not current_user_meta_data or current_user_meta_data.get("name") != user_name or current_user_meta_data.get("email") != user_id:
        users_ref.child(firebase_key).update({"name": user_name, "email": user_id})
        st.success(f"사용자 정보가 업데이트되었습니다: {user_name} ({user_id})")

# --- 사용자 모드 (Admin이 아닌 경우) ---
if not is_admin_mode:
    st.subheader(f"{user_name}님의 등록 환자 목록") # 사용자 이름 표시

    # 해당 사용자의 기존 환자 데이터 로드
    existing_patient_data = patients_ref_for_user.get()

    if existing_patient_data:
        for key, val in existing_patient_data.items():
            with st.container():
                col1, col2 = st.columns([0.85, 0.15])
                with col1:
                    department_display = val.get('등록과', '미지정')
                    st.markdown(f"환자명: {val['환자명']} / 진료번호: {val['진료번호']} / 등록과: {department_display}")
                with col2:
                    if st.button("삭제", key=key): # 각 항목마다 고유한 삭제 버튼 키
                        patients_ref_for_user.child(key).delete()
                        st.success("환자가 성공적으로 삭제되었습니다.")
                        st.rerun() # 삭제 후 화면 새로고침

    else:
        st.info("등록된 환자가 없습니다.")

    # 환자 등록 폼
    with st.form("register_form"):
        name = st.text_input("환자명")
        pid = st.text_input("진료번호")

        # 등록 가능한 진료과 목록 생성 및 선택 박스
        departments_for_registration = sorted(list(set(sheet_keyword_to_department_map.values())))
        selected_department = st.selectbox("등록 과", departments_for_registration)

        submitted = st.form_submit_button("등록")
        if submitted:
            if not name or not pid:
                st.warning("모든 항목을 입력해주세요.")
            # 중복 환자 등록 방지
            elif existing_patient_data and any(
                v["환자명"] == name and v["진료번호"] == pid and v.get("등록과") == selected_department
                for v in existing_patient_data.values()):
                st.error("이미 등록된 환자입니다.")
            else:
                # Firebase에 환자 정보 저장
                patients_ref_for_user.push().set({"환자명": name, "진료번호": pid, "등록과": selected_department})
                st.success(f"{name} ({pid}) [{selected_department}] 환자 등록 완료")
                st.rerun() # 등록 후 화면 새로고침

# --- 관리자 모드 (Admin인 경우) ---
else:
    st.subheader("엑셀 처리 및 내원 확인 알림 보내기 (관리자 모드)")
    uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])

    if uploaded_file:
        password = None
        # 파일이 암호화되어 있으면 비밀번호 입력 필드 표시
        if is_encrypted_excel(uploaded_file):
            password = st.text_input("엑셀 파일 비밀번호 입력", type="password")
            if not password:
                st.info("암호화된 파일입니다. 비밀번호를 입력해주세요.")
                st.stop()

        try:
            # 파일 이름에서 날짜 추출
            file_name = uploaded_file.name
            date_match = re.search(r'(\d{4})', file_name)
            extracted_date = date_match.group(1) if date_match else None

            # 엑셀 파일 로드 및 처리
            xl_object, raw_file_io = load_excel(uploaded_file, password)
            excel_data_dfs, styled_excel_bytes = process_excel_file_and_style(raw_file_io)

            if excel_data_dfs is None or styled_excel_bytes is None:
                st.warning("엑셀 파일 처리 중 문제가 발생했거나 처리할 데이터가 없습니다.")
                st.stop()

            # 이메일 전송을 위한 발신자 정보 (secrets.toml에서 로드)
            sender = st.secrets["GMAIL_SENDER"]
            sender_pw = st.secrets["GMAIL_APP_PASSWORD"]

            # Firebase에서 모든 사용자 메타 정보 및 모든 환자 데이터 로드
            all_users_meta = users_ref.get()
            all_patients_data = db.reference("patients").get()

            # 데이터 로드 여부에 따른 안내
            if not all_users_meta and not all_patients_data:
                st.warning("Firebase에 등록된 사용자 또는 환자 데이터가 없습니다. 이메일 전송은 불가능합니다.")
            elif not all_users_meta:
                st.warning("Firebase users 노드에 등록된 사용자 메타 정보가 없습니다. 이메일 전송 시 이름 대신 이메일이 사용됩니다.")
            elif not all_patients_data:
                st.warning("Firebase patients 노드에 등록된 환자 데이터가 없습니다. 매칭할 수 없습니다.")

            matched_users = []

            if all_patients_data: # 환자 데이터가 있어야 매칭 로직 실행
                # 모든 환자 데이터를 순회하며 매칭
                for uid_safe, registered_patients_for_this_user in all_patients_data.items():
                    user_email = recover_email(uid_safe) # Firebase 키에서 이메일 복원
                    user_display_name = user_email # 기본 표시 이름은 이메일

                    # users 노드에서 사용자 이름 정보 가져오기
                    if all_users_meta and uid_safe in all_users_meta:
                        user_meta = all_users_meta[uid_safe]
                        if "name" in user_meta:
                            user_display_name = user_meta["name"]
                        if "email" in user_meta:
                            user_email = user_meta["email"] # users 노드에 저장된 실제 이메일 사용

                    registered_patients_data = []
                    if registered_patients_for_this_user:
                        for key, val in registered_patients_for_this_user.items():
                            registered_patients_data.append({
                                "환자명": val["환자명"].strip(),
                                "진료번호": val["진료번호"].strip().zfill(8),
                                "등록과": val.get("등록과", "")
                            })

                    matched_rows_for_user = []

                    # 엑셀 시트별로 매칭 진행
                    for sheet_name_excel_raw, df_sheet in excel_data_dfs.items():
                        excel_sheet_name_lower = sheet_name_excel_raw.strip().lower()

                        excel_sheet_department = None
                        for keyword, department_name in sorted(sheet_keyword_to_department_map.items(), key=lambda item: len(item[0]), reverse=True):
                            if keyword.lower() in excel_sheet_name_lower:
                                excel_sheet_department = department_name
                                break

                        if not excel_sheet_department:
                            continue

                        for _, excel_row in df_sheet.iterrows():
                            excel_patient_name = excel_row["환자명"].strip()
                            excel_patient_pid = excel_row["진료번호"].strip().zfill(8)

                            # 등록된 환자 정보와 엑셀 데이터 매칭
                            for registered_patient in registered_patients_data:
                                if (registered_patient["환자명"] == excel_patient_name and
                                    registered_patient["진료번호"] == excel_patient_pid and
                                    registered_patient["등록과"] == excel_sheet_department):

                                    matched_row_copy = excel_row.copy()
                                    matched_row_copy["시트"] = sheet_name_excel_raw
                                    matched_rows_for_user.append(matched_row_copy)
                                    break

                    if matched_rows_for_user:
                        combined_matched_df = pd.DataFrame(matched_rows_for_user)
                        matched_users.append({"email": user_email, "name": user_display_name, "data": combined_matched_df})

            # 매칭 결과 표시 및 이메일 전송 버튼
            if matched_users:
                st.success(f"{len(matched_users)}명의 사용자와 일치하는 환자 발견됨.")

                for user_match_info in matched_users:
                    st.markdown(f"**수신자:** {user_match_info['name']} ({user_match_info['email']})")
                    st.dataframe(user_match_info['data'])

                if st.button("매칭된 환자에게 메일 보내기"):
                    for user_match_info in matched_users:
                        real_email = user_match_info['email']
                        df_matched = user_match_info['data']
                        result = send_email(real_email, df_matched, sender, sender_pw, date_str=extracted_date)
                        if result is True:
                            st.success(f"**{user_match_info['name']}** ({real_email}) 전송 완료")
                        else:
                            st.error(f"**{user_match_info['name']}** ({real_email}) 전송 실패: {result}")
            else:
                st.info("엑셀 파일 처리 완료. 매칭된 환자가 없습니다.")

            # 처리된 엑셀 파일 다운로드 버튼
            output_filename = uploaded_file.name.replace(".xlsx", "_processed.xlsx").replace(".xlsm", "_processed.xlsm")
            st.download_button(
                "처리된 엑셀 다운로드",
                data=styled_excel_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except ValueError as ve:
            st.error(f"파일 처리 실패: {ve}")
        except Exception as e:
            st.error(f"예상치 못한 오류 발생: {e}")
