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
# Firebase 관리자 SDK를 초기화합니다.
# `st.secrets`에서 Firebase 서비스 계정 자격 증명을 가져옵니다.
if not firebase_admin._apps:
    cred = credentials.Certificate(st.secrets["firebase_credentials"])
    firebase_admin.initialize_app(cred, {
        'databaseURL': st.secrets["firebase"]["database_url"]
    })

# 📌 Firebase-safe 경로 변환
# 이메일 주소를 Firebase Realtime Database 경로에 안전하게 사용할 수 있도록 변환합니다.
# '.'는 '_dot_', '@'는 '_at_'으로 대체합니다.
def sanitize_path(email):
    return email.replace(".", "_dot_").replace("@", "_at_")

# 📩 이메일 주소 복원
# Firebase에 저장된 안전한 경로를 원래 이메일 주소로 복원합니다.
def recover_email(safe_id: str) -> str:
    email = safe_id.replace("_at_", "@").replace("_dot_", ".")
    # '.com'으로 끝나는 경우를 위한 특정 처리 (필요에 따라 수정 가능)
    if email.endswith("_com"):
        email = email[:-4] + ".com"
    return email

# 🔒 암호화된 엑셀 여부 확인
# 업로드된 파일이 msoffcrypto 라이브러리로 암호화되었는지 확인합니다.
def is_encrypted_excel(file):
    try:
        file.seek(0) # 파일 포인터를 시작으로 이동
        # msoffcrypto.OfficeFile 객체를 생성하여 파일이 암호화되었는지 확인
        return msoffcrypto.OfficeFile(file).is_encrypted()
    except Exception:
        # 파일이 유효한 Office 파일이 아니거나 암호화 확인 중 오류 발생 시 False 반환
        return False

# 📂 엑셀 로드
# 엑셀 파일을 로드하고, 암호화된 경우 비밀번호로 복호화합니다.
# 복호화된 파일 또는 원본 파일을 BytesIO 객체로 반환합니다.
def load_excel(file, password=None):
    try:
        file.seek(0) # 파일 포인터를 시작으로 이동
        office_file = msoffcrypto.OfficeFile(file)
        if office_file.is_encrypted():
            if not password:
                raise ValueError("암호화된 파일입니다. 비밀번호를 입력해주세요.")
            decrypted = io.BytesIO()
            office_file.load_key(password=password) # 비밀번호로 키 로드
            office_file.decrypt(decrypted) # 파일 복호화
            # Pandas ExcelFile 객체와 복호화된 BytesIO 객체 반환
            return pd.ExcelFile(decrypted), decrypted
        else:
            # 암호화되지 않은 경우, Pandas ExcelFile 객체와 원본 파일 객체 반환
            return pd.ExcelFile(file), file
    except Exception as e:
        raise ValueError(f"엑셀 로드 또는 복호화 실패: {e}")

# 📧 이메일 전송
# 지정된 수신자에게 환자 내원 알림 이메일을 전송합니다.
# `st.secrets`에서 Gmail 발신자 정보와 앱 비밀번호를 사용합니다.
def send_email(receiver, rows, sender, password):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver
        msg['Subject'] = "📌 등록 환자 내원 알림"
        # Pandas DataFrame을 HTML 테이블로 변환하여 이메일 본문에 포함
        html_table = rows.to_html(index=False, escape=False)
        body = f"다음 등록 환자가 내원했습니다:<br><br>{html_table}"
        msg.attach(MIMEText(body, 'html'))

        # SMTP 서버를 통해 이메일 전송
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls() # TLS 암호화 시작
        server.login(sender, password) # 발신자 계정 로그인
        server.send_message(msg) # 메시지 전송
        server.quit() # 서버 연결 종료
        return True
    except Exception as e:
        # 이메일 전송 실패 시 오류 메시지 반환
        return str(e)

# --- 코드 2의 엑셀 처리 관련 상수 및 함수 ---
# 시트 이름 매핑: 엑셀 시트 이름을 표준화된 키로 매핑합니다.
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
    '임플실': '임플란트',
    '원진실': '원내생'
}

# 교수진 사전: 각 시트 키에 해당하는 교수진 목록을 정의합니다.
professors_dict = {
    '소치': ['김현태', '장기택', '김정욱', '현홍근', '김영재', '신터전', '송지수'],
    '보존': ['이인복', '금기연', '이우철', '유연지', '서덕규', '이창하', '김선영', '손원준'],
    '외과': ['최진영', '서병무', '명훈', '김성민', '박주영', '양훈주', '한정준', '권익재'],
    '치주': ['구영', '이용무', '설양조', '구기태', '김성태', '조영단'],
    '보철': ['곽재영', '김성균', '임영준', '김명주', '권호범', '여인성', '윤형인', '박지만', '이재현', '조준호'],
    '교정': [], '내과': [], '원내생': [], '원스톱': [], '임플란트': [],
}

# 📑 엑셀 시트 파싱 및 정제 (코드 2의 process_sheet_v8 함수)
# DataFrame을 정렬하고 교수/비교수 데이터를 분리하여 특정 형식으로 재구성합니다.
def process_sheet_v8(df, professors_list, sheet_key):
    # '예약일시' 컬럼이 있으면 삭제합니다.
    df = df.drop(columns=['예약일시'], errors='ignore')
    # 필수 컬럼 ('예약의사', '예약시간')이 존재하는지 확인합니다.
    if '예약의사' not in df.columns or '예약시간' not in df.columns:
        st.error(f"시트 처리 오류: '예약의사' 또는 '예약시간' 컬럼이 DataFrame에 없습니다.")
        # 필수 컬럼이 없는 경우 빈 DataFrame을 반환하여 오류 확산을 방지합니다.
        return pd.DataFrame(columns=['진료번호', '예약시간', '환자명', '예약의사', '진료내역'])

    # '예약의사'와 '예약시간'을 기준으로 정렬합니다.
    df = df.sort_values(by=['예약의사', '예약시간'])
    # 교수진 목록에 포함된 의사와 그렇지 않은 의사로 DataFrame을 분리합니다.
    professors = df[df['예약의사'].isin(professors_list)]
    non_professors = df[~df['예약의사'].isin(professors_list)]

    # '보철' 시트가 아닌 경우 '예약시간'을 기준으로, '보철' 시트인 경우 '예약의사'를 기준으로 정렬합니다.
    if sheet_key != '보철':
        non_professors = non_professors.sort_values(by=['예약시간', '예약의사'])
    else:
        non_professors = non_professors.sort_values(by=['예약의사', '예약시간'])

    final_rows = []
    current_time = None
    current_doctor = None

    # 비(非)교수 데이터를 처리하고 시간/의사 변경 시 빈 행을 추가합니다.
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

    # 빈 행과 '<교수님>' 헤더를 추가합니다.
    final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
    final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
    final_rows.append(pd.Series(["<교수님>"] + [" "] * (len(df.columns) - 1), index=df.columns))

    current_professor = None
    # 교수 데이터를 처리하고 의사 변경 시 빈 행을 추가합니다.
    for _, row in professors.iterrows():
        if current_professor != row['예약의사']:
            if current_professor is not None:
                final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
            current_professor = row['예약의사']
        final_rows.append(row)

    # 최종 DataFrame을 생성하고 필요한 컬럼만 선택하여 순서를 맞춥니다.
    final_df = pd.DataFrame(final_rows, columns=df.columns)
    required_cols = ['진료번호', '예약시간', '환자명', '예약의사', '진료내역']
    final_df = final_df[[col for col in required_cols if col in final_df.columns]]
    return final_df

# 엑셀 파일 처리 및 스타일링 (코드 2의 process_excel_file을 Streamlit에 맞게 수정)
# 업로드된 엑셀 파일을 복호화하고, 시트별로 데이터를 처리 및 정렬한 후,
# 특정 조건에 따라 스타일을 적용한 새 엑셀 파일을 BytesIO 객체로 반환합니다.
def process_excel_file_and_style(file_bytes_io, password):
    decrypted_file_io = io.BytesIO()
    try:
        file_bytes_io.seek(0) # 파일 포인터를 시작으로 이동
        office_file = msoffcrypto.OfficeFile(file_bytes_io)
        if office_file.is_encrypted():
            if not password:
                raise ValueError("암호화된 파일입니다. 비밀번호를 입력해주세요.")
            office_file.load_key(password=password)
            office_file.decrypt(decrypted_file_io)
        else:
            # 암호화되지 않은 경우, 원본 파일 내용을 복사합니다.
            decrypted_file_io.write(file_bytes_io.read())
        decrypted_file_io.seek(0) # 복호화된 파일 포인터를 시작으로 이동
    except Exception as e:
        raise ValueError(f"파일 복호화 또는 로드 실패: {e}")

    try:
        # 복호화된(또는 원본) BytesIO 객체로부터 워크북을 로드합니다.
        wb_raw = load_workbook(filename=decrypted_file_io, data_only=True)
    except Exception as e:
        raise ValueError(f"엑셀 워크북 로드 실패: {e}")

    processed_sheets_dfs = {} # 처리된 DataFrame을 저장할 딕셔너리

    for sheet_name in wb_raw.sheetnames:
        ws = wb_raw[sheet_name]
        values = list(ws.values)
        # 시트 상단의 빈 행을 제거합니다.
        while values and (values[0] is None or all(v is None for v in values[0])):
            values.pop(0)
        # 헤더와 최소 한 줄의 데이터가 있는지 확인합니다.
        if len(values) < 2:
            st.warning(f"⚠ 시트 '{sheet_name}'에 유효한 데이터가 충분하지 않습니다. 건너뜁니다.")
            continue

        df = pd.DataFrame(values)
        df.columns = df.iloc[0] # 첫 번째 행을 컬럼 헤더로 설정
        df = df.drop([0]).reset_index(drop=True) # 헤더 행을 데이터에서 제거
        df = df.fillna("").astype(str) # NaN 값을 빈 문자열로 채우고 모든 데이터를 문자열로 변환
        
        # '예약의사' 컬럼 전처리: 공백 제거 및 " 교수님" 문자열 제거
        if '예약의사' in df.columns:
            df['예약의사'] = df['예약의사'].str.strip().str.replace(" 교수님", "", regex=False)
        else:
            st.warning(f"시트 '{sheet_name}': '예약의사' 컬럼이 없습니다. 이 시트는 처리되지 않습니다.")
            continue

        sheet_key = sheet_name_mapping.get(sheet_name.strip(), None)
        if not sheet_key:
            st.warning(f"❌ 시트 '{sheet_name}'을 인식할 수 없습니다. 건너뜁니다.")
            continue

        professors_list = professors_dict.get(sheet_key, [])
        try:
            # `process_sheet_v8` 함수를 사용하여 시트 데이터 처리
            processed_df = process_sheet_v8(df, professors_list, sheet_key)
            processed_sheets_dfs[sheet_name] = processed_df
        except KeyError as e:
            st.error(f"시트 '{sheet_name}' 처리 중 컬럼 오류: {e}. 이 시트는 건너뜁니다.")
            continue
        except Exception as e:
            st.error(f"시트 '{sheet_name}' 처리 중 알 수 없는 오류: {e}. 이 시트는 건너뜁니다.")
            continue

    if not processed_sheets_dfs:
        st.info("❌ 처리된 시트가 없습니다.")
        return None, None # 처리된 시트가 없으면 None 반환

    # 처리된 DataFrame들을 메모리 내 엑셀 파일로 작성하여 스타일링을 적용합니다.
    output_buffer_for_styling = io.BytesIO()
    with pd.ExcelWriter(output_buffer_for_styling, engine='openpyxl') as writer:
        for sheet_name, df in processed_sheets_dfs.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    output_buffer_for_styling.seek(0) # 파일 포인터를 시작으로 이동
    wb_styled = load_workbook(output_buffer_for_styling) # 스타일링을 위해 워크북 다시 로드

    # 스타일링 적용
    for sheet_name in wb_styled.sheetnames:
        ws = wb_styled[sheet_name]
        # 헤더 행의 컬럼 이름을 기반으로 인덱스를 매핑합니다.
        header = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}

        # 데이터 행을 순회하며 스타일을 적용합니다. (헤더 다음 행부터 시작)
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
            # '<교수님>' 행의 모든 셀을 볼드 처리합니다.
            if row[0].value == "<교수님>":
                for cell in row:
                    if cell.value:
                        cell.font = Font(bold=True)

            # '교정' 시트에서 '진료내역' 컬럼에 '본딩' 또는 'bonding'이 포함된 경우 볼드 처리합니다.
            if sheet_name.strip() == "교정" and '진료내역' in header:
                idx = header['진료내역'] - 1 # 0-기반 인덱스
                # 셀이 존재하는지 확인 후 접근
                if len(row) > idx:
                    cell = row[idx]
                    text = str(cell.value)
                    if any(keyword in text for keyword in ['본딩', 'bonding']):
                        cell.font = Font(bold=True)

    final_output_bytes = io.BytesIO()
    wb_styled.save(final_output_bytes) # 스타일링된 워크북을 BytesIO에 저장
    final_output_bytes.seek(0) # 파일 포인터를 시작으로 이동

    # 처리된 DataFrame 딕셔너리와 스타일링된 엑셀 파일의 BytesIO 객체를 모두 반환합니다.
    return processed_sheets_dfs, final_output_bytes

# --- Streamlit 애플리케이션 시작 ---
st.title("🩺 환자 내원 확인 시스템")

# 사용자 아이디 입력 필드
user_id = st.text_input("아이디를 입력하세요")
if not user_id:
    st.stop() # 아이디가 입력되지 않으면 애플리케이션 실행 중지

# Firebase 경로에 사용할 안전한 키 생성
firebase_key = sanitize_path(user_id)

# 👤 사용자 모드 (admin이 아닌 경우)
if user_id != "admin":
    st.subheader("📝 내 환자 등록")
    ref = db.reference(f"patients/{firebase_key}") # Firebase 참조 설정
    existing_data = ref.get() # Firebase에서 기존 환자 데이터 가져오기

    if existing_data:
        # 등록된 환자 목록을 표시하고 삭제 버튼 제공
        for key, val in existing_data.items():
            with st.container():
                col1, col2 = st.columns([0.85, 0.15])
                with col1:
                    st.markdown(f"👤 {val['환자명']} / 🆔 {val['진료번호']}")
                with col2:
                    if st.button("❌ 삭제", key=key):
                        db.reference(f"patients/{firebase_key}/{key}").delete() # Firebase에서 환자 삭제
                        st.success("삭제 완료")
                        st.rerun() # 변경 사항 반영을 위해 앱 다시 실행
    else:
        st.info("등록된 환자가 없습니다.")

    # 새 환자 등록 폼
    with st.form("register_form"):
        name = st.text_input("환자명")
        pid = st.text_input("진료번호")
        submitted = st.form_submit_button("등록")
        if submitted:
            if not name or not pid:
                st.warning("모든 항목을 입력해주세요.")
            # 이미 등록된 환자인지 확인
            elif existing_data and any(
                v["환자명"] == name and v["진료번호"] == pid for v in existing_data.values()):
                st.error("이미 등록된 환자입니다.")
            else:
                ref.push().set({"환자명": name, "진료번호": pid}) # Firebase에 새 환자 등록
                st.success(f"{name} ({pid}) 등록 완료")
                st.rerun() # 변경 사항 반영을 위해 앱 다시 실행

# 🔑 관리자 모드 (admin으로 로그인한 경우)
else:
    st.subheader("📂 엑셀 업로드 및 사용자 일치 검사")
    # 엑셀 파일 업로드 위젯
    uploaded_file = st.file_uploader("암호화된 Excel 파일을 업로드하세요", type=["xlsx", "xlsm"])

    if uploaded_file:
        password = None
        # 업로드된 파일이 암호화되었는지 확인하고 비밀번호 입력 필드를 표시
        if is_encrypted_excel(uploaded_file):
            password = st.text_input("🔑 엑셀 파일 비밀번호 입력", type="password")
            if not password:
                st.info("암호화된 파일입니다. 비밀번호를 입력해주세요.")
                st.stop() # 비밀번호가 입력될 때까지 실행 중지

        try:
            # 엑셀 파일을 로드하고 (필요시 복호화), 원본/복호화된 파일 객체를 얻습니다.
            xl_object, raw_file_io = load_excel(uploaded_file, password)

            # 코드 2의 로직을 사용하여 엑셀 파일을 처리하고 스타일을 적용합니다.
            # 처리된 DataFrame 딕셔너리와 스타일링된 엑셀 파일의 BytesIO 객체를 반환받습니다.
            excel_data_dfs, styled_excel_bytes = process_excel_file_and_style(raw_file_io, password)

            if excel_data_dfs is None or styled_excel_bytes is None:
                st.warning("엑셀 파일 처리 중 문제가 발생했거나 처리할 데이터가 없습니다.")
                st.stop()

            # Gmail 발신자 정보 가져오기
            sender = st.secrets["gmail"]["sender"]
            sender_pw = st.secrets["gmail"]["app_password"]

            users_ref = db.reference("patients") # 모든 환자 데이터에 대한 Firebase 참조
            all_users = users_ref.get() # 모든 등록된 환자 데이터 가져오기
            if not all_users:
                st.warning("❗ 등록된 사용자가 없습니다.")
                st.stop()

            matched_users = [] # 엑셀 데이터와 일치하는 환자를 가진 사용자 목록

            # Firebase에 등록된 모든 사용자를 순회합니다.
            for uid, plist in all_users.items():
                # 각 사용자가 등록한 환자 정보를 (환자명, 진료번호) 튜플의 집합으로 만듭니다.
                registered_set = set(
                    (v["환자명"].strip(), v["진료번호"].strip().zfill(8)) for v in plist.values()
                )
                matched_rows_for_user = [] # 현재 사용자와 일치하는 엑셀 행 목록

                # 처리된 엑셀 데이터의 각 시트(DataFrame)를 순회합니다.
                for sheet_name, df_sheet in excel_data_dfs.items():
                    # 엑셀 시트의 각 행을 순회하며 등록된 환자와 일치하는지 확인합니다.
                    matched = df_sheet[df_sheet.apply(
                        lambda row: (row["환자명"].strip(), row["진료번호"].strip().zfill(8)) in registered_set, axis=1
                    )]
                    if not matched.empty:
                        matched["시트"] = sheet_name # 일치하는 행에 시트 이름 추가
                        matched_rows_for_user.append(matched)

                if matched_rows_for_user:
                    # 현재 사용자와 일치하는 모든 행을 하나의 DataFrame으로 결합합니다.
                    combined_matched_df = pd.concat(matched_rows_for_user, ignore_index=True)
                    matched_users.append((uid, combined_matched_df)) # 일치하는 사용자 목록에 추가

            if matched_users:
                st.success(f"🔍 {len(matched_users)}명의 사용자와 일치하는 환자 발견됨.")

                # 일치하는 환자 데이터를 각 사용자별로 표시합니다.
                for uid, df_matched in matched_users:
                    st.markdown(f"### 📧 {recover_email(uid)}")
                    st.dataframe(df_matched)

                # 메일 전송 버튼
                if st.button("📤 메일 보내기"):
                    for uid, df_matched in matched_users:
                        real_email = recover_email(uid)
                        result = send_email(real_email, df_matched, sender, sender_pw)
                        if result is True:
                            st.success(f"✅ {real_email} 전송 완료")
                        else:
                            st.error(f"❌ {real_email} 전송 실패: {result}")

                # 처리된 엑셀 파일 다운로드 버튼
                output_filename = uploaded_file.name.replace(".xlsx", "_processed.xlsx").replace(".xlsm", "_processed.xlsx")
                st.download_button(
                    "📥 처리된 엑셀 다운로드",
                    data=styled_excel_bytes, # 스타일링이 적용된 엑셀 파일의 BytesIO 객체 사용
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("📭 매칭된 사용자 없음")

        except ValueError as ve:
            st.error(f"❌ 파일 처리 실패: {ve}")
        except Exception as e:
            st.error(f"❌ 예상치 못한 오류 발생: {e}")
