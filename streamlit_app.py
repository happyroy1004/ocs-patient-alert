import json
import streamlit as st
import pandas as pd
import msoffcrypto
import io
from openpyxl import load_workbook
from openpyxl.styles import Font
import firebase_admin
from firebase_admin import credentials, firestore

# st.secrets["FIREBASE_KEY"]는 SectionProxy이므로 dict로 변환
firebase_config = dict(st.secrets["FIREBASE_KEY"])

# credentials.Certificate()에 dict 그대로 전달
cred = credentials.Certificate(firebase_config)
firebase_admin.initialize_app(cred)
db = firestore.client()

st.title("🔒 OCS 환자 알림 시스템")

# 사용자 이메일
user_email = st.text_input("📧 이메일을 입력하세요:")
if not user_email:
    st.stop()

# 환자 등록
st.subheader("📝 환자 등록")
name_input = st.text_input("환자 이름")
id_input = st.text_input("환자 번호")
if st.button("환자 등록") and name_input and id_input:
    doc_ref = db.collection("users").document(user_email)
    doc_ref.set({
        "patients": firestore.ArrayUnion([{
            "name": name_input.strip(),
            "id": id_input.strip()
        }])
    }, merge=True)
    st.success("환자 등록 완료!")

# 환자 목록 표시
st.subheader("📋 등록된 환자 목록")
doc = db.collection("users").document(user_email).get()
user_patients = doc.to_dict().get("patients", []) if doc.exists else []
if user_patients:
    for p in user_patients:
        st.write(f"👤 {p['name']} ({p['id']})")
else:
    st.info("아직 등록된 환자가 없습니다.")

# 시트 이름 및 교수 리스트 설정
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
    '임플실': '임플란트',
    '원진실': '원내생'
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
    if '예약의사' not in df.columns or '예약시간' not in df.columns:
        raise KeyError("예약의사, 예약시간 열 필요")

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
    return final_df[['진료번호', '예약시간', '환자명', '예약의사', '진료내역']]

def process_excel(file, password):
    decrypted = io.BytesIO()
    mso_file = msoffcrypto.OfficeFile(file)
    mso_file.load_key(password=password)
    mso_file.decrypt(decrypted)
    decrypted.seek(0)

    wb = load_workbook(decrypted, data_only=True)
    result = {}

    for sheet_name in wb.sheetnames:
        values = list(wb[sheet_name].values)
        while values and (values[0] is None or all(v is None for v in values[0])):
            values.pop(0)
        if len(values) < 2:
            continue
        df = pd.DataFrame(values)
        df.columns = df.iloc[0]
        df = df.drop([0]).reset_index(drop=True).fillna("").astype(str)
        df['예약의사'] = df['예약의사'].str.replace(" 교수님", "", regex=False)

        key = sheet_name_mapping.get(sheet_name.strip(), None)
        if not key:
            continue
        professors_list = professors_dict.get(key, [])
        result[sheet_name] = process_sheet(df, professors_list, key)

    return result

# 파일 업로드
st.subheader("📂 OCS 엑셀 파일 업로드")
uploaded_file = st.file_uploader("암호화된 .xlsx 파일을 업로드하세요", type="xlsx")
password = st.text_input("파일 암호 입력", type="password")

if uploaded_file and password:
    try:
        sheets = process_excel(uploaded_file, password)
        for name, df in sheets.items():
            st.markdown(f"### 📄 {name}")
            st.dataframe(df, use_container_width=True)

            matched = df[df.apply(lambda row: any(
                p['name'] in row['환자명'] and p['id'] in row['진료번호'] for p in user_patients
            ), axis=1)]
            if not matched.empty:
                st.warning(f"🚨 등록된 환자 발견: {len(matched)}명")
                st.dataframe(matched)
            else:
                st.success("✅ 등록된 환자가 없습니다.")
    except Exception as e:
        st.error(f"❌ 오류 발생: {str(e)}")
