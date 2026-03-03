# excel_utils.py

import streamlit as st
import pandas as pd
import io
import msoffcrypto
import re
from openpyxl import load_workbook
from openpyxl.styles import Font
from config import PROFESSORS_DICT, SHEET_KEYWORD_TO_DEPARTMENT_MAP

# --- 유효성 검사 ---
def is_daily_schedule(file_name):
    """OCS 스케줄 파일 이름 형식(ocs_YYYY.xlsx/xlsm)을 확인합니다."""
    pattern = r'^ocs_\d{4}\.(?:xlsx|xlsm)$'
    return re.match(pattern, file_name, re.IGNORECASE) is not None
    
def is_encrypted_excel(file_path):
    """엑셀 파일이 암호화되었는지 확인합니다."""
    try:
        file_path.seek(0)
        return msoffcrypto.OfficeFile(file_path).is_encrypted()
    except Exception:
        return False

# --- 엑셀 로드 및 복호화 스트림 반환 ---
def load_excel_stream(file, password=None):
    """업로드된 파일을 읽어 복호화된 BytesIO 스트림을 반환합니다."""
    try:
        file.seek(0)
        file_bytes = file.read()
        
        input_stream = io.BytesIO(file_bytes)
        
        # 1. 파일이 암호화되었는지 확인
        is_encrypted = False
        try:
            if msoffcrypto.OfficeFile(input_stream).is_encrypted():
                is_encrypted = True
        except Exception:
            # 파일 구조상 암호화 체크 실패 시 일반 파일로 간주
            pass
        
        # 2. 암호화되어 있다면 복호화 진행
        if is_encrypted:
            if not password:
                raise ValueError("암호화된 엑셀 파일입니다. UI에서 비밀번호를 입력해주세요.")
            
            decrypted_stream = io.BytesIO()
            input_stream.seek(0)
            
            office_file = msoffcrypto.OfficeFile(input_stream)
            try:
                office_file.load_key(password=password)
                office_file.decrypt(decrypted_stream)
            except Exception as e:
                raise ValueError(f"비밀번호가 틀렸거나 복호화에 실패했습니다: {e}")
            
            decrypted_stream.seek(0)
            return decrypted_stream
        else:
            # 3. 일반 파일이면 원본 스트림 반환
            input_stream.seek(0)
            return input_stream
            
    except Exception as e:
        raise ValueError(f"엑셀 스트림 생성 실패: {e}")

# --- 데이터 처리 및 정렬 ---
def process_sheet_v8(df, professors_list, sheet_key): 
    """OCS 시트 데이터를 교수/비교수 기준으로 정렬합니다."""
    
    required_cols = ['진료번호', '예약일시', '예약시간', '환자명', '예약의사', '진료내역']
    if not all(col in df.columns for col in ['예약의사', '예약시간']):
        st.error(f"시트 처리 오류: '예약의사' 또는 '예약시간' 컬럼이 DataFrame에 없습니다.")
        return pd.DataFrame(columns=[col for col in required_cols if col in df.columns])

    df = df.sort_values(by=['예약의사', '예약시간'])
    professors = df[df['예약의사'].isin(professors_list)]
    non_professors = df[~df['예약의사'].isin(professors_list)]

    # 정렬 기준 설정
    if sheet_key != '보철':
        non_professors = non_professors.sort_values(by=['예약시간', '예약의사'])
    else:
        non_professors = non_professors.sort_values(by=['예약의사', '예약시간'])

    final_rows = []
    current_time = None
    current_doctor = None

    # 비교수 환자 처리 (시간 또는 의사별 그룹핑)
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

    # 교수님 섹션 구분자 추가
    if not non_professors.empty:
        final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
    final_rows.append(pd.Series(["<교수님>"] + [" "] * (len(df.columns) - 1), index=df.columns))

    # 교수 환자 처리 (의사별 그룹핑)
    current_professor = None
    for _, row in professors.iterrows():
        if current_professor != row['예약의사']:
            if current_professor is not None:
                final_rows.append(pd.Series([" "] * len(df.columns), index=df.columns))
            current_professor = row['예약의사']
        final_rows.append(row)

    final_df = pd.DataFrame(final_rows, columns=df.columns)
    final_df = final_df[[col for col in required_cols if col in final_df.columns]]
    return final_df

def process_excel_file_and_style(uploaded_file, db_ref_func, excel_password=None):
    """
    엑셀 파일을 읽고, 정렬/스타일링을 적용한 후, 분석용 DataFrame 딕셔너리를 반환합니다.
    (ui_manager.py의 파라미터 규격에 완벽히 맞췄습니다.)
    """
    # 1. 파일 복호화 및 읽기 가능한 스트림 추출
    decrypted_stream = load_excel_stream(uploaded_file, password=excel_password)
    
    output_buffer_for_styling = io.BytesIO()

    # 2. Openpyxl로 워크북 로드
    try:
        wb_raw = load_workbook(filename=decrypted_stream, keep_vba=False, data_only=True)
    except Exception as e:
        raise ValueError(f"엑셀 워크북 로드 실패 (파일 형식 오류): {e}")

    processed_sheets_dfs = {} # 스타일링된 DF (출력용)
    cleaned_raw_dfs = {}       # 정리된 Raw DF (분석
