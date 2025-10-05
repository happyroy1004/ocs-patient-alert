# excel_utils.py

import streamlit as st
import pandas as pd
import io
import msoffcrypto
import re
from openpyxl import load_workbook
from openpyxl.styles import Font
from .config import PROFESSORS_DICT, SHEET_KEYWORD_TO_DEPARTMENT_MAP

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

# --- 엑셀 로드 및 복호화 ---
def load_excel(file, password=None):
    """업로드된 엑셀 파일을 로드하고 필요시 복호화합니다."""
    try:
        file.seek(0)
        file_bytes = file.read()
        
        input_stream = io.BytesIO(file_bytes)
        decrypted_bytes_io = None
        
        # 파일이 암호화되었는지 확인
        is_encrypted = False
        try:
            if msoffcrypto.OfficeFile(input_stream).is_encrypted():
                is_encrypted = True
        except:
             # 파일 구조상 암호화 체크 실패 시 일반 파일로 간주
            pass
        
        if is_encrypted:
            if not password:
                raise ValueError("암호화된 파일입니다. 비밀번호를 입력해주세요.")
            
            decrypted_bytes_io = io.BytesIO()
            input_stream.seek(0)
            
            office_file = msoffcrypto.OfficeFile(input_stream)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted_bytes_io)
            
            decrypted_bytes_io.seek(0)
            return pd.ExcelFile(decrypted_bytes_io), decrypted_bytes_io

        else:
            input_stream.seek(0)
            return pd.ExcelFile(input_stream), input_stream
            
    except Exception as e:
        raise ValueError(f"엑셀 로드 또는 복호화 실패: {e}")

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

def process_excel_file_and_style(file_bytes_io):
    """엑셀 파일을 읽고, 정렬/스타일링을 적용한 후, 분석용 DataFrame 딕셔너리를 반환합니다."""
    file_bytes_io.seek(0)
    output_buffer_for_styling = io.BytesIO()

    try:
        wb_raw = load_workbook(filename=file_bytes_io, keep_vba=False, data_only=True)
    except Exception as e:
        raise ValueError(f"엑셀 워크북 로드 실패: {e}")

    processed_sheets_dfs = {}
    
    # 1. 시트별 데이터 처리 및 정렬
    for sheet_name_raw in wb_raw.sheetnames:
        sheet_name_lower = sheet_name_raw.strip().lower()

        # 시트 이름 매핑
        sheet_key = None
        for keyword, department_name in sorted(SHEET_KEYWORD_TO_DEPARTMENT_MAP.items(), key=lambda item: len(item[0]), reverse=True):
            if keyword.lower() in sheet_name_lower:
                sheet_key = department_name
                break
        if not sheet_key: continue

        ws = wb_raw[sheet_name_raw]
        values = list(ws.values)
        while values and (values[0] is None or all((v is None or str(v).strip() == "") for v in values[0])):
            values.pop(0)
        if len(values) < 2: continue

        df = pd.DataFrame(values)
        if df.empty or df.iloc[0].isnull().all(): continue

        df.columns = df.iloc[0]
        df = df.drop([0]).reset_index(drop=True)
        df = df.fillna("").astype(str)

        if '예약의사' not in df.columns: continue
        df['예약의사'] = df['예약의사'].str.strip().str.replace(" 교수님", "", regex=False)

        professors_list = PROFESSORS_DICT.get(sheet_key, [])
        
        try:
            # 원본 데이터프레임과 정렬된 데이터프레임 모두 저장
            processed_df = process_sheet_v8(df.copy(), professors_list, sheet_key)
            processed_sheets_dfs[sheet_name_raw] = processed_df
        except Exception as e:
            st.error(f"시트 '{sheet_name_raw}' 처리 중 오류: {e}")
            continue

    if not processed_sheets_dfs:
        # 처리된 데이터가 없는 경우 원본 DF 딕셔너리만 반환
        file_bytes_io.seek(0)
        all_sheet_dfs = pd.read_excel(file_bytes_io, sheet_name=None)
        return all_sheet_dfs, None

    # 2. 정렬된 데이터로 새 엑셀 파일 생성 및 스타일링
    with pd.ExcelWriter(output_buffer_for_styling, engine='openpyxl') as writer:
        for sheet_name_raw, df in processed_sheets_dfs.items():
            df.to_excel(writer, sheet_name=sheet_name_raw, index=False)

    output_buffer_for_styling.seek(0)
    wb_styled = load_workbook(output_buffer_for_styling, keep_vba=False, data_only=True)

    # 스타일링 로직
    for sheet_name in wb_styled.sheetnames:
        ws = wb_styled[sheet_name]
        header = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}

        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
            if row[0].value == "<교수님>":
                for cell in row:
                    if cell.value:
                        cell.font = Font(bold=True)

            # 교정 Bonding 강조 스타일
            if sheet_name.strip() == "교정" and '진료내역' in header:
                idx = header['진료내역'] - 1
                if len(row) > idx:
                    cell = row[idx]
                    text = str(cell.value).strip().lower()
                    
                    if ('bonding' in text or '본딩' in text) and 'debonding' not in text:
                        cell.font = Font(bold=True)

    final_output_bytes = io.BytesIO()
    wb_styled.save(final_output_bytes)
    final_output_bytes.seek(0)
    
    file_bytes_io.seek(0)
    all_sheet_dfs_raw = pd.read_excel(file_bytes_io, sheet_name=None)
    
    return all_sheet_dfs_raw, final_output_bytes

# --- OCS 데이터 분석 ---
def run_analysis(df_dict):
    """OCS 데이터를 기반으로 소치/보존/교정의 통계를 분석합니다."""
    analysis_results = {}
    
    # 분석에 필요한 시트 이름 매핑
    sheet_department_map = {
        '소치': '소치', '소아치과': '소치', '소아 치과': '소치', '보존': '보존', '보존과': '보존', '치과보존과': '보존',
        '교정': '교정', '교정과': '교정', '치과교정과': '교정'
    }
    
    mapped_dfs = {}
    for sheet_name, df in df_dict.items():
        processed_sheet_name = sheet_name.replace(" ", "").lower()
        for key, dept in sheet_department_map.items():
            if processed_sheet_name == key.replace(" ", "").lower():
                # run_analysis에는 정렬되기 전의 원본 DF가 필요합니다.
                if all(col in df.columns for col in ['예약의사', '예약시간', '진료내역']):
                     mapped_dfs[dept] = df.copy()
                break

    # 1. 소치 분석
    if '소치' in mapped_dfs:
        df = mapped_dfs['소치']
        non_professors_df = df[~df['예약의사'].isin(PROFESSORS_DICT.get('소치', []))]
        non_professors_df['예약시간'] = non_professors_df['예약시간'].astype(str).str.strip()
        non_professors_df = non_professors_df[non_professors_df['예약시간'].str.contains(':')] # 유효한 시간만
        
        # 오전: 08:00 ~ 12:50
        morning_patients = non_professors_df[(non_professors_df['예약시간'] >= '08:00') & (non_professors_df['예약시간'] <= '12:50')].shape[0]
        # 오후: 13:00 이후
        afternoon_patients = non_professors_df[non_professors_df['예약시간'] >= '13:00'].shape[0]
        
        # OCS 데이터 특성상 빈 줄이 카운트될 수 있으므로, 최소 1명 이상일 경우에만 조정
        # if afternoon_patients > 0: afternoon_patients -= 1 # 이 조정은 데이터 특성에 따라 달라질 수 있으므로 일단 제거
        analysis_results['소치'] = {'오전': morning_patients, '오후': afternoon_patients}

    # 2. 보존 분석
    if '보존' in mapped_dfs:
        df = mapped_dfs['보존']
        non_professors_df = df[~df['예약의사'].isin(PROFESSORS_DICT.get('보존', []))]
        non_professors_df['예약시간'] = non_professors_df['예약시간'].astype(str).str.strip()
        non_professors_df = non_professors_df[non_professors_df['예약시간'].str.contains(':')]

        # 오전: 08:00 ~ 12:30
        morning_patients = non_professors_df[(non_professors_df['예약시간'] >= '08:00') & (non_professors_df['예약시간'] <= '12:30')].shape[0]
        # 오후: 12:50 이후
        afternoon_patients = non_professors_df[non_professors_df['예약시간'] >= '12:50'].shape[0]
        
        # if afternoon_patients > 0: afternoon_patients -= 1 # 이 조정은 데이터 특성에 따라 달라질 수 있으므로 일단 제거
        analysis_results['보존'] = {'오전': morning_patients, '오후': afternoon_patients}

    # 3. 교정 분석 (Bonding)
    if '교정' in mapped_dfs:
        df = mapped_dfs['교정']
        # Bonding이 포함되고 debonding이 없는 환자 필터링
        bonding_patients_df = df[df['진료내역'].str.contains('bonding|본딩', case=False, na=False) & ~df['진료내역'].str.contains('debonding', case=False, na=False)]
        bonding_patients_df['예약시간'] = bonding_patients_df['예약시간'].astype(str).str.strip()
        bonding_patients_df = bonding_patients_df[bonding_patients_df['예약시간'].str.contains(':')]

        # 오전: 08:00 ~ 12:30
        morning_bonding_patients = bonding_patients_df[(bonding_patients_df['예약시간'] >= '08:00') & (bonding_patients_df['예약시간'] <= '12:30')].shape[0]
        # 오후: 12:50 이후
        afternoon_bonding_patients = bonding_patients_df[bonding_patients_df['예약시간'] >= '12:50'].shape[0]
        
        analysis_results['교정'] = {'오전': morning_bonding_patients, '오후': afternoon_bonding_patients}
        
    return analysis_results
