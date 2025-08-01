# 엑셀 파일 전체 처리 및 스타일 적용
def process_excel_file_and_style(file_bytes_io):
    file_bytes_io.seek(0)

    try:
        # data_only=True로 로드하여 현재 표시된 값을 가져옵니다.
        # 하지만, 이렇게 로드해도 저장 시 문제가 생길 수 있으므로, 아래에서 명시적 텍스트 저장을 적용합니다.
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
    # openpyxl 엔진을 사용하여 ExcelWriter를 생성합니다.
    # 이렇게 하면 openpyxl 워크북 객체를 직접 조작할 수 있습니다.
    with pd.ExcelWriter(output_buffer_for_styling, engine='openpyxl') as writer:
        for sheet_name_raw, df_to_write in processed_sheets_dfs.items():
            df_to_write.to_excel(writer, sheet_name=sheet_name_raw, index=False)

        # pd.ExcelWriter의 workbook 객체를 직접 가져와서 스타일을 적용하고 셀 값을 강제로 텍스트로 만듭니다.
        wb_styled = writer.book

        # 각 시트에 스타일 적용 및 셀 값 앞에 ' 붙이기
        for sheet_name in wb_styled.sheetnames:
            ws = wb_styled[sheet_name]
            header = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}

            for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row), start=1): # 헤더 포함 모든 행 순회
                for cell in row:
                    # 셀 값이 존재하고, 숫자나 None이 아닌 경우에만 처리
                    if cell.value is not None and not isinstance(cell.value, (int, float)):
                        original_value = str(cell.value).strip()
                        # '='으로 시작하는 경우, 앞에 "'"를 붙여 텍스트로 강제 변환
                        # 이미 '로 시작하는 경우 중복해서 붙이지 않도록 확인
                        if original_value.startswith("=") and not original_value.startswith("'="):
                            cell.value = "'" + original_value
                        # 다른 모든 문자열 값도 명시적으로 텍스트로 저장될 수 있도록 처리
                        # 다만, 여기서는 '=' 문제에 집중하므로, 필요한 경우에만 추가합니다.
                        # Excel에서 일반적인 텍스트는 자동으로 텍스트로 인식됩니다.
                        # 이 부분은 현재 문제의 핵심이 아니므로 주석 처리합니다.
                        # elif original_value and not original_value.startswith("'") and not original_value.isdigit():
                        #     cell.value = "'" + original_value


                # 스타일 적용 (기존 로직 유지)
                if row_idx >= 2: # 데이터 행부터 적용
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
