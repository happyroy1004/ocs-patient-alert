uploaded_file = st.file_uploader("📂 Excel 파일 업로드", type=["xlsx"])
if uploaded_file and firebase_key:
    # 엑셀 파일 읽기
    df = pd.read_excel(uploaded_file, sheet_name=None)
    
    # Firebase에서 기존 등록된 환자 정보 가져오기
    ref = db.reference(f"patients/{firebase_key}")
    existing_data = ref.get()
    existing_set = set()
    if existing_data:
        for item in existing_data.values():
            existing_set.add((str(item.get("name")).strip(), str(item.get("number")).strip()))

    for sheet_name, sheet_df in df.items():
        st.subheader(f"📄 시트: {sheet_name}")

        if "환자명" not in sheet_df.columns or "진료번호" not in sheet_df.columns:
            st.warning("❌ '환자명' 또는 '진료번호' 열이 없습니다.")
            continue

        results = []
        for _, row in sheet_df.iterrows():
            name = str(row["환자명"]).strip()
            number = str(row["진료번호"]).strip()
            exists = (name, number) in existing_set
            results.append({
                "환자명": name,
                "진료번호": number,
                "등록 여부": "✅ 등록됨" if exists else "➕ 미등록"
            })

        result_df = pd.DataFrame(results)
        st.dataframe(result_df)
