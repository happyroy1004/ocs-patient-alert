# config.py (수정된 버전)

import streamlit as st
import datetime

# --- 💡 필수 인증 정보 로드 (secrets.toml에서 가져옴) ---
try:
    # Firebase Realtime Database 및 Admin SDK 인증 정보
    # secrets.toml에 [firebase] 섹션으로 저장된 내용을 딕셔너리로 로드합니다.
    FIREBASE_CREDENTIALS = st.secrets["firebase"]
    DB_URL = st.secrets["firebase"]["FIREBASE_DATABASE_URL"] 

    # Google Calendar OAuth 인증 정보
    # secrets.toml에 [google_calendar] 섹션으로 저장된 내용을 딕셔너리로 로드합니다.
    GOOGLE_CALENDAR_CLIENT_SECRET = st.secrets["google_calendar"]
    
except KeyError as e:
    # 인증 정보 로드 실패 시 명시적인 오류 메시지 출력
    st.error(f"🚨 중요: Secrets.toml 설정 오류. '{e.args[0]}' 키를 찾을 수 없습니다. secrets.toml 파일을 확인해 주세요.")
    # 임시로 None을 할당하여 앱이 바로 크래시되는 것을 방지합니다.
    FIREBASE_CREDENTIALS = None
    DB_URL = None
    GOOGLE_CALENDAR_CLIENT_SECRET = None


# --- 전역 상수 정의 ---
# 환자 데이터의 진료과 플래그 키 목록 (DB에 저장되는 T/F 플래그)
PATIENT_DEPT_FLAGS = ["보철", "외과", "내과", "소치", "교정", "원진실", "보존"] 
# 등록 시 선택할 수 있는 모든 진료과
DEPARTMENTS_FOR_REGISTRATION = ["교정", "내과", "보존", "보철", "소치", "외과", "치주", "원진실"]
# Google Calendar Scope
SCOPES = ["https://www.googleapis.com/auth/calendar.events"]
# 초기 비밀번호 (보안상 secrets.toml에 두는 것이 좋으나, 예시를 위해 여기에 명시)
DEFAULT_PASSWORD = "1234" 

# OCS 시트 이름 매핑
SHEET_KEYWORD_TO_DEPARTMENT_MAP = {
    '치과보철과': '보철', '보철과': '보철', '보철': '보철', '치과교정과' : '교정', '교정과': '교정', '교정': '교정',
    '구강 악안면외과' : '외과', '구강악안면외과': '외과', '외과': '외과', '구강 내과' : '내과', '구강내과': '내과', '내과': '내과',
    '치과보존과' : '보존', '보존과': '보존', '보존': '보존', '소아치과': '소치', '소치': '소치', '소아 치과': '소치',
    '원내생진료센터': '원내생', '원내생': '원내생','원내생 진료센터': '원내생','원진실':'원내생',
    '원스톱 협진센터' : '원스톱', '원스톱협진센터': '원스톱', '원스톱': '원스톱',
    '임플란트 진료센터' : '임플란트', '임플란트진료센터': '임플란트', '임플란트': '임플란트',
    '임플' : '임플란트', '치주과': '치주', '치주': '치주', '임플실': '임플란트', '병리': '병리'
}

# 환자 등록과별 OCS 시트 매핑 (매칭 로직에 필요)
PATIENT_DEPT_TO_SHEET_MAP = {
    '보철': ['보철', '임플란트'], '치주': ['치주', '임플란트'], '외과': ['외과', '원스톱', '임플란트'], 
    '교정': ['교정'], '내과': ['내과'], '보존': ['보존'], '소치': ['소치'], '원내생': ['원내생'], '병리': ['병리']
}

# 교수 명단 (분석/정렬에 사용)
PROFESSORS_DICT = {
    '소치': ['김현태', '장기택', '김정욱', '현홍근', '김영재', '신터전', '송지수'],
    '보존': ['이인복', '금기연', '이우철', '유연지', '서덕규', '이창하', '김선영', '손원준'],
    '외과': ['최진영', '서병무', '명훈', '김성민', '박주영', '양훈주', '한정준', '권익재', '서미현'],
    '치주': ['구영', '이용무', '설양조', '구기태', '김성태', '조영단'],
    '보철': ['곽재영', '김성균', '임영준', '김명주', '권호범', '여인성', '윤형인', '박지만', '이재현', '조준호'],
    '교정': [], '내과': [], '원진실': [], '원스톱': [], '임플란트': [], '병리': []
}
