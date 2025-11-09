# professor_reviews_module.py (학생 추가 및 검색 기능 통합)

import streamlit as st
import datetime
import os
import re
import pandas as pd

# 기존 유틸리티 모듈 임포트
from firebase_utils import get_db_refs, sanitize_path

# Firebase 레퍼런스 초기화
users_ref, doctor_users_ref, db_ref_func = get_db_refs()
professor_reviews_ref = db_ref_func("professor_reviews") 
# 💡 [추가] 교수님 목록을 저장할 새로운 레퍼런스
professors_ref = db_ref_func("professors_list")

# 사용자가 선택할 수 있는 과 목록 (config.py 또는 별도 DB에서 가져오는 것이 이상적이나, 여기서는 임시 정의)
DEPARTMENTS = ["외과", "보철", "보존", "치주", "소치", "관악", "영상", "내과", "교정"] 
ALL_DEPARTMENTS_OPTION = "모든 과"

# --- 내부 로직 함수 ---

@st.cache_data(ttl=360)
def load_professor_list():
    """Firebase에서 교수님 목록을 로드합니다."""
    # Firebase에서 전체 교수 목록을 {key: {name: "이름", dept: "과"}} 형태로 가져옴
    data = professors_ref.get()
    if not data:
        # 💡 [초기 목록 설정] 데이터가 없으면 기본 교수 목록을 등록 (최초 1회 실행)
        initial_list = [
            {"name": "김철수", "dept": "외과"}, 
            {"name": "이영희", "dept": "보철"}, 
            {"name": "김철수", "dept": "보존"}, # 동명이인 예시
        ]
        for prof in initial_list:
            key = f"{prof['name']}_{prof['dept']}"
            professors_ref.child(sanitize_path(key)).set(prof)
        
        # 기본 목록 등록 후 다시 로드
        data = professors_ref.get()
    
    # 딕셔너리 데이터를 리스트 형태로 변환하여 반환
    return list(data.values()) if data else []


def _handle_professor_addition(name, dept):
    """새로운 교수님 정보를 Firebase에 추가합니다."""
    if not name or not dept:
        st.error("교수님 이름과 과를 모두 입력해주세요.")
        return

    key = f"{name}_{dept}"
    safe_key = sanitize_path(key)

    # 중복 확인
    existing = professors_ref.child(safe_key).get()
    if existing:
        st.warning(f"'{name}' 교수님 ({dept})은 이미 등록되어 있습니다.")
        return

    # 등록
    professors_ref.child(safe_key).set({"name": name, "dept": dept})
    
    # 캐시 무효화 및 새로고침
    load_professor_list.clear() 
    st.success(f"✅ 교수님 '{name}' ({dept})이(가) 목록에 추가되었습니다.")
    st.rerun()


def _handle_review_submission(professor_name, professor_dept, rating, review_text):
    """익명 평가를 Firebase에 저장합니다."""
    # 고유 키: 이름_과
    unique_key = f"{professor_name}_{professor_dept}"
    if not unique_key or not review_text:
        st.error("평가할 교수님 정보와 내용을 입력해주세요.")
        return

    try:
        new_review = {
            "professor_name": professor_name,
            "professor_dept": professor_dept,
            "rating": rating,
            "review": review_text,
            "timestamp": datetime.datetime.now().isoformat(),
            "user_id": "anonymous_" + os.urandom(8).hex() 
        }
        
        # 고유 키 아래에 자동 생성 키로 평가 저장
        safe_key = sanitize_path(unique_key)
        professor_reviews_ref.child(safe_key).push(new_review)
        st.success(f"🎉 **{professor_name}** 교수님 ({professor_dept})에 대한 익명 평가가 등록되었습니다.")
        
        st.rerun() 
        
    except Exception as e:
        st.error(f"평가 등록 실패: {e}")


def _show_existing_reviews(professor_name, professor_dept):
    """선택된 교수님의 기존 평가를 표시하고 평균 평점을 계산합니다."""
    unique_key = f"{professor_name}_{professor_dept}"
    safe_key = sanitize_path(unique_key)
    all_reviews = professor_reviews_ref.child(safe_key).get()
    
    full_name = f"{professor_name} 교수님 ({professor_dept})"

    if all_reviews and isinstance(all_reviews, dict):
        review_list = list(all_reviews.values())
        
        ratings = [r.get('rating', 0) for r in review_list if isinstance(r, dict)]
        avg_rating = sum(ratings) / len(ratings) if ratings else 0

        st.subheader(f"✅ {full_name} 평가 결과 (총 {len(ratings)}개)")
        st.markdown(f"**평균 평점: {avg_rating:.2f} / 5.0**")
        st.markdown("---")

        for review_data in sorted(review_list, key=lambda x: x.get('timestamp', ''), reverse=True):
            if isinstance(review_data, dict):
                st.markdown(f"**⭐️ 평점: {review_data.get('rating', 'N/A')}**")
                st.text(review_data.get('review', '평가 내용 없음'))
                st.caption(f"등록일: {review_data.get('timestamp', 'N/A')[:10]}")
                st.divider()

    else:
        st.info(f"아직 {full_name}에 대한 등록된 평가가 없습니다.")


# --- 메인 UI 함수 (streamlit_app.py에서 호출) ---

def show_professor_review_system():
    """교수님 평가 시스템의 메인 UI를 표시합니다."""
    st.header("🧑‍🏫 외래 교수님 후기 방명록")
    st.info("학생만 접근 가능하며, 등록된 평가는 익명으로 처리됩니다.")
    st.markdown("---")
    
    # 전체 교수 목록 로드
    all_professors_data = load_professor_list()


    # 2. 검색 UI
    st.subheader("외래교수님 후기검색")
    
    # 💡 [변경] 검색 입력 및 과 필터링
    search_query = st.text_input("이름으로 교수님 검색", key="prof_search_query", placeholder="예: 김철수")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        # 과 필터 (선택적)
        selected_dept_filter = st.selectbox(
            "과 필터 (선택사항)", 
            options=[ALL_DEPARTMENTS_OPTION] + DEPARTMENTS, 
            key="dept_filter"
        )
    
    # 3. 검색 결과 필터링 및 표시
    filtered_professors = []
    
    if search_query:
        search_term = search_query.strip().lower()
        
        for prof in all_professors_data:
            name_match = search_term in prof.get('name', '').lower()
            dept_match = selected_dept_filter == ALL_DEPARTMENTS_OPTION or prof.get('dept') == selected_dept_filter

            if name_match and dept_match:
                filtered_professors.append(prof)
    
    # 4. 검색 결과 또는 전체 목록 표시
    if not search_query and selected_dept_filter != ALL_DEPARTMENTS_OPTION:
        # 검색 없이 과 필터만 사용한 경우
        for prof in all_professors_data:
            if prof.get('dept') == selected_dept_filter:
                 filtered_professors.append(prof)

    if not search_query and not filtered_professors:
         st.info(f"현재 등록된 교수님은 총 **{len(all_professors_data)}명**입니다. 검색하거나 과를 선택해 주세요.")
         # 검색하지 않은 경우 전체 목록을 보여줄 필요는 없습니다 (너무 많을 수 있으므로).

    if filtered_professors:
        st.subheader(f"검색 결과 (총 {len(filtered_professors)}명)")
        
        # 사용자에게 최종 선택할 교수님 목록을 제공
        prof_options_for_select = [
            f"{p['name']} ({p['dept']})" for p in filtered_professors
        ]
        
        # 💡 [변경] 교수님 선택
        selected_prof_str = st.selectbox(
            "평가를 보거나 등록할 교수님을 선택하세요", 
            options=prof_options_for_select, 
            key="final_prof_select"
        )
        
        # 선택된 교수님 정보 추출
        if selected_prof_str:
            # 이름과 과를 분리 (예: 김철수 (외과) -> name='김철수', dept='외과')
            name_match = re.search(r"(.+)\s*\((.+)\)", selected_prof_str)
            if name_match:
                final_name = name_match.group(1).strip()
                final_dept = name_match.group(2).strip()
            else:
                final_name, final_dept = None, None
                st.error("선택된 교수님 정보가 올바르지 않습니다.")
            
            if final_name and final_dept:
                # 5. 평가 등록 폼 및 기존 평가 표시
                st.markdown("---")

                # 5-1. 기존 평가 표시
                _show_existing_reviews(final_name, final_dept)
                
                # 5-2. 평가 등록 폼
                with st.expander(f"📝 {final_name} 교수님 ({final_dept}) 평가 등록", expanded=True):
                    with st.form("new_review_form"):
                        rating = st.slider("평점 (5점 만점)", 1, 5, 3)
                        review_text = st.text_area("익명 평가 내용 (500자 이내)", max_chars=500, height=100)
                        submit_review = st.form_submit_button("평가 등록 (익명)")

                        if submit_review:
                            _handle_review_submission(final_name, final_dept, rating, review_text)
                            
                st.markdown("---")
                


    elif search_query:
        st.warning(f"'{search_query}'(으)로 검색된 교수님이 없습니다.")


    st.markdown("---")
    

    # 2. 교수 추가 폼
    with st.expander("➕ 목록에 새로운 교수님 추가 (학생용)", expanded=False):
        st.subheader("새 교수님 등록")
        with st.form("add_professor_form"):
            new_prof_name = st.text_input("교수님 성함")
            new_prof_dept = st.selectbox("소속 과", DEPARTMENTS)
            add_submitted = st.form_submit_button("교수님 목록에 추가")

            if add_submitted:
                _handle_professor_addition(new_prof_name, new_prof_dept)

