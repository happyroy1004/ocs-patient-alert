# professor_reviews_module.py

import streamlit as st
import datetime
import os
import re

# 기존 유틸리티 모듈 임포트
# Note: 이 파일들이 'ui_manager.py'와 동일 레벨에 있어야 합니다.
from firebase_utils import get_db_refs, sanitize_path

# Firebase 레퍼런스 초기화 (이 모듈 내에서 독립적으로 처리)
# db_ref_func는 데이터베이스 경로를 받아서 레퍼런스를 반환하는 함수입니다.
users_ref, doctor_users_ref, db_ref_func = get_db_refs()
professor_reviews_ref = db_ref_func("professor_reviews") 

# --- 내부 로직 함수 ---

def _handle_review_submission(professor_name, rating, review_text):
    """익명 평가를 Firebase에 저장합니다."""
    # professor_reviews_ref는 이 모듈 상단에서 정의되었습니다.
    if not professor_name or not review_text:
        st.error("교수님 이름과 평가 내용을 입력해주세요.")
        return

    try:
        # 💡 익명성 보장: 사용자 ID 대신 랜덤 키를 사용하거나 아예 저장하지 않습니다.
        new_review = {
            "professor": professor_name,
            "rating": rating,
            "review": review_text,
            "timestamp": datetime.datetime.now().isoformat(),
            "user_id": "anonymous_" + os.urandom(8).hex() 
        }
        
        # 교수님 이름 아래에 고유한 자동 생성 키로 저장
        # sanitize_path를 사용하여 교수님 이름의 특수문자를 처리합니다.
        safe_prof_key = sanitize_path(professor_name)
        professor_reviews_ref.child(safe_prof_key).push(new_review)
        st.success(f"🎉 **{professor_name}**에 대한 익명 평가가 등록되었습니다.")
        
        # 성공 시, 폼 데이터를 클리어하기 위해 st.rerun()을 호출하는 것이 일반적입니다.
        st.rerun() 
        
    except Exception as e:
        st.error(f"평가 등록 실패: {e}")

def _show_existing_reviews(professor_name):
    """선택된 교수님의 기존 평가를 표시하고 평균 평점을 계산합니다."""
    safe_prof_key = sanitize_path(professor_name)
    all_reviews = professor_reviews_ref.child(safe_prof_key).get()
    
    if all_reviews and isinstance(all_reviews, dict):
        review_list = list(all_reviews.values())
        
        # 평균 평점 계산
        ratings = [r.get('rating', 0) for r in review_list if isinstance(r, dict)]
        avg_rating = sum(ratings) / len(ratings) if ratings else 0

        st.subheader(f"✅ {professor_name} 평가 결과 (총 {len(ratings)}개)")
        st.markdown(f"**평균 평점: {avg_rating:.2f} / 5.0**")
        st.markdown("---")

        # 평가 목록 표시 (최신순)
        for review_data in sorted(review_list, key=lambda x: x.get('timestamp', ''), reverse=True):
            if isinstance(review_data, dict):
                st.markdown(f"**⭐️ 평점: {review_data.get('rating', 'N/A')}**")
                st.text(review_data.get('review', '평가 내용 없음'))
                st.caption(f"등록일: {review_data.get('timestamp', 'N/A')[:10]}")
                st.divider()
    else:
        st.info(f"아직 {professor_name}에 대한 등록된 평가가 없습니다.")

# --- 메인 UI 함수 (streamlit_app.py에서 호출) ---

def show_professor_review_system():
    """교수님 평가 시스템의 메인 UI를 표시합니다."""
    st.header("🧑‍🏫 교수님 익명 평가 시스템")
    st.info("로그인 여부와 관계없이 평가를 확인하고 등록할 수 있습니다. 등록된 평가는 익명으로 처리됩니다.")
    st.markdown("---")
    
    # 1. 교수님 목록 (실제로는 DB에서 가져오는 것이 이상적입니다.)
    # 여기서는 예시로 하드코딩
    professor_list = ["김철수 교수님", "이영희 교수님", "박민준 교수님", "최지원 교수님"] 
    selected_professor = st.selectbox("평가를 보거나 등록할 교수님을 선택하세요", professor_list, key="prof_select")

    # 2. 평가 등록 폼
    with st.expander(f"📝 {selected_professor} 평가 등록", expanded=False):
        with st.form("new_review_form"):
            rating = st.slider("평점 (5점 만점)", 1, 5, 3)
            review_text = st.text_area("익명 평가 내용 (500자 이내)", max_chars=500, height=100)
            submit_review = st.form_submit_button("평가 등록 (익명)")

            if submit_review:
                _handle_review_submission(selected_professor, rating, review_text)
                
    st.markdown("---")
    
    # 3. 기존 평가 표시
    _show_existing_reviews(selected_professor)
