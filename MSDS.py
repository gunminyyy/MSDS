import streamlit as st

# 1. 페이지 설정 (웹 브라우저 탭에 표시될 내용)
st.set_page_config(page_title="PDF to Excel Converter", layout="wide")

# 2. 프로그램 제목
st.title("MSDS 양식 변환기")
st.markdown("---")

# 3. 선택 박스 (양식 선택 등)
option = st.selectbox(
    "적용할 양식을 선택하세요",
    ("양식 A (기본)", "양식 B (정밀 분석)", "기타 양식")
)

st.write("") # 간격 조절

# 4. 메인 레이아웃 (왼쪽: 업로드 / 가운데: 버튼 / 오른쪽: 다운로드)
col1, col2, col3 = st.columns([4, 2, 4])

with col1:
    st.subheader("원본 파일 업로드")
    uploaded_file = st.file_uploader(
        "PDF 파일을 드래그해서 넣어주세요", 
        type="pdf",
        help="변환하고자 하는 원본 PDF 파일을 선택하세요."
    )

with col2:
    st.write("") # 버튼 위치를 내리기 위한 공백
    st.write("")
    st.write("")
    st.write("")
    # 변환 버튼
    if st.button("▶ 변환 시작", use_container_width=True):
        if uploaded_file is not None:
            with st.spinner("변환 중..."):
                # --- 여기에 나중에 로직을 추가할 예정입니다 ---
                # 1. PDF 읽기
                # 2. 데이터 추출
                # 3. 엑셀 양식에 쓰기
                # ----------------------------------------
                st.success("변환 완료!")
        else:
            st.error("파일을 먼저 업로드해주세요.")

with col3:
    st.subheader("변환된 파일 다운로드")
    # 변환이 완료된 후 파일이 나타나는 목록 (예시용 데이터)
    if uploaded_file is not None:
        st.info("변환된 파일이 여기에 표시됩니다.")
        
        # 실제 배포 시에는 변환된 파일 경로를 연결합니다.
        # st.download_button(
        #     label="엑셀 파일 다운로드",
        #     data=None, # 여기에 실제 데이터가 들어갑니다.
        #     file_name="result.xlsx",
        #     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        # )
    else:
        st.write("파일을 업로드하면 다운로드 목록이 활성화됩니다.")

# 5. 하단 안내문 (선택 사항)
st.markdown("---")
st.caption("© 2024 PDF to Excel Auto System - 깃허브 및 스트림릿 배포용")

