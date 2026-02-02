import streamlit as st

# 1. 페이지 설정
st.set_page_config(page_title="MSDS PDF to Excel Converter", layout="wide")

# 2. 프로그램 제목
st.title("MSDS 양식 변환기")
st.markdown("---")

# 3. 양식 선택 박스
option = st.selectbox(
    "적용할 양식을 선택하세요",
    ("양식 A", "양식 B", "기타 양식")
)

st.write("") 

# 4. 메인 레이아웃 (왼쪽: 여러 파일 업로드 / 가운데: 일괄 변환 버튼 / 오른쪽: 다운로드 목록)
col1, col2, col3 = st.columns([4, 2, 4])

# 변환된 파일들을 저장할 리스트 (나중에 로직에서 채워짐)
if 'converted_files' not in st.session_state:
    st.session_state['converted_files'] = []

with col1:
    st.subheader("원본 파일 업로드")
    # accept_multiple_files=True 옵션 추가
    uploaded_files = st.file_uploader(
        "여러 PDF 파일을 드래그해서 넣어주세요", 
        type="pdf",
        accept_multiple_files=True,
        help="변환하고자 하는 모든 PDF 파일을 선택하세요."
    )

with col2:
    st.write("") 
    st.write("")
    st.write("")
    st.write("")
    # 변환 버튼
    if st.button("▶ 변환 시작", use_container_width=True):
        if uploaded_files:
            with st.spinner(f"{len(uploaded_files)}개의 파일 변환 중..."):
                # --- [로직 추가 구간] ---
                # 임시로 성공 메시지만 표시 (나중에 여기에 for문으로 파일별 처리 로직 삽입)
                st.session_state['converted_files'] = [f"{f.name.split('.')[0]}.xlsx" for f in uploaded_files]
                # -----------------------
                st.success(f"{len(uploaded_files)}개 파일 변환 완료!")
        else:
            st.error("파일을 하나 이상 업로드해주세요.")

with col3:
    st.subheader("변환된 파일 목록")
    
    if uploaded_files and st.session_state['converted_files']:
        st.write(f"총 {len(st.session_state['converted_files'])}개의 결과물:")
        
        # 파일별로 다운로드 버튼 생성 (UI 예시)
        for i, file_name in enumerate(st.session_state['converted_files']):
            c_left, c_right = st.columns([3, 1])
            with c_left:
                st.text(f"📄 {file_name}")
            with c_right:
                # 실제 로직 구현 시 data에 변환된 엑셀 바이너리를 넣어야 함
                st.download_button(
                    label="받기",
                    data=b"", # 실제 엑셀 데이터가 들어갈 자리
                    file_name=file_name,
                    key=f"dl_btn_{i}" # 고유 키 필요
                )
    else:
        st.info("파일을 업로드하고 변환 시작을 눌러주세요.")

