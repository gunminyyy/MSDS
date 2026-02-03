import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io

# 1. 페이지 설정
st.set_page_config(page_title="MSDS PDF to Excel Converter", layout="wide")

# 2. 프로그램 제목
st.title("MSDS 양식 변환기")
st.markdown("---")

# 3. 데이터 관리 섹션 (임시 추가: 웹 테스트용 중앙 데이터 업로드)
with st.expander("📂 중앙 데이터베이스 설정", expanded=True):
    master_data_file = st.file_uploader("master_data.xlsx 파일을 업로드하세요", type="xlsx")
    if master_data_file:
        st.success("데이터베이스 로드 완료")

# 제품명 입력 칸 추가
product_name_input = st.text_input("제품명을 입력하세요", help="엑셀 양식에 기입될 제품명입니다.")

# 4. 양식 선택 박스 (4가지 양식으로 수정)
option = st.selectbox(
    "적용할 양식을 선택하세요",
    ("CFF(K)", "CFF(E)", "HP(K)", "HP(E)")
)

st.write("") 

# 5. 메인 레이아웃
col1, col2, col3 = st.columns([4, 2, 4])

if 'converted_files' not in st.session_state:
    st.session_state['converted_files'] = []
    st.session_state['download_data'] = {} # 실제 파일 데이터를 저장할 딕셔너리

with col1:
    st.subheader("원본 파일 업로드")
    uploaded_files = st.file_uploader(
        "여러 PDF 파일을 드래그해서 넣어주세요", 
        type="pdf",
        accept_multiple_files=True
    )

with col2:
    st.write("") ; st.write("") ; st.write("") ; st.write("")
    
    if st.button("▶ 변환 시작", use_container_width=True):
        if uploaded_files and master_data_file:
            with st.spinner(f"{len(uploaded_files)}개의 파일 변환 중..."):
                
                # --- [변환 핵심 로직 시작 구역] ---
                new_files = []
                new_download_data = {}
                
                # 중앙 데이터 읽기
                df_master = pd.read_excel(master_data_file)
                
                for pdf_file in uploaded_files:
                    # 1. PDF에서 키워드 추출 (나중에 구체화)
                    # keyword = extract_keyword_from_pdf(pdf_file)
                    
                    # 2. 중앙 데이터에서 매칭 정보 찾기
                    # matched_info = df_master[df_master['키워드'] == keyword]
                    
                    # 3. 양식 로드 및 데이터 쓰기 (나중에 파일 보내주시면 구현)
                    # output_excel = write_to_template(option, matched_info, product_name_input)
                    
                    file_name = f"{pdf_file.name.split('.')[0]}_{option}.xlsx"
                    new_files.append(file_name)
                    new_download_data[file_name] = b"" # 실제 결과 바이너리 들어갈 곳
                
                st.session_state['converted_files'] = new_files
                st.session_state['download_data'] = new_download_data
                # ----------------------------------
                
                st.success(f"{len(uploaded_files)}개 파일 변환 완료!")
        elif not master_data_file:
            st.error("중앙 데이터베이스 파일을 먼저 업로드해주세요.")
        else:
            st.error("파일을 하나 이상 업로드해주세요.")

with col3:
    st.subheader("변환된 파일 목록")
    if uploaded_files and st.session_state['converted_files']:
        for i, file_name in enumerate(st.session_state['converted_files']):
            c_left, c_right = st.columns([3, 1])
            with c_left:
                st.text(f"📄 {file_name}")
            with c_right:
                st.download_button(
                    label="받기",
                    data=st.session_state['download_data'].get(file_name, b""),
                    file_name=file_name,
                    key=f"dl_btn_{i}"
                )
    else:
        st.info("파일을 업로드하고 변환 시작을 눌러주세요.")

st.markdown("---")
st.caption("© 2024 PDF to Excel Auto System - 깃허브 및 스트림릿 배포용")
