import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment
import io
import copy

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
    # 원본이 엑셀일 경우를 대비해 xlsx 추가
    uploaded_files = st.file_uploader(
        "원본 파일(PDF/Excel)을 드래그해서 넣어주세요", 
        type=["pdf", "xlsx"],
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
                
                # 중앙 데이터 읽기 (참조용, 현재 로직에서는 원본 파일 내용을 주로 사용함)
                df_master = pd.read_excel(master_data_file)
                
                for uploaded_file in uploaded_files:
                    # 결과물 파일명 생성
                    file_name = f"{uploaded_file.name.split('.')[0]}_{option}.xlsx"
                    
                    # -----------------------------------------------------------
                    # [로직 적용] CFF(K) 양식 선택 시
                    # -----------------------------------------------------------
                    if option == "CFF(K)" and uploaded_file.name.endswith('.xlsx'):
                        try:
                            # 1. 원본 엑셀 읽기
                            src_wb = load_workbook(uploaded_file, data_only=True)
                            src_ws = src_wb.active
                            
                            # 2. 양식 엑셀 로드 (여기서는 메모리 상에서 생성한다고 가정하거나 실제 파일 경로 필요)
                            # 테스트를 위해 빈 워크북을 생성하지만, 실제로는 load_workbook("templates/CFF_K.xlsx")여야 함
                            # dest_wb = load_workbook("templates/CFF_K.xlsx")
                            dest_wb = load_workbook(io.BytesIO(master_data_file.getvalue())) # 임시: 양식 파일이 없으므로 마스터파일을 복사해서 씀(교체 필요)
                            dest_ws = dest_wb.active
                            
                            # A. [제품명] 입력 (B7, B10)
                            dest_ws['B7'] = product_name_input
                            dest_ws['B10'] = product_name_input
                            
                            # B. [텍스트 추출] 2. 유해성~ 밑부터 예방조치~ 위까지 D열 복사
                            start_row = 0
                            end_row = 0
                            
                            # 행을 순회하며 키워드 위치 찾기
                            for i, row in enumerate(src_ws.iter_rows(values_only=True), 1):
                                row_str = str(row[0]) if row[0] else "" # A열 기준 검색
                                if "2. 유해성" in row_str and "위험성" in row_str:
                                    start_row = i
                                if "나. 예방조치문구를 포함한 경고표지 항목" in row_str:
                                    end_row = i
                                    break
                            
                            # 내용 추출 및 병합
                            if start_row > 0 and end_row > 0:
                                extracted_texts = []
                                for r in range(start_row + 1, end_row):
                                    cell_val = src_ws.cell(row=r, column=4).value # D열(4번째)
                                    if cell_val:
                                        extracted_texts.append(str(cell_val))
                                
                                # 한 줄씩 줄바꿈하여 B20에 입력
                                final_text = "\n".join(extracted_texts)
                                dest_ws['B20'] = final_text
                                dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center')

                            # C. [그림문자] 이미지 복사 (B23부터 나열)
                            # 그림문자 행 찾기
                            img_header_row = 0
                            for i, row in enumerate(src_ws.iter_rows(values_only=True), 1):
                                if row[0] and "그림문자" in str(row[0]):
                                    img_header_row = i
                                    break
                            
                            if img_header_row > 0:
                                target_img_row_idx = img_header_row # 그림은 헤더 바로 밑 행(index는 0부터 시작하므로 row-1이 아니라 그대로 사용)
                                
                                # 이미지 찾기 (openpyxl의 _images 활용)
                                found_images = []
                                for img in src_ws._images:
                                    # 앵커의 행 위치 확인 (0-based index이므로 -1 보정 필요할 수 있음, 상황에 따라 조정)
                                    if img.anchor._from.row == img_header_row: 
                                        found_images.append(img)
                                
                                # B23 셀 위치부터 붙여넣기
                                start_col_idx = 2 # B열
                                target_row_idx = 23
                                
                                for idx, img in enumerate(found_images):
                                    # 이미지 데이터 복제
                                    img_data = io.BytesIO(img._data())
                                    new_img = XLImage(img_data)
                                    
                                    # 크기 조절 (1.77cm approx 67 pixels)
                                    new_img.width = 67
                                    new_img.height = 67
                                    
                                    # 위치 지정 (B23, C23, D23... 순차적으로)
                                    # 셀 좌표 문자로 변환 (ASCII 코드 활용: B=66)
                                    col_char = chr(66 + idx) 
                                    cell_pos = f"{col_char}{target_row_idx}"
                                    
                                    dest_ws.add_image(new_img, cell_pos)

                            # 저장
                            output = io.BytesIO()
                            dest_wb.save(output)
                            new_download_data[file_name] = output.getvalue()
                            new_files.append(file_name)

                        except Exception as e:
                            st.error(f"변환 중 오류 발생 ({file_name}): {e}")
                    
                    # 다른 양식 또는 PDF인 경우 (기존 로직 유지 또는 패스)
                    else:
                        new_files.append(file_name)
                        new_download_data[file_name] = b"" 
                
                st.session_state['converted_files'] = new_files
                st.session_state['download_data'] = new_download_data
                # ----------------------------------
                
                if new_files:
                    st.success(f"변환 완료!")
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
