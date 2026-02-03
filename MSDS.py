import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import re

# 1. 페이지 설정
st.set_page_config(page_title="MSDS 스마트 변환기", layout="wide")

# 2. 제목
st.title("MSDS 양식 변환기 (수식 자동 연결)")
st.markdown("---")

# 3. 파일 설정
with st.expander("📂 파일 설정 (필수)", expanded=True):
    col_master, col_template = st.columns(2)
    with col_master:
        # 이 파일의 첫 번째 시트 내용을 '위험 안전문구' 시트로 만들어 넣습니다.
        master_data_file = st.file_uploader("1. 중앙 데이터 (ingredients...xlsx)", type="xlsx", help="수식에 사용될 데이터 원본입니다.")
    with col_template:
        cff_k_template_file = st.file_uploader("2. CFF(K) 양식 파일", type="xlsx")

product_name_input = st.text_input("제품명을 입력하세요", help="이 값이 B7, B10에 입력됩니다.")
option = st.selectbox("적용할 양식", ("CFF(K)", "CFF(E)", "HP(K)", "HP(E)"))

st.write("") 

# 4. 메인 로직
col_left, col_center, col_right = st.columns([4, 2, 4])

if 'converted_files' not in st.session_state:
    st.session_state['converted_files'] = []
    st.session_state['download_data'] = {}

with col_left:
    st.subheader("3. 원본 파일 업로드")
    uploaded_files = st.file_uploader("원본 데이터(텍스트/그림 포함)", type=["xlsx"], accept_multiple_files=True)

with col_center:
    st.write("") ; st.write("") ; st.write("")
    
    if st.button("▶ 변환 시작", use_container_width=True):
        if uploaded_files and cff_k_template_file and product_name_input and master_data_file:
            with st.spinner("데이터 동기화 및 수식 경로 수정 중..."):
                
                new_files = []
                new_download_data = {}
                
                # 1. 중앙 데이터 읽기 (첫 번째 시트만 사용)
                df_master = pd.read_excel(master_data_file, sheet_name=0)
                
                for uploaded_file in uploaded_files:
                    if option == "CFF(K)":
                        try:
                            # 2. 파일 로드
                            src_wb = load_workbook(uploaded_file, data_only=True)
                            src_ws = src_wb.active
                            
                            # 수식 유지를 위해 data_only=False
                            dest_wb = load_workbook(io.BytesIO(cff_k_template_file.getvalue()))
                            dest_ws = dest_wb.active
                            
                            # ---------------------------------------------------
                            # [1] 중앙 데이터 주입 ('위험 안전문구' 시트 생성)
                            # ---------------------------------------------------
                            target_sheet_name = '위험 안전문구' # 수식이 참조하는 시트명
                            
                            if target_sheet_name in dest_wb.sheetnames:
                                data_ws = dest_wb[target_sheet_name]
                                data_ws.delete_rows(1, data_ws.max_row)
                            else:
                                data_ws = dest_wb.create_sheet(target_sheet_name)
                                
                            for r in dataframe_to_rows(df_master, index=False, header=True):
                                data_ws.append(r)

                            # ---------------------------------------------------
                            # [2] 수식 경로 청소 (외부 링크 -> 내부 링크)
                            # ---------------------------------------------------
                            # 예: 'D:\...\[파일]시트'! -> '시트'! 로 변경
                            # 정규표현식: '문자열[문자열]' 패턴을 찾아서 제거
                            
                            for row in dest_ws.iter_rows():
                                for cell in row:
                                    if cell.data_type == 'f': # 수식인 경우
                                        formula_str = str(cell.value)
                                        # 외부 파일 경로 패턴 찾기 (작은 따옴표 안의 경로 + 대괄호 파일명)
                                        # 간단하게 '['로 시작해서 ']'로 끝나는 파일명 부분을 포함한 경로를 날림
                                        
                                        # 1단계: 알려주신 특정 경로가 있다면 확실하게 제거
                                        if "ingredients CAS and EC 통합.xlsx]" in formula_str:
                                            # 경로가 포함된 형태: 'D:\...\[file]Sheet'
                                            # 이것을 'Sheet'로 바꿔야 함.
                                            # 가장 쉬운 방법: path string을 찾아서 empty로 치환
                                            
                                            # 정규식: '드라이브명: ... [파일명]' 패턴을 찾음
                                            new_formula = re.sub(r"'?[a-zA-Z]:\\[^']*\['?[^']*'?.xlsx\]", "'", formula_str)
                                            # 혹시 경로 없이 [파일명]만 있는 경우도 처리
                                            new_formula = re.sub(r"\[[^\]]*\.xlsx\]", "", new_formula)
                                            
                                            cell.value = new_formula

                            # ---------------------------------------------------
                            # A. 제품명 입력
                            # ---------------------------------------------------
                            dest_ws['B7'] = product_name_input
                            dest_ws['B10'] = product_name_input
                            
                            # ---------------------------------------------------
                            # B. 텍스트 복사 (기존 로직)
                            # ---------------------------------------------------
                            start_row = 0; end_row = 0
                            for i, row in enumerate(src_ws.iter_rows(values_only=True), 1):
                                row_text = " ".join([str(c) for c in row if c])
                                if "2. 유해성" in row_text and "위험성" in row_text: start_row = i
                                if "나. 예방조치문구를 포함한 경고표지 항목" in row_text: end_row = i; break
                            
                            if start_row > 0 and end_row > 0:
                                texts = []
                                for r in range(start_row + 1, end_row):
                                    val = src_ws.cell(row=r, column=4).value
                                    if val: texts.append(str(val).strip())
                                dest_ws['B20'] = "\n".join(texts)
                                dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
                            
                            # ---------------------------------------------------
                            # C. 그림 복사 (기존 로직)
                            # ---------------------------------------------------
                            img_row = 0
                            for i, row in enumerate(src_ws.iter_rows(values_only=True), 1):
                                if "그림문자" in str(row[0]): img_row = i; break
                            
                            if img_row > 0:
                                imgs = [img for img in src_ws._images if img.anchor._from.row >= img_row - 1 and img.anchor._from.row <= img_row + 1]
                                for idx, src_img in enumerate(imgs):
                                    img_bytes = io.BytesIO(src_img._data())
                                    new_img = XLImage(img_bytes)
                                    new_img.width = 67; new_img.height = 67
                                    dest_ws.add_image(new_img, f"{get_column_letter(2 + idx)}23")

                            # 저장
                            output = io.BytesIO()
                            dest_wb.save(output)
                            
                            final_name = f"{product_name_input} GHS MSDS(K).xlsx"
                            if final_name in new_download_data:
                                final_name = f"{product_name_input}_{uploaded_file.name.split('.')[0]} GHS MSDS(K).xlsx"
                            
                            new_download_data[final_name] = output.getvalue()
                            new_files.append(final_name)
                            
                        except Exception as e:
                            st.error(f"오류 ({uploaded_file.name}): {e}")

                st.session_state['converted_files'] = new_files
                st.session_state['download_data'] = new_download_data
                
                if new_files:
                    st.success("완료! (수식 외부 경로가 내부 시트로 자동 연결되었습니다)")
        else:
            st.error("모든 파일과 제품명을 입력해주세요.")

with col_right:
    st.subheader("결과 다운로드")
    if st.session_state['converted_files']:
        for i, fname in enumerate(st.session_state['converted_files']):
            c1, c2 = st.columns([3, 1])
            with c1: st.text(f"📄 {fname}")
            with c2:
                st.download_button("받기", st.session_state['download_data'][fname], file_name=fname, key=i)
