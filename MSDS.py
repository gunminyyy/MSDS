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
st.title("MSDS 양식 변환기 (진단 모드)")
st.info("파일이 손상되는 원인을 찾기 위해, 왼쪽 사이드바의 기능을 하나씩 켜면서 테스트해보세요.")
st.markdown("---")

# 3. [진단 옵션] 사이드바
with st.sidebar:
    st.header("🔧 기능 선택 (하나씩 켜보세요)")
    st.write("아래 순서대로 하나씩 체크하며 변환해보세요. 언제 파일이 안 열리는지 확인해야 합니다.")
    
    opt_basic_save = st.checkbox("0. 아무것도 안 하고 저장만 하기", value=True, disabled=True, help="기본 파일 입출력 테스트입니다.")
    opt_prod_name = st.checkbox("1. 제품명 입력 (B7, B10)", value=True)
    opt_text_copy = st.checkbox("2. 본문 텍스트 복사 (B20)", value=False)
    opt_data_sync = st.checkbox("3. 중앙 데이터 시트 생성", value=False, help="이걸 켰을 때 안 열리면 데이터 시트 생성 문제입니다.")
    opt_formula_fix = st.checkbox("4. 수식 경로 자동 수정 (가장 의심됨)", value=False, help="이걸 켰을 때 안 열리면 수식 수정 로직 문제입니다.")
    opt_img_copy = st.checkbox("5. 그림 복사", value=False)

# 4. 파일 업로드
with st.expander("📂 필수 파일 업로드", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        master_data_file = st.file_uploader("1. 중앙 데이터 (master_data.xlsx)", type="xlsx")
    with col2:
        template_file = st.file_uploader("2. 양식 파일 (통합 양식 GHS MSDS(K).xlsx)", type="xlsx")

product_name_input = st.text_input("제품명을 입력하세요")
option = st.selectbox("적용할 양식", ("CFF(K)", "CFF(E)", "HP(K)", "HP(E)"))

st.write("") 

# 5. 메인 로직
col_left, col_center, col_right = st.columns([4, 2, 4])

if 'converted_files' not in st.session_state:
    st.session_state['converted_files'] = []
    st.session_state['download_data'] = {}

with col_left:
    st.subheader("3. 원본 파일 업로드")
    uploaded_files = st.file_uploader("원본 데이터(엑셀)", type=["xlsx"], accept_multiple_files=True)

with col_center:
    st.write("") ; st.write("") ; st.write("")
    
    if st.button("▶ 변환 시작", use_container_width=True):
        if uploaded_files and master_data_file and template_file:
            with st.spinner("진단 모드로 변환 중..."):
                
                new_files = []
                new_download_data = {}
                
                # 1. 중앙 데이터 읽기
                try:
                    df_master = pd.read_excel(master_data_file, sheet_name=0)
                except:
                    df_master = pd.DataFrame() # 에러 방지용 빈 데이터프레임

                for uploaded_file in uploaded_files:
                    try:
                        # ---------------------------------------------------
                        # [Step 0] 파일 열기 (가장 기본)
                        # ---------------------------------------------------
                        template_file.seek(0)
                        dest_wb = load_workbook(io.BytesIO(template_file.read()))
                        dest_ws = dest_wb.active
                        
                        # 원본 파일 로드 (텍스트/그림 복사용)
                        src_wb = load_workbook(uploaded_file, data_only=True)
                        src_ws = src_wb.active

                        # ---------------------------------------------------
                        # [Step 3] 중앙 데이터 동기화 (옵션)
                        # ---------------------------------------------------
                        if opt_data_sync:
                            target_sheet_name = '위험 안전문구'
                            # 시트 삭제 대신 clear 방식으로 시도 (안전성 향상)
                            if target_sheet_name in dest_wb.sheetnames:
                                # 기존 시트 제거
                                del dest_wb[target_sheet_name]
                            
                            # 새 시트 생성
                            data_ws = dest_wb.create_sheet(target_sheet_name)
                            for r in dataframe_to_rows(df_master, index=False, header=True):
                                data_ws.append(r)

                        # ---------------------------------------------------
                        # [Step 4] 수식 경로 청소 (옵션)
                        # ---------------------------------------------------
                        if opt_formula_fix:
                            for row in dest_ws.iter_rows():
                                for cell in row:
                                    if cell.data_type == 'f':
                                        formula_str = str(cell.value)
                                        if "ingredients CAS and EC 통합.xlsx]" in formula_str:
                                            # 가장 보수적인 치환 (단순화)
                                            new_formula = formula_str.replace("'D:\\Naver MYBOX\\★공유\\업체제출자료양식\\MSDS\\업체별\\[ingredients CAS and EC 통합.xlsx]위험 안전문구'", "'위험 안전문구'")
                                            new_formula = new_formula.replace("[ingredients CAS and EC 통합.xlsx]", "")
                                            cell.value = new_formula

                        # ---------------------------------------------------
                        # [Step 1] 제품명 입력 (옵션)
                        # ---------------------------------------------------
                        if opt_prod_name:
                            dest_ws['B7'] = product_name_input
                            dest_ws['B10'] = product_name_input
                        
                        # ---------------------------------------------------
                        # [Step 2] 텍스트 복사 (옵션)
                        # ---------------------------------------------------
                        if opt_text_copy:
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
                        # [Step 5] 그림 복사 (옵션)
                        # ---------------------------------------------------
                        if opt_img_copy:
                            img_row = 0
                            for i, row in enumerate(src_ws.iter_rows(values_only=True), 1):
                                if "그림문자" in str(row[0]): img_row = i; break
                            
                            if img_row > 0:
                                imgs = []
                                if hasattr(src_ws, '_images'):
                                    for img in src_ws._images:
                                        if hasattr(img, 'anchor'):
                                            row_idx = img.anchor._from.row
                                            if row_idx >= img_row - 2 and row_idx <= img_row:
                                                imgs.append(img)
                                
                                for idx, src_img in enumerate(imgs):
                                    if hasattr(src_img, '_data'):
                                        img_bytes = io.BytesIO(src_img._data())
                                        new_img = XLImage(img_bytes)
                                        new_img.width = 67; new_img.height = 67
                                        dest_ws.add_image(new_img, f"{get_column_letter(2 + idx)}23")

                        # ---------------------------------------------------
                        # 저장
                        # ---------------------------------------------------
                        output = io.BytesIO()
                        dest_wb.save(output)
                        output.seek(0)
                        
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
                    st.success("변환 완료! 다운로드 파일을 확인하세요.")
        else:
            st.error("필수 파일을 모두 업로드해주세요.")

with col_right:
    st.subheader("결과 다운로드")
    if st.session_state['converted_files']:
        for i, fname in enumerate(st.session_state['converted_files']):
            c1, c2 = st.columns([3, 1])
            with c1: st.text(f"📄 {fname}")
            with c2:
                st.download_button(
                    label="받기", 
                    data=st.session_state['download_data'][fname], 
                    file_name=fname, 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=i
                )
