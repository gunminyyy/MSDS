import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from PIL import Image as PILImage # 이미지 처리 라이브러리
import io
import re

# 1. 페이지 설정
st.set_page_config(page_title="MSDS 스마트 변환기", layout="wide")
st.title("MSDS 양식 변환기 (그림 병합 배치)")
st.markdown("---")

# 2. 파일 업로드
with st.expander("📂 필수 파일 업로드", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        master_data_file = st.file_uploader("1. 중앙 데이터 (master_data.xlsx)", type="xlsx")
    with col2:
        template_file = st.file_uploader("2. 양식 파일 (통합 양식 GHS MSDS(K).xlsx)", type="xlsx")

product_name_input = st.text_input("제품명 입력 (B7, B10)")
option = st.selectbox("적용할 양식", ("CFF(K)", "CFF(E)", "HP(K)", "HP(E)"))
st.write("") 

# 3. 메인 로직
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
            with st.spinner("그림을 하나로 합쳐서 배치하는 중..."):
                
                new_files = []
                new_download_data = {}
                
                # 중앙 데이터 읽기
                try: df_master = pd.read_excel(master_data_file, sheet_name=0)
                except: df_master = pd.DataFrame()

                for uploaded_file in uploaded_files:
                    if option == "CFF(K)":
                        try:
                            # 1. 파일 로드
                            src_wb = load_workbook(uploaded_file, data_only=True)
                            src_ws = src_wb.active
                            
                            template_file.seek(0)
                            dest_wb = load_workbook(io.BytesIO(template_file.read()))
                            dest_ws = dest_wb.active

                            # ---------------------------------------------------
                            # [데이터 동기화 & 수식 수정]
                            # ---------------------------------------------------
                            target_sheet = '위험 안전문구'
                            if target_sheet in dest_wb.sheetnames: del dest_wb[target_sheet]
                            data_ws = dest_wb.create_sheet(target_sheet)
                            for r in dataframe_to_rows(df_master, index=False, header=True): data_ws.append(r)

                            for row in dest_ws.iter_rows():
                                for cell in row:
                                    if cell.data_type == 'f':
                                        f_str = str(cell.value)
                                        if "ingredients CAS and EC 통합.xlsx]" in f_str:
                                            new_f = re.sub(r"'?[a-zA-Z]:\\[^']*\['?[^']*'?.xlsx\]", "'", f_str)
                                            new_f = re.sub(r"\[[^\]]*\.xlsx\]", "", new_f)
                                            cell.value = new_f

                            # 제품명 및 텍스트 복사
                            dest_ws['B7'] = product_name_input
                            dest_ws['B10'] = product_name_input
                            
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
                            # [핵심 수정] 그림 병합 배치 (Image Merging)
                            # ---------------------------------------------------
                            
                            # 1. 기존 그림 삭제 (B23 근처)
                            target_anchor_row = 22
                            if hasattr(dest_ws, '_images'):
                                preserved_imgs = []
                                for img in dest_ws._images:
                                    try:
                                        if not (target_anchor_row - 2 <= img.anchor._from.row <= target_anchor_row + 2):
                                            preserved_imgs.append(img)
                                    except: preserved_imgs.append(img)
                                dest_ws._images = preserved_imgs

                            # 2. 원본 그림 수집
                            img_row = 0
                            for i, row in enumerate(src_ws.iter_rows(values_only=True), 1):
                                if "그림문자" in str(row[0]): img_row = i; break
                            
                            collected_pil_images = []
                            if img_row > 0 and hasattr(src_ws, '_images'):
                                for img in src_ws._images:
                                    if hasattr(img, 'anchor'):
                                        r = img.anchor._from.row
                                        if img_row - 2 <= r <= img_row + 1:
                                            # PIL 이미지로 변환하여 리스트에 저장
                                            if hasattr(img, '_data'):
                                                pil_img = PILImage.open(io.BytesIO(img._data()))
                                                collected_pil_images.append(pil_img)
                            
                            # 3. 그림 합치기 (Stitching)
                            if collected_pil_images:
                                # 개별 그림 크기 설정 (1.77cm ≈ 67px)
                                unit_size = 67 
                                total_width = unit_size * len(collected_pil_images)
                                total_height = unit_size
                                
                                # 투명 배경의 빈 캔버스 생성
                                merged_img = PILImage.new('RGBA', (total_width, total_height), (255, 255, 255, 0))
                                
                                for idx, p_img in enumerate(collected_pil_images):
                                    # 크기 리사이징 (깨짐 방지 위해 고품질 리샘플링 사용)
                                    p_img_resized = p_img.resize((unit_size, unit_size), PILImage.LANCZOS)
                                    # 캔버스에 붙여넣기 (x 좌표를 이동시켜가며)
                                    merged_img.paste(p_img_resized, (idx * unit_size, 0))
                                
                                # 4. 합친 이미지를 엑셀에 삽입
                                img_byte_arr = io.BytesIO()
                                merged_img.save(img_byte_arr, format='PNG') # PNG로 저장해야 투명도 유지됨
                                img_byte_arr.seek(0)
                                
                                final_xl_img = XLImage(img_byte_arr)
                                dest_ws.add_image(final_xl_img, 'B23') # B23 셀 하나에만 넣음

                            # 저장
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
                    st.success("완료! 그림들이 깔끔하게 이어졌습니다.")
        else:
            st.error("파일을 모두 업로드해주세요.")

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
