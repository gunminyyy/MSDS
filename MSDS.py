import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from PIL import Image as PILImage
import io
import re
import gc
import numpy as np
import os
import fitz  # PyMuPDF

# 1. 페이지 설정
st.set_page_config(page_title="MSDS 스마트 변환기", layout="wide")
st.title("MSDS 양식 변환기 (PDF 지원 정밀 모드)")
st.markdown("---")

# --------------------------------------------------------------------------
# [함수] 이미지 정규화
# --------------------------------------------------------------------------
def normalize_image(pil_img):
    try:
        if pil_img.mode in ('RGBA', 'LA') or (pil_img.mode == 'P' and 'transparency' in pil_img.info):
            background = PILImage.new('RGB', pil_img.size, (255, 255, 255))
            if pil_img.mode == 'P': pil_img = pil_img.convert('RGBA')
            background.paste(pil_img, mask=pil_img.split()[3])
            pil_img = background
        else:
            pil_img = pil_img.convert('RGB')
        return pil_img.resize((32, 32)).convert('L')
    except:
        return pil_img.resize((32, 32)).convert('L')

# [함수] 리소스 경로 찾기
def get_reference_images():
    img_folder = "reference_imgs"
    ref_images = {}
    if not os.path.exists(img_folder): return {}, False
    try:
        file_list = sorted(os.listdir(img_folder)) 
        for fname in file_list:
            if fname.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.tif', '.tiff')):
                full_path = os.path.join(img_folder, fname)
                try:
                    pil_img = PILImage.open(full_path)
                    ref_images[fname] = pil_img
                except: continue
        return ref_images, True
    except: return {}, False

# [함수] 이미지 비교 매칭
def find_best_match_name(src_img, ref_images):
    best_score = float('inf')
    best_name = None
    try:
        src_norm = normalize_image(src_img)
        src_arr = np.array(src_norm, dtype=np.int16)
        for name, ref_img in ref_images.items():
            ref_norm = normalize_image(ref_img)
            ref_arr = np.array(ref_norm, dtype=np.int16)
            diff = np.mean(np.abs(src_arr - ref_arr))
            if diff < best_score:
                best_score = diff
                best_name = name
        if best_score < 65: return best_name
        else: return None
    except: return None

# [함수] 파일명에서 숫자 추출
def extract_number(filename):
    nums = re.findall(r'\d+', filename)
    return int(nums[0]) if nums else 999

# 2. 파일 업로드
with st.expander("📂 필수 파일 업로드", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        master_data_file = st.file_uploader("1. 중앙 데이터 (master_data.xlsx)", type="xlsx")
        loaded_refs, folder_exists = get_reference_images()
        if folder_exists and loaded_refs:
            st.success(f"✅ 기준 그림 {len(loaded_refs)}개 로드됨 (폴더: reference_imgs)")
        elif not folder_exists:
            st.warning("⚠️ 'reference_imgs' 폴더를 만들고 그림 파일들을 넣어주세요.")

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
    uploaded_files = st.file_uploader("원본 데이터(PDF)", type=["pdf"], accept_multiple_files=True)

with col_center:
    st.write("") ; st.write("") ; st.write("")
    
    if st.button("▶ 변환 시작", use_container_width=True):
        if uploaded_files and master_data_file and template_file:
            with st.spinner("PDF 분석 및 데이터 변환 중..."):
                
                new_files = []
                new_download_data = {}
                
                # 중앙 데이터 읽기
                try: 
                    df_master = pd.read_excel(master_data_file, sheet_name=0)
                    code_map = {}
                    for idx, row in df_master.iterrows():
                        code_val = str(row.iloc[0]).strip()
                        desc_val = str(row.iloc[1]).strip()
                        code_map[code_val] = desc_val
                except: 
                    df_master = pd.DataFrame()
                    code_map = {}

                for uploaded_file in uploaded_files:
                    if option == "CFF(K)":
                        try:
                            # 1. PDF 로드
                            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                            
                            # 2. 양식 파일 준비
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

                            # 제품명 입력
                            dest_ws['B7'] = product_name_input
                            dest_ws['B10'] = product_name_input
                            
                            # ---------------------------------------------------
                            # [PDF 텍스트 분석]
                            # ---------------------------------------------------
                            full_text = ""
                            for page in doc:
                                full_text += page.get_text()

                            # 줄바꿈 제거 (검색 용이성)
                            clean_text = full_text.replace("\n", " ")

                            # [A] 유해성 본문 (B20)
                            start_match = re.search(r"2\.\s*유해성.*?위험성", clean_text)
                            end_match = re.search(r"예방조치문구", clean_text)
                            
                            b20_text = ""
                            if start_match and end_match:
                                start_idx = start_match.end()
                                end_idx = end_match.start()
                                b20_text = clean_text[start_idx:end_idx].strip()[:1000]
                            
                            if b20_text:
                                dest_ws['B20'] = b20_text
                                dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

                            # [B] H코드 추출 및 입력 (B25 ~ B30)
                            extracted_h_codes = []
                            # clean_text 전체에서 H코드 찾기
                            found_h_codes = re.findall(r"H\d{3}", clean_text)
                            for code in found_h_codes:
                                if code not in extracted_h_codes: extracted_h_codes.append(code)
                            
                            # B25 입력
                            current_target_row = 25
                            for code in extracted_h_codes:
                                if current_target_row > 30: break
                                dest_ws.cell(row=current_target_row, column=2).value = code
                                dest_ws.cell(row=current_target_row, column=4).value = code_map.get(code, "")
                                current_target_row += 1
                            
                            # B25~B30 숨김 처리
                            for r in range(25, 31):
                                if not dest_ws.cell(row=r, column=2).value:
                                    dest_ws.row_dimensions[r].hidden = True
                                else:
                                    dest_ws.row_dimensions[r].hidden = False

                            # ---------------------------------------------------
                            # [신규] P코드 추출 (예방, 대응, 저장, 폐기)
                            # ---------------------------------------------------
                            
                            # 섹션별 인덱스 찾기 (검색 범위 한정을 위해)
                            idx_prevention = clean_text.find("예방", end_match.start() if end_match else 0)
                            idx_response = clean_text.find("대응", idx_prevention)
                            idx_storage = clean_text.find("저장", idx_response)
                            idx_disposal = clean_text.find("폐기", idx_storage)
                            
                            # 섹션별 텍스트 자르기
                            txt_prevention = ""
                            txt_response = ""
                            txt_storage = ""
                            txt_disposal = ""
                            
                            if idx_prevention != -1 and idx_response != -1:
                                txt_prevention = clean_text[idx_prevention:idx_response]
                            if idx_response != -1 and idx_storage != -1:
                                txt_response = clean_text[idx_response:idx_storage]
                            if idx_storage != -1 and idx_disposal != -1:
                                txt_storage = clean_text[idx_storage:idx_disposal]
                            if idx_disposal != -1:
                                # 폐기 다음 섹션("3.") 전까지
                                next_section = re.search(r"3\.\s", clean_text[idx_disposal:])
                                end_disposal = idx_disposal + next_section.start() if next_section else len(clean_text)
                                txt_disposal = clean_text[idx_disposal:end_disposal]

                            # 공통 함수: P코드 추출 및 셀 입력
                            def fill_p_codes(target_text, start_row, end_row):
                                # P코드 정규식 (P300+P310 형태 포함)
                                p_codes = re.findall(r"P\d{3}(?:\+P\d{3})*", target_text)
                                unique_p = []
                                for p in p_codes:
                                    if p not in unique_p: unique_p.append(p)
                                
                                # 숨김 취소 (내용 넣기 전 초기화)
                                for r in range(start_row, end_row + 1):
                                    dest_ws.row_dimensions[r].hidden = False
                                
                                curr = start_row
                                for p_code in unique_p:
                                    if curr > end_row: break
                                    dest_ws.cell(row=curr, column=2).value = p_code
                                    dest_ws.cell(row=curr, column=4).value = code_map.get(p_code, "")
                                    curr += 1
                                
                                # 내용 없는 행 숨기기
                                for r in range(start_row, end_row + 1):
                                    if not dest_ws.cell(row=r, column=2).value:
                                        dest_ws.row_dimensions[r].hidden = True

                            # 1. 예방 (B32 ~ B41)
                            fill_p_codes(txt_prevention, 32, 41)
                            
                            # 2. 대응 (B42 ~ B49)
                            fill_p_codes(txt_response, 42, 49)
                            
                            # 3. 저장 (B50 ~ B52) - 기존 숨김 행 포함
                            fill_p_codes(txt_storage, 50, 52)
                            
                            # 4. 폐기 (B53)
                            fill_p_codes(txt_disposal, 53, 53)

                            # ---------------------------------------------------
                            # [기존 기능] PDF 이미지 추출 및 정렬
                            # ---------------------------------------------------
                            target_anchor_row = 22
                            if hasattr(dest_ws, '_images'):
                                preserved_imgs = []
                                for img in dest_ws._images:
                                    try:
                                        if not (target_anchor_row - 2 <= img.anchor._from.row <= target_anchor_row + 2):
                                            preserved_imgs.append(img)
                                    except: preserved_imgs.append(img)
                                dest_ws._images = preserved_imgs
                            
                            collected_pil_images = []
                            for page_index in range(len(doc)):
                                image_list = doc.get_page_images(page_index)
                                for img_info in image_list:
                                    xref = img_info[0]
                                    base_image = doc.extract_image(xref)
                                    image_bytes = base_image["image"]
                                    try:
                                        pil_img = PILImage.open(io.BytesIO(image_bytes))
                                        matched_name = None
                                        if loaded_refs:
                                            matched_name = find_best_match_name(pil_img, loaded_refs)
                                        
                                        if matched_name:
                                            sort_key = extract_number(matched_name)
                                            collected_pil_images.append((sort_key, pil_img))
                                    except: continue
                            
                            # 중복 제거 및 정렬
                            unique_images = {}
                            for key, img in collected_pil_images:
                                if key not in unique_images:
                                    unique_images[key] = img
                            
                            final_images = sorted(unique_images.items(), key=lambda x: x[0])
                            sorted_imgs = [item[1] for item in final_images]
                            
                            if sorted_imgs:
                                unit_size = 67 
                                icon_size = 60 
                                padding_top = 4 
                                padding_left = (unit_size - icon_size) // 2 
                                total_width = unit_size * len(sorted_imgs)
                                total_height = unit_size 
                                merged_img = PILImage.new('RGBA', (total_width, total_height), (255, 255, 255, 0))
                                
                                for idx, p_img in enumerate(sorted_imgs):
                                    p_img_resized = p_img.resize((icon_size, icon_size), PILImage.LANCZOS)
                                    merged_img.paste(p_img_resized, ((idx * unit_size) + padding_left, padding_top))
                                
                                img_byte_arr = io.BytesIO()
                                merged_img.save(img_byte_arr, format='PNG') 
                                img_byte_arr.seek(0)
                                dest_ws.add_image(XLImage(img_byte_arr), 'B23')

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
                
                del df_master
                if 'doc' in locals(): doc.close()
                if 'dest_wb' in locals(): del dest_wb
                if 'output' in locals(): del output
                gc.collect()

                if new_files:
                    st.success("완료! PDF 분석 및 변환이 끝났습니다.")
        else:
            st.error("모든 파일을 업로드해주세요.")

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
