import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.cell.cell import MergedCell # [필수] 병합 셀 감지용
from PIL import Image as PILImage
import io
import re
import gc
import numpy as np
import os
import fitz  # PyMuPDF

# 1. 페이지 설정
st.set_page_config(page_title="MSDS 스마트 변환기", layout="wide")
st.title("MSDS 양식 변환기 (최종 - 병합 충돌 자동 해결)")
st.markdown("---")

# --------------------------------------------------------------------------
# [함수] 이미지 처리
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

def extract_number(filename):
    nums = re.findall(r'\d+', filename)
    return int(nums[0]) if nums else 999

# --------------------------------------------------------------------------
# [함수] PDF 텍스트 정밀 파싱
# --------------------------------------------------------------------------
def parse_pdf_ghs_logic(doc):
    clean_lines = []
    NOISE_KEYWORDS = [
        "물질안전보건자료", "MSDS", "Material Safety Data Sheet",
        "Corea flavors", "주식회사 고려", "HAIR CARE", "Ver.", "발행일", "개정일",
        "제 품 명", "GHS", "페이지", "PAGE", "---"
    ]

    for page in doc:
        blocks = page.get_text("blocks", sort=True)
        for b in blocks:
            text = b[4]
            lines = text.split('\n')
            for line in lines:
                line_str = line.strip()
                if not line_str: continue
                is_noise = False
                for kw in NOISE_KEYWORDS:
                    if kw.replace(" ", "") in line_str.replace(" ", ""):
                        is_noise = True; break
                if not is_noise: clean_lines.append(line_str)

    result = {
        "hazard_cls": [], "signal_word": "", "h_codes": [],
        "p_prev": [], "p_resp": [], "p_stor": [], "p_disp": []
    }

    ZONE_NONE = 0
    ZONE_HAZARD_CLS = 1
    ZONE_LABEL_INFO = 2
    
    SUBZONE_PREV = 11; SUBZONE_RESP = 12; SUBZONE_STOR = 13; SUBZONE_DISP = 14

    current_zone = ZONE_NONE
    current_subzone = None
    
    regex_code = re.compile(r"([HP]\d{3}(?:\s*\+\s*[HP]\d{3})*)")
    BLACKLIST_IN_HAZARD = ["공급자정보", "회사명", "주소", "긴급전화번호", "권고용도", "사용상의제한"]

    for line in clean_lines:
        line_ns = line.replace(" ", "")
        
        if "가.유해성" in line_ns and "분류" in line_ns:
            current_zone = ZONE_HAZARD_CLS; continue
        if "나.예방조치" in line_ns:
            current_zone = ZONE_LABEL_INFO; current_subzone = None; continue
        if "3.구성성분" in line_ns or "다.기타" in line_ns:
            current_zone = ZONE_NONE; break

        if current_zone == ZONE_HAZARD_CLS:
            is_blacklisted = False
            for bl in BLACKLIST_IN_HAZARD:
                if bl in line_ns: is_blacklisted = True; break
            if not is_blacklisted:
                result["hazard_cls"].append(line)
                codes = regex_code.findall(line)
                for c in codes:
                    if c.startswith("H"): result["h_codes"].append(c)

        elif current_zone == ZONE_LABEL_INFO:
            if "신호어" in line_ns:
                val = line.replace("신호어", "").replace(":", "").strip()
                if val: result["signal_word"] = val
            
            if line_ns.startswith("예방") and len(line_ns) < 10: current_subzone = SUBZONE_PREV
            elif line_ns.startswith("대응") and len(line_ns) < 10: current_subzone = SUBZONE_RESP
            elif line_ns.startswith("저장") and len(line_ns) < 10: current_subzone = SUBZONE_STOR
            elif line_ns.startswith("폐기") and len(line_ns) < 10: current_subzone = SUBZONE_DISP

            codes = regex_code.findall(line)
            for c in codes:
                if c.startswith("H"): result["h_codes"].append(c)
                elif c.startswith("P"):
                    if current_subzone == SUBZONE_PREV: result["p_prev"].append(c)
                    elif current_subzone == SUBZONE_RESP: result["p_resp"].append(c)
                    elif current_subzone == SUBZONE_STOR: result["p_stor"].append(c)
                    elif current_subzone == SUBZONE_DISP: result["p_disp"].append(c)

    return result

# --------------------------------------------------------------------------
# [핵심] 안전 쓰기 함수 (병합된 셀이면 해제 후 입력)
# --------------------------------------------------------------------------
def safe_write(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    
    # 해당 셀이 병합된 셀의 일부인지 확인 (MergedCell 객체인지)
    if isinstance(cell, MergedCell):
        # 병합된 범위 찾아서 해제
        for merged_range in list(ws.merged_cells.ranges):
            if cell.coordinate in merged_range:
                ws.unmerge_cells(str(merged_range))
                break
        # 해제 후 다시 셀 객체 가져오기 (이제 일반 Cell임)
        cell = ws.cell(row=row, column=col)
    
    cell.value = value

# --------------------------------------------------------------------------
# [함수] 동적 쓰기 (safe_write 적용)
# --------------------------------------------------------------------------
def write_section_dynamic(ws, start_keyword, next_keyword, codes, code_map):
    start_row = -1
    for row in range(1, 300):
        val1 = str(ws.cell(row=row, column=1).value)
        val2 = str(ws.cell(row=row, column=2).value)
        if (start_keyword in val1) or (start_keyword in val2):
            start_row = row; break
    
    if start_row == -1: return 

    next_row = -1
    for row in range(start_row + 1, 300):
        val1 = str(ws.cell(row=row, column=1).value)
        val2 = str(ws.cell(row=row, column=2).value)
        if next_keyword and (next_keyword in val1 or next_keyword in val2):
            next_row = row; break
    
    if next_row == -1: next_row = start_row + 10 

    available_rows = next_row - start_row - 1
    
    unique_codes = []
    for c in codes:
        clean_c = c.replace(" ", "").strip().upper()
        if clean_c not in unique_codes: unique_codes.append(clean_c)
    
    required_rows = len(unique_codes)

    if required_rows > available_rows:
        rows_to_add = required_rows - available_rows
        ws.insert_rows(next_row, amount=rows_to_add)
    
    current_r = start_row + 1
    
    for code in unique_codes:
        ws.row_dimensions[current_r].height = 19
        ws.row_dimensions[current_r].hidden = False
        
        # [수정] safe_write 사용하여 병합 충돌 방지
        safe_write(ws, current_r, 2, code) # B열
        
        desc = code_map.get(code, "")
        safe_write(ws, current_r, 4, desc) # D열
        
        current_r += 1
    
    # 빈 행 처리 (다시 위치 찾기)
    real_next_row = -1
    for row in range(current_r, 300):
        val1 = str(ws.cell(row=row, column=1).value)
        val2 = str(ws.cell(row=row, column=2).value)
        if next_keyword and (next_keyword in val1 or next_keyword in val2):
            real_next_row = row; break
    
    if real_next_row == -1: real_next_row = current_r 

    for r in range(current_r, real_next_row):
        safe_write(ws, r, 2, "")
        safe_write(ws, r, 4, "")
        ws.row_dimensions[r].hidden = True

# 2. 파일 업로드
with st.expander("📂 필수 파일 업로드", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        master_data_file = st.file_uploader("1. 중앙 데이터 (master_data.xlsx)", type="xlsx")
        loaded_refs, folder_exists = get_reference_images()
        if folder_exists and loaded_refs:
            st.success(f"✅ 기준 그림 {len(loaded_refs)}개 로드됨")
        elif not folder_exists:
            st.warning("⚠️ 'reference_imgs' 폴더 필요")

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
            with st.spinner("PDF 분석 및 동적 양식 생성 중..."):
                
                new_files = []
                new_download_data = {}
                
                try: 
                    df_master = pd.read_excel(master_data_file, sheet_name=0)
                    code_map = {}
                    for idx, row in df_master.iterrows():
                        if pd.notna(row.iloc[0]):
                            code_key = str(row.iloc[0]).replace(" ", "").strip().upper()
                            desc_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                            code_map[code_key] = desc_val
                except: 
                    df_master = pd.DataFrame()
                    code_map = {}

                for uploaded_file in uploaded_files:
                    if option == "CFF(K)":
                        try:
                            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                            parsed_data = parse_pdf_ghs_logic(doc)
                            
                            template_file.seek(0)
                            dest_wb = load_workbook(io.BytesIO(template_file.read()))
                            dest_ws = dest_wb.active

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

                            dest_ws['B7'] = product_name_input
                            dest_ws['B10'] = product_name_input
                            
                            if parsed_data["hazard_cls"]:
                                b20_text = "\n".join(parsed_data["hazard_cls"])
                                safe_write(dest_ws, 20, 2, b20_text) # B20
                                dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

                            if parsed_data["signal_word"]:
                                safe_write(dest_ws, 24, 2, parsed_data["signal_word"]) # B24
                                dest_ws['B24'].alignment = Alignment(horizontal='center', vertical='center')

                            write_section_dynamic(dest_ws, "유해·위험문구", "예방", parsed_data["h_codes"], code_map)
                            write_section_dynamic(dest_ws, "예방", "대응", parsed_data["p_prev"], code_map)
                            write_section_dynamic(dest_ws, "대응", "저장", parsed_data["p_resp"], code_map)
                            write_section_dynamic(dest_ws, "저장", "폐기", parsed_data["p_stor"], code_map)
                            write_section_dynamic(dest_ws, "폐기", "3.", parsed_data["p_disp"], code_map)

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
                            
                            unique_images = {}
                            for key, img in collected_pil_images:
                                if key not in unique_images: unique_images[key] = img
                            
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
                    st.success("완료! 행 동적 추가 및 병합 오류 해결.")
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
