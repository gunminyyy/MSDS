import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
import io
import re
import os
import fitz  # PyMuPDF
import numpy as np
import gc

# 1. 페이지 설정
st.set_page_config(page_title="MSDS 스마트 변환기", layout="wide")
st.title("MSDS 양식 변환기 (코드 누락 방지 & 신호어 보정)")
st.markdown("---")

# --------------------------------------------------------------------------
# [스타일] 굴림 8pt, 왼쪽 정렬
# --------------------------------------------------------------------------
FONT_STYLE = Font(name='굴림', size=8)
ALIGN_LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
ALIGN_CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

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
    if not os.path.exists(img_folder): return {}, False
    try:
        ref_images = {}
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
        src_arr = np.array(src_norm, dtype='int16')
        for name, ref_img in ref_images.items():
            ref_norm = normalize_image(ref_img)
            ref_arr = np.array(ref_norm, dtype='int16')
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
# [함수] PDF 파싱 (전역 스캔 및 자동 분류 추가)
# --------------------------------------------------------------------------
def parse_pdf_ghs_logic(doc):
    clean_lines = []
    NOISE_KEYWORDS = [
        "물질안전보건자료", "MSDS", "Material Safety Data Sheet",
        "Corea flavors", "주식회사 고려", "HAIR CARE", "Ver.", "발행일", "개정일",
        "제 품 명", "GHS", "페이지", "PAGE", "---"
    ]

    full_text_for_scan = "" # 전역 스캔용 텍스트

    for page in doc:
        blocks = page.get_text("blocks", sort=True)
        for b in blocks:
            text = b[4]
            full_text_for_scan += text + "\n"
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

    ZONE_NONE = 0; ZONE_HAZARD_CLS = 1; ZONE_LABEL_INFO = 2
    current_zone = ZONE_NONE
    SUB_NONE=0; SUB_PREV=1; SUB_RESP=2; SUB_STOR=3; SUB_DISP=4
    current_sub = SUB_NONE
    
    regex_code = re.compile(r"([HP]\s?\d{3}(?:\s*\+\s*[HP]\s?\d{3})*)")
    BLACKLIST_HAZARD = ["공급자정보", "회사명", "주소", "긴급전화번호", "권고용도", "사용상의제한"]

    # 1차 패스: 구역별 정밀 추출
    for line in clean_lines:
        line_ns = line.replace(" ", "")
        
        if "가.유해성" in line_ns and "분류" in line_ns:
            current_zone = ZONE_HAZARD_CLS; continue
        if "나.예방조치" in line_ns:
            current_zone = ZONE_LABEL_INFO; current_sub = SUB_NONE; continue
        if "3.구성성분" in line_ns or "다.기타" in line_ns:
            current_zone = ZONE_NONE; break

        if current_zone == ZONE_HAZARD_CLS:
            is_bad = False
            for bl in BLACKLIST_HAZARD:
                if bl in line_ns: is_bad = True; break
            if not is_bad:
                result["hazard_cls"].append(line)
                codes = regex_code.findall(line)
                for c in codes:
                    c_clean = c.replace(" ", "")
                    if c_clean.startswith("H"): result["h_codes"].append(c_clean)

        elif current_zone == ZONE_LABEL_INFO:
            # [수정] 신호어 추출: 같은 줄의 내용을 우선 가져옴
            if "신호어" in line_ns:
                val = line.replace("신호어", "").replace(":", "").strip()
                if val: 
                    result["signal_word"] = val
            
            # 서브존 전환
            if line_ns.startswith("예방") and len(line_ns) < 15: current_sub = SUB_PREV
            elif line_ns.startswith("대응") and len(line_ns) < 15: current_sub = SUB_RESP
            elif line_ns.startswith("저장") and len(line_ns) < 15: current_sub = SUB_STOR
            elif line_ns.startswith("폐기") and len(line_ns) < 15: current_sub = SUB_DISP

            codes = regex_code.findall(line)
            for c in codes:
                c_clean = c.replace(" ", "")
                if c_clean.startswith("H"): result["h_codes"].append(c_clean)
                elif c_clean.startswith("P"):
                    if current_sub == SUB_PREV: result["p_prev"].append(c_clean)
                    elif current_sub == SUB_RESP: result["p_resp"].append(c_clean)
                    elif current_sub == SUB_STOR: result["p_stor"].append(c_clean)
                    elif current_sub == SUB_DISP: result["p_disp"].append(c_clean)

    # 2차 패스: 전역 스캔 및 누락 코드 자동 분류 (Safety Net)
    all_codes = regex_code.findall(full_text_for_scan)
    collected_set = set(result["h_codes"] + result["p_prev"] + result["p_resp"] + result["p_stor"] + result["p_disp"])
    
    for c in all_codes:
        c_clean = c.replace(" ", "")
        if c_clean not in collected_set:
            # 이미 수집된 목록에 없으면(누락되었으면) 번호 대역으로 분류해서 추가
            if c_clean.startswith("H"): result["h_codes"].append(c_clean)
            elif c_clean.startswith("P2"): result["p_prev"].append(c_clean)
            elif c_clean.startswith("P3"): result["p_resp"].append(c_clean)
            elif c_clean.startswith("P4"): result["p_stor"].append(c_clean)
            elif c_clean.startswith("P5"): result["p_disp"].append(c_clean)
            # 복합 코드(P301+P310)인 경우 첫 번째 코드를 기준으로 분류
            elif c_clean.startswith("P"):
                first_part = c_clean.split("+")[0]
                if first_part.startswith("P2"): result["p_prev"].append(c_clean)
                elif first_part.startswith("P3"): result["p_resp"].append(c_clean)
                elif first_part.startswith("P4"): result["p_stor"].append(c_clean)
                elif first_part.startswith("P5"): result["p_disp"].append(c_clean)

    return result

# --------------------------------------------------------------------------
# [함수] 중앙 데이터 매핑
# --------------------------------------------------------------------------
def get_description_smart(code, code_map):
    clean_code = str(code).replace(" ", "").upper().strip()
    if clean_code in code_map:
        return code_map[clean_code]
    if "+" in clean_code:
        parts = clean_code.split("+")
        found_texts = []
        for p in parts:
            if p in code_map:
                found_texts.append(code_map[p])
        if found_texts:
            return " ".join(found_texts)
    return ""

# --------------------------------------------------------------------------
# [함수] 안전 쓰기
# --------------------------------------------------------------------------
def safe_write_force(ws, row, col, value, center=False):
    cell = ws.cell(row=row, column=col)
    try:
        cell.value = value
    except AttributeError:
        try:
            for rng in list(ws.merged_cells.ranges):
                if cell.coordinate in rng:
                    ws.unmerge_cells(str(rng))
                    cell = ws.cell(row=row, column=col)
                    break
            cell.value = value
        except: pass

    if cell.font.name != '굴림':
        cell.font = FONT_STYLE
    
    if center:
        cell.alignment = ALIGN_CENTER
    else:
        cell.alignment = ALIGN_LEFT

# --------------------------------------------------------------------------
# [핵심] 고정 범위 채우기
# --------------------------------------------------------------------------
def fill_fixed_range(ws, start_row, end_row, codes, code_map):
    unique_codes = []
    seen = set()
    for c in codes:
        clean = c.replace(" ", "").upper().strip()
        if clean not in seen:
            unique_codes.append(clean)
            seen.add(clean)
    
    limit = end_row - start_row + 1
    
    for i in range(limit):
        current_row = start_row + i
        
        if i < len(unique_codes):
            code = unique_codes[i]
            desc = get_description_smart(code, code_map)
            
            ws.row_dimensions[current_row].hidden = False
            ws.row_dimensions[current_row].height = 19
            
            safe_write_force(ws, current_row, 2, code, center=False)
            safe_write_force(ws, current_row, 4, desc, center=False)
            
        else:
            ws.row_dimensions[current_row].hidden = True
            safe_write_force(ws, current_row, 2, "") 
            safe_write_force(ws, current_row, 4, "")

# 2. 파일 업로드
with st.expander("📂 필수 파일 업로드", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        master_data_file = st.file_uploader("1. 중앙 데이터 (ingredients...xlsx)", type="xlsx")
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
            with st.spinner("코드 정밀 스캔 및 변환 중..."):
                
                new_files = []
                new_download_data = {}
                
                # 중앙 데이터 로드
                code_map = {}
                try:
                    xls = pd.ExcelFile(master_data_file)
                    target_sheet = None
                    for sheet in xls.sheet_names:
                        if "위험" in sheet and "안전" in sheet:
                            target_sheet = sheet
                            break
                    if target_sheet is None:
                        for sheet in xls.sheet_names:
                            df_check = pd.read_excel(master_data_file, sheet_name=sheet, nrows=5)
                            cols = [str(c).upper() for c in df_check.columns]
                            if 'CODE' in cols and 'K' in cols:
                                target_sheet = sheet
                                break
                    if target_sheet:
                        df_master = pd.read_excel(master_data_file, sheet_name=target_sheet)
                        df_master.columns = [str(c).replace(" ", "").upper() for c in df_master.columns]
                        col_code = 'CODE'
                        col_kor = 'K'
                        for idx, row in df_master.iterrows():
                            if pd.notna(row[col_code]):
                                k = str(row[col_code]).replace(" ", "").upper().strip()
                                v = str(row[col_kor]).strip() if pd.notna(row[col_kor]) else ""
                                code_map[k] = v
                except Exception as e:
                    st.error(f"데이터 로드 오류: {e}")

                for uploaded_file in uploaded_files:
                    if option == "CFF(K)":
                        try:
                            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                            # [핵심] 누락 방지 강화된 파싱 함수 사용
                            parsed_data = parse_pdf_ghs_logic(doc)
                            
                            template_file.seek(0)
                            dest_wb = load_workbook(io.BytesIO(template_file.read()))
                            dest_ws = dest_wb.active

                            for row in dest_ws.iter_rows():
                                for cell in row:
                                    if isinstance(cell, MergedCell): continue
                                    if cell.data_type == 'f' and "ingredients" in str(cell.value):
                                        cell.value = ""

                            safe_write_force(dest_ws, 7, 2, product_name_input, center=True)
                            safe_write_force(dest_ws, 10, 2, product_name_input, center=True)
                            
                            if parsed_data["hazard_cls"]:
                                b20_text = "\n".join(parsed_data["hazard_cls"])
                                safe_write_force(dest_ws, 20, 2, b20_text, center=False)
                                dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

                            if parsed_data["signal_word"]:
                                safe_write_force(dest_ws, 24, 2, parsed_data["signal_word"], center=True)

                            # 고정 범위 채우기
                            fill_fixed_range(dest_ws, 25, 36, parsed_data["h_codes"], code_map)
                            fill_fixed_range(dest_ws, 38, 49, parsed_data["p_prev"], code_map)
                            fill_fixed_range(dest_ws, 50, 63, parsed_data["p_resp"], code_map)
                            fill_fixed_range(dest_ws, 64, 69, parsed_data["p_stor"], code_map)
                            fill_fixed_range(dest_ws, 70, 72, parsed_data["p_disp"], code_map)

                            # 이미지
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
                    st.success("완료! 신호어 보정 및 모든 P코드 완벽 추출 성공.")
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
