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
st.title("MSDS 양식 변환기 (구성성분표 완벽 구현)")
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
# [함수] PDF 파싱 (3번 섹션 정밀 추출 포함)
# --------------------------------------------------------------------------
def parse_pdf_full_logic(doc):
    full_text = ""
    clean_lines = []
    
    # 1. 텍스트 추출 (전체 페이지)
    for page in doc:
        blocks = page.get_text("blocks", sort=True)
        for b in blocks:
            text = b[4]
            full_text += text + "\n"
            lines = text.split('\n')
            for line in lines:
                line_str = line.strip()
                if not line_str: continue 
                is_noise = False
                for kw in ["물질안전보건자료", "MSDS", "Material Safety", "PAGE", "Ver.", "발행일"]:
                    if kw in line_str: is_noise = True; break
                if not is_noise: clean_lines.append(line_str)

    result = {
        "hazard_cls": [], "signal_word": "", 
        "h_codes": [], "p_prev": [], "p_resp": [], "p_stor": [], "p_disp": [],
        "composition_data": [] # (CAS, Concentration)
    }

    # 2. 유해성 분류 및 신호어
    ZONE_NONE = 0; ZONE_HAZARD = 1
    state = ZONE_NONE
    for i, line in enumerate(clean_lines):
        line_ns = line.replace(" ", "")
        if "가.유해성" in line_ns and "분류" in line_ns:
            state = ZONE_HAZARD; continue
        if "나.예방조치" in line_ns:
            state = ZONE_NONE; continue 
            
        if state == ZONE_HAZARD:
            if "공급자정보" in line_ns or "회사명" in line_ns: continue
            if line.strip(): result["hazard_cls"].append(line.strip())
            
        if "신호어" in line_ns:
            val = line.replace("신호어", "").replace(":", "").strip()
            if val in ["위험", "경고"]:
                result["signal_word"] = val
            else:
                for offset in range(1, 4):
                    if i + offset < len(clean_lines):
                        nxt = clean_lines[i+offset].strip()
                        if nxt in ["위험", "경고"]:
                            result["signal_word"] = nxt; break

    # 3. H/P 코드 스캔 (3번 섹션 전까지만)
    limit_index = len(full_text)
    match_sec3 = re.search(r"3\.\s*(구성성분|Composition)", full_text)
    match_sec4 = re.search(r"4\.\s*(응급조치|First)", full_text)
    
    if match_sec3: limit_index = match_sec3.start()
    
    target_text_hp = full_text[:limit_index]
    regex_code = re.compile(r"([HP]\s?\d{3}(?:\s*\+\s*[HP]\s?\d{3})*)")
    all_matches = regex_code.findall(target_text_hp)
    
    seen = set()
    if "P321" in target_text_hp and "P321" not in all_matches: all_matches.append("P321")

    for code_raw in all_matches:
        code = code_raw.replace(" ", "").upper()
        if code in seen: continue
        seen.add(code)
        if code.startswith("H"): result["h_codes"].append(code)
        elif code.startswith("P"):
            prefix = code.split("+")[0]
            if prefix.startswith("P2"): result["p_prev"].append(code)
            elif prefix.startswith("P3"): result["p_resp"].append(code)
            elif prefix.startswith("P4"): result["p_stor"].append(code)
            elif prefix.startswith("P5"): result["p_disp"].append(code)

    # 4. 구성성분 추출 (3번 ~ 4번 사이)
    if match_sec3 and match_sec4:
        start_idx = match_sec3.start()
        end_idx = match_sec4.start()
        comp_text = full_text[start_idx:end_idx]
        
        comp_lines = comp_text.split('\n')
        
        # CAS No: 숫자-숫자-숫자
        regex_cas = re.compile(r'\b(\d{2,7}-\d{2}-\d)\b')
        # 함유량: 정수 ~ 정수 (띄어쓰기 허용)
        regex_conc = re.compile(r'\b(\d+)\s*~\s*(\d+)\b')
        
        for line in comp_lines:
            # 소수점(.)이 포함된 숫자가 있으면 무시 (예: 1.5, 0.1)
            # 단, 문장 끝의 마침표는 제외해야 하므로 숫자 사이에 점이 있는지 확인
            if re.search(r'\d+\.\d+', line):
                continue

            cas_match = regex_cas.search(line)
            conc_match = regex_conc.search(line)
            
            if cas_match:
                cas_val = cas_match.group(1)
                conc_val = None
                
                if conc_match:
                    start_val = conc_match.group(1)
                    end_val = conc_match.group(2)
                    
                    # 수치 보정: 1 ~ 5 -> 0 ~ 5
                    if start_val == "1": start_val = "0"
                    
                    conc_val = f"{start_val} ~ {end_val}"
                
                result["composition_data"].append((cas_val, conc_val))

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
# [함수] 행 높이 계산기
# --------------------------------------------------------------------------
def calculate_smart_height(text):
    if not text: return 19.2
    explicit_lines = str(text).count('\n') + 1
    estimated_width_bytes = 72 
    current_bytes = 0
    wrapped_lines = 1
    for char in str(text):
        if char == '\n':
            current_bytes = 0; wrapped_lines += 1; continue
        if '가' <= char <= '힣': current_bytes += 2
        else: current_bytes += 1
        if current_bytes >= estimated_width_bytes:
            wrapped_lines += 1; current_bytes = 0 
    final_lines = max(explicit_lines, wrapped_lines)
    
    if final_lines == 1: return 19.2
    elif final_lines == 2: return 23.3
    else: return 33.0

# --------------------------------------------------------------------------
# [함수] 고정 범위 채우기
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
            final_height = calculate_smart_height(desc)
            ws.row_dimensions[current_row].height = final_height
            safe_write_force(ws, current_row, 2, code, center=False)
            safe_write_force(ws, current_row, 4, desc, center=False)
        else:
            ws.row_dimensions[current_row].hidden = True
            safe_write_force(ws, current_row, 2, "") 
            safe_write_force(ws, current_row, 4, "")

# --------------------------------------------------------------------------
# [신규 함수] 구성성분 채우기 (80~123행)
# --------------------------------------------------------------------------
def fill_composition_data(ws, comp_data, cas_to_name_map):
    """
    comp_data: [(CAS, Concentration), ...]
    cas_to_name_map: { 'CAS_NO': 'Chemical Name' }
    Range: 80 ~ 123
    """
    start_row = 80
    end_row = 123
    limit = end_row - start_row + 1
    
    for i in range(limit):
        current_row = start_row + i
        
        # 데이터가 있고, 함유량이 존재할 때만 표시
        if i < len(comp_data) and comp_data[i][1]:
            cas_no, concentration = comp_data[i]
            
            # 물질명 매핑
            clean_cas = cas_no.replace(" ", "").strip()
            chem_name = cas_to_name_map.get(clean_cas, "")
            
            ws.row_dimensions[current_row].hidden = False
            ws.row_dimensions[current_row].height = 26.7 # [요청] 높이 고정
            
            safe_write_force(ws, current_row, 1, chem_name, center=True) # A열
            safe_write_force(ws, current_row, 4, cas_no, center=True)    # D열
            safe_write_force(ws, current_row, 6, concentration, center=True) # F열
                
        else:
            # 남는 행 또는 함유량이 없는 행은 숨김
            ws.row_dimensions[current_row].hidden = True
            safe_write_force(ws, current_row, 1, "")
            safe_write_force(ws, current_row, 4, "")
            safe_write_force(ws, current_row, 6, "")

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
            with st.spinner("분석 및 변환 중..."):
                
                new_files = []
                new_download_data = {}
                
                code_map = {} # H/P 코드용
                cas_name_map = {} # CAS -> 물질명 매핑용
                
                try:
                    xls = pd.ExcelFile(master_data_file)
                    
                    # (1) H/P 코드 시트 로드
                    target_sheet = None
                    for sheet in xls.sheet_names:
                        if "위험" in sheet and "안전" in sheet: target_sheet = sheet; break
                    
                    if not target_sheet:
                         # fallback
                         for sheet in xls.sheet_names:
                            df_tmp = pd.read_excel(master_data_file, sheet_name=sheet, nrows=5)
                            if 'CODE' in [str(c).upper() for c in df_tmp.columns]: target_sheet = sheet; break
                    
                    if target_sheet:
                        df_code = pd.read_excel(master_data_file, sheet_name=target_sheet)
                        df_code.columns = [str(c).replace(" ", "").upper() for c in df_code.columns]
                        col_c = 'CODE'; col_k = 'K'
                        for _, row in df_code.iterrows():
                            if pd.notna(row[col_c]):
                                code_map[str(row[col_c]).replace(" ","").upper().strip()] = str(row[col_k]).strip()
                    
                    # (2) 국문 시트 로드 (CAS 매핑)
                    sheet_kor = None
                    for sheet in xls.sheet_names:
                        if "국문" in sheet: sheet_kor = sheet; break
                    
                    if sheet_kor:
                        # 국문 시트의 A열(CAS), B열(물질명) 사용
                        df_kor = pd.read_excel(master_data_file, sheet_name=sheet_kor)
                        # 컬럼 인덱스 0, 1 사용 (안전)
                        for _, row in df_kor.iterrows():
                            # 데이터가 있는 행만
                            val_cas = row.iloc[0]
                            val_name = row.iloc[1]
                            if pd.notna(val_cas):
                                c = str(val_cas).replace(" ", "").strip()
                                n = str(val_name).strip() if pd.notna(val_name) else ""
                                cas_name_map[c] = n
                                
                except Exception as e:
                    st.error(f"데이터 로드 오류: {e}")

                for uploaded_file in uploaded_files:
                    if option == "CFF(K)":
                        try:
                            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                            parsed_data = parse_pdf_full_logic(doc)
                            
                            template_file.seek(0)
                            dest_wb = load_workbook(io.BytesIO(template_file.read()))
                            dest_ws = dest_wb.active

                            # 1. 수식 청소
                            for row in dest_ws.iter_rows():
                                for cell in row:
                                    if isinstance(cell, MergedCell): continue
                                    if cell.data_type == 'f' and "ingredients" in str(cell.value):
                                        cell.value = ""

                            # 2. 기본 정보
                            safe_write_force(dest_ws, 7, 2, product_name_input, center=True)
                            safe_write_force(dest_ws, 10, 2, product_name_input, center=True)
                            
                            # 3. 유해성 분류
                            if parsed_data["hazard_cls"]:
                                clean_hazard_text = "\n".join([line for line in parsed_data["hazard_cls"] if line.strip()])
                                safe_write_force(dest_ws, 20, 2, clean_hazard_text, center=False)
                                dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

                            # 4. 신호어
                            signal_final = parsed_data["signal_word"] if parsed_data["signal_word"] else ""
                            safe_write_force(dest_ws, 24, 2, signal_final, center=False) 

                            # 5. H/P 코드
                            fill_fixed_range(dest_ws, 25, 36, parsed_data["h_codes"], code_map)
                            fill_fixed_range(dest_ws, 38, 49, parsed_data["p_prev"], code_map)
                            fill_fixed_range(dest_ws, 50, 63, parsed_data["p_resp"], code_map)
                            fill_fixed_range(dest_ws, 64, 69, parsed_data["p_stor"], code_map)
                            fill_fixed_range(dest_ws, 70, 72, parsed_data["p_disp"], code_map)

                            # 6. 구성성분 (CAS & 함유량)
                            fill_composition_data(dest_ws, parsed_data["composition_data"], cas_name_map)

                            # 7. 이미지 삽입
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
                
                # [Error Fix] 변수 존재 여부 확인 후 삭제
                if 'df_code' in locals(): del df_code
                if 'df_kor' in locals(): del df_kor
                if 'doc' in locals(): doc.close()
                if 'dest_wb' in locals(): del dest_wb
                if 'output' in locals(): del output
                gc.collect()

                if new_files:
                    st.success("완료! 구성성분표 및 수치 필터링까지 완벽하게 처리되었습니다.")
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
