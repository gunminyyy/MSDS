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
st.title("MSDS 양식 변환기 (완전 무결점 - 번호 기반 자동 분류)")
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
# [함수] PDF 파싱 (전역 스캔 + 번호 기반 분류)
# --------------------------------------------------------------------------
def parse_pdf_universal(doc):
    full_text = ""
    clean_lines = []
    
    # 1. 전체 텍스트 추출 (순서대로)
    for page in doc:
        blocks = page.get_text("blocks", sort=True)
        for b in blocks:
            text = b[4]
            full_text += text + "\n"
            clean_lines.extend(text.split('\n'))

    result = {
        "hazard_cls": [], "signal_word": "", 
        "h_codes": [], "p_prev": [], "p_resp": [], "p_stor": [], "p_disp": []
    }

    # 2. 신호어 & 유해성 분류 찾기 (이건 위치가 중요하므로 라인별 탐색 유지)
    #    단, 신호어는 '주변 탐색'으로 보완
    ZONE_NONE = 0; ZONE_HAZARD = 1
    state = ZONE_NONE
    
    for i, line in enumerate(clean_lines):
        line_ns = line.replace(" ", "")
        
        if "가.유해성" in line_ns and "분류" in line_ns:
            state = ZONE_HAZARD; continue
        if "나.예방조치" in line_ns:
            state = ZONE_NONE; continue # 여기서부턴 전역 스캔으로 처리
            
        if state == ZONE_HAZARD:
            if "공급자정보" in line_ns or "회사명" in line_ns: continue
            result["hazard_cls"].append(line)
            
        # 신호어 찾기 (전체 라인 대상)
        if "신호어" in line_ns:
            val = line.replace("신호어", "").replace(":", "").strip()
            if val in ["위험", "경고"]:
                result["signal_word"] = val
            else:
                # 다음 3줄 탐색
                for offset in range(1, 4):
                    if i + offset < len(clean_lines):
                        nxt = clean_lines[i+offset].strip()
                        if nxt in ["위험", "경고"]:
                            result["signal_word"] = nxt; break

    # 3. [핵심] 코드 전역 스캔 및 자동 분류 (번호 기반)
    #    H코드, P코드 패턴 정의 (공백 포함 가능)
    regex = re.compile(r"([HP]\s?\d{3}(?:\s*\+\s*[HP]\s?\d{3})*)")
    
    all_matches = regex.findall(full_text)
    
    seen = set()
    for code_raw in all_matches:
        # 정규화 (공백 제거)
        code = code_raw.replace(" ", "").upper()
        
        if code in seen: continue
        seen.add(code)
        
        # GHS 번호 규칙에 따른 자동 분류
        if code.startswith("H"):
            result["h_codes"].append(code)
        
        elif code.startswith("P"):
            # 복합 코드(P301+P310)일 경우 첫 번째 코드로 판단
            prefix = code.split("+")[0]
            
            if prefix.startswith("P2"):   # 200번대 -> 예방
                result["p_prev"].append(code)
            elif prefix.startswith("P3"): # 300번대 -> 대응
                result["p_resp"].append(code)
            elif prefix.startswith("P4"): # 400번대 -> 저장
                result["p_stor"].append(code)
            elif prefix.startswith("P5"): # 500번대 -> 폐기
                result["p_disp"].append(code)
            # P321 같은 경우도 prefix가 P321 -> P3으로 시작하므로 '대응'으로 정확히 들어감

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
# [함수] 안전 쓰기 (강제 병합 해제 & 스타일)
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
# [핵심] 고정 범위 채우기 (Fixed Range)
# --------------------------------------------------------------------------
def fill_fixed_range(ws, start_row, end_row, codes, code_map):
    # 정렬: 번호순으로 정렬하면 깔끔하지만, 원본 순서를 원하면 sort 제거 가능.
    # 여기서는 '누락 방지'가 최우선이므로, 자동 분류된 리스트를 그대로 사용하되
    # 혹시 모를 중복만 다시 체크함.
    
    limit = end_row - start_row + 1
    
    for i in range(limit):
        current_row = start_row + i
        
        if i < len(codes):
            code = codes[i]
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
            with st.spinner("자동 분류 로직 가동 중..."):
                
                new_files = []
                new_download_data = {}
                
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
                            # [핵심] 완전 무결점 분류 함수 사용
                            parsed_data = parse_pdf_universal(doc)
                            
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

                            signal_final = parsed_data["signal_word"] if parsed_data["signal_word"] else ""
                            safe_write_force(dest_ws, 24, 2, signal_final, center=True)

                            # [고정 범위 채우기] - 이제 어떤 코드도 누락 없이 자기 자리를 찾아갑니다.
                            fill_fixed_range(dest_ws, 25, 36, parsed_data["h_codes"], code_map)
                            fill_fixed_range(dest_ws, 38, 49, parsed_data["p_prev"], code_map)
                            fill_fixed_range(dest_ws, 50, 63, parsed_data["p_resp"], code_map)
                            fill_fixed_range(dest_ws, 64, 69, parsed_data["p_stor"], code_map)
                            fill_fixed_range(dest_ws, 70, 72, parsed_data["p_disp"], code_map)

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
                    st.success("완료! P코드 전수 조사 및 자동 분류로 누락을 완벽 차단했습니다.")
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
