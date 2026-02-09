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
import math

# 1. 페이지 설정
st.set_page_config(page_title="MSDS 스마트 변환기", layout="wide")
st.title("MSDS 양식 변환기 (동일 라인 내용 복구 & 정밀 Gap 분석)")
st.markdown("---")

# --------------------------------------------------------------------------
# [스타일] 굴림 8pt
# --------------------------------------------------------------------------
FONT_STYLE = Font(name='굴림', size=8)
ALIGN_DATA = Alignment(horizontal='left', vertical='center', wrap_text=True)
ALIGN_TITLE = Alignment(horizontal='left', vertical='center', wrap_text=True)
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
# [핵심] PDF 라인 정밀 추출 (핀셋 노이즈 제거)
# --------------------------------------------------------------------------
def get_all_clean_lines_with_coords(doc):
    all_lines = []
    
    # 노이즈 패턴 (정규식) - 헤더/푸터 제거용
    noise_regexs = [
        r'^\s*\d+\s*/\s*\d+\s*$', # "3 / 11" 같은 페이지 번호 (라인 전체)
        r'물질안전보건자료', r'Material Safety Data Sheet', 
        r'PAGE', r'Ver\.\s*:?\s*\d+\.?\d*', r'발행일\s*:?.*', 
        r'주식회사\s*고려.*', r'Cff', r'Corea\s*flavors.*', 
        r'제\s*품\s*명\s*:?.*'
    ]
    
    global_y_offset = 0
    
    for page in doc:
        page_h = page.rect.height
        # 상하단 20px 안전하게 제외 (내용 잘림 최소화)
        clip_rect = fitz.Rect(0, 20, page.rect.width, page_h - 20)
        
        words = page.get_text("words", clip=clip_rect)
        words.sort(key=lambda w: (w[1], w[0])) 
        
        current_y = -100
        line_buffer = []
        page_lines = [] 
        
        for w in words:
            text, y0, y1 = w[4], w[1], w[3]
            if abs(y0 - current_y) > 3:
                if line_buffer:
                    full_text = " ".join([item[0] for item in line_buffer])
                    l_y0 = min([item[1] for item in line_buffer])
                    l_y1 = max([item[2] for item in line_buffer])
                    page_lines.append({'text': full_text, 'y0': l_y0, 'y1': l_y1})
                
                line_buffer = [(text, y0, y1)]
                current_y = y0
            else:
                line_buffer.append((text, y0, y1))
        
        if line_buffer:
            full_text = " ".join([item[0] for item in line_buffer])
            l_y0 = min([item[1] for item in line_buffer])
            l_y1 = max([item[2] for item in line_buffer])
            page_lines.append({'text': full_text, 'y0': l_y0, 'y1': l_y1})
            
        # 노이즈 필터링
        for line in page_lines:
            clean_txt = line['text']
            
            # 라인 전체가 노이즈인지 확인
            is_full_noise = False
            for pat in noise_regexs:
                if re.fullmatch(pat, clean_txt.strip(), re.IGNORECASE):
                    is_full_noise = True; break
            if is_full_noise: continue

            # 부분 노이즈 제거 (문장 중간의 헤더 등)
            for pat in noise_regexs:
                clean_txt = re.sub(pat, '', clean_txt, flags=re.IGNORECASE).strip()
            
            if clean_txt:
                line['text'] = clean_txt
                line['global_y0'] = line['y0'] + global_y_offset
                line['global_y1'] = line['y1'] + global_y_offset
                all_lines.append(line)
        
        global_y_offset += page_h
        
    return all_lines

# --------------------------------------------------------------------------
# [핵심 수정] 섹션 데이터 추출 (동일 라인 내용 보존 + Gap Logic)
# --------------------------------------------------------------------------
def extract_section_smart(all_lines, start_kw, end_kw):
    start_idx = -1
    end_idx = -1
    
    # Start 찾기
    for i, line in enumerate(all_lines):
        if start_kw in line['text']:
            start_idx = i
            break
    if start_idx == -1: return ""
    
    # End 찾기
    if isinstance(end_kw, str): end_kw = [end_kw]
    for i in range(start_idx + 1, len(all_lines)):
        line_text = all_lines[i]['text']
        for ek in end_kw:
            if ek in line_text:
                end_idx = i; break
        if end_idx != -1: break
    if end_idx == -1: end_idx = len(all_lines)
    
    # [FIX] start_idx 포함 (같은 줄에 있는 내용 살리기 위해)
    target_lines_raw = all_lines[start_idx : end_idx]
    if not target_lines_raw: return ""
    
    # 첫 줄(start_idx) 처리: 제목 제거하고 내용만 남기기
    first_line = target_lines_raw[0].copy() # 복사본 사용
    txt = first_line['text']
    
    # 제목(start_kw) 기준으로 자르기
    # 예: "나. 눈에 들어갔을 때 : 즉시 씻으시오" -> "즉시 씻으시오"
    if start_kw in txt:
        # split 후 뒷부분 가져오기
        parts = txt.split(start_kw, 1)
        if len(parts) > 1:
            content_part = parts[1].strip()
            # 앞부분 특수문자 제거 (: , - 등)
            content_part = re.sub(r"^[:\.\-\s]+", "", content_part)
            first_line['text'] = content_part
        else:
            first_line['text'] = "" # 내용 없음
    
    # 첫 줄 업데이트 (내용이 있을 때만 포함)
    target_lines = []
    if first_line['text'].strip():
        target_lines.append(first_line)
    
    # 나머지 줄 추가
    target_lines.extend(target_lines_raw[1:])
    
    if not target_lines: return ""
    
    # [Cleaning] 제목 잔여물 제거 (핀셋 방식)
    garbage_starts = [
        "에 접촉했을 때", "에 들어갔을 때", "들어갔을 때", "접촉했을 때", "했을 때", 
        "흡입했을 때", "먹었을 때", "주의사항", "내용물", 
        "취급요령", "저장방법", "보호구", "조치사항", "제거 방법",
        "소화제", "유해성", "로부터 생기는", "착용할 보호구", "및 예방조치",
        "방법", "경고표지 항목", "그림문자", "화학물질"
    ]
    
    cleaned_lines = []
    for line in target_lines:
        txt = line['text'].strip()
        
        # 줄의 시작 부분에 쓰레기가 있는지 확인
        for gb in garbage_starts:
            # 1. 아예 해당 문구로 시작하는 경우
            if txt.startswith(gb):
                txt = txt[len(gb):].strip()
            # 2. 문구와 매우 유사하게 시작하는 경우 (공백 등)
            elif gb in txt[:20]: 
                txt = txt.replace(gb, "").strip()
        
        # 특수문자 재정리
        txt = re.sub(r"^[:\.\)\s]+", "", txt)
        
        if txt:
            line['text'] = txt
            cleaned_lines.append(line)
            
    if not cleaned_lines: return ""

    # [Smart Merge] 간격(Gap) 기반 문장 병합
    final_text = ""
    if len(cleaned_lines) > 0:
        final_text = cleaned_lines[0]['text']
        for i in range(1, len(cleaned_lines)):
            prev = cleaned_lines[i-1]
            curr = cleaned_lines[i]
            
            # Gap 계산
            gap = curr['global_y0'] - prev['global_y1']
            
            # Gap이 작으면(6px 미만) 같은 문장 (Wrapping)
            # Gap이 크면(6px 이상) 다른 항목 (New line)
            if gap < 6.0: 
                final_text += " " + curr['text']
            else:
                final_text += "\n" + curr['text']
                
    return final_text

# --------------------------------------------------------------------------
# [함수] PDF 파싱 메인
# --------------------------------------------------------------------------
def parse_pdf_final(doc):
    all_lines = get_all_clean_lines_with_coords(doc)
    
    result = {
        "hazard_cls": [], "signal_word": "", "h_codes": [], 
        "p_prev": [], "p_resp": [], "p_stor": [], "p_disp": [],
        "composition_data": [], "sec4_to_7": {} 
    }

    # H/P 코드, 신호어 등 (기존 로직 유지)
    limit_y = 999999
    for line in all_lines:
        if "3. 구성성분" in line['text'] or "3. 성분" in line['text']:
            limit_y = line['global_y0']; break
            
    # 텍스트 풀 (섹션 2용)
    full_text_hp = "\n".join([l['text'] for l in all_lines if l['global_y0'] < limit_y])
    
    # 신호어
    for line in full_text_hp.split('\n'):
        if "신호어" in line:
            val = line.replace("신호어", "").replace(":", "").strip()
            if val in ["위험", "경고"]: result["signal_word"] = val
        elif line.strip() in ["위험", "경고"] and not result["signal_word"]:
            result["signal_word"] = line.strip()
    
    # 유해성 분류
    lines_hp = full_text_hp.split('\n')
    state = 0
    for l in lines_hp:
        l_ns = l.replace(" ", "")
        if "가.유해성" in l_ns and "분류" in l_ns: state=1; continue
        if "나.예방조치" in l_ns: state=0; continue
        if state==1 and l.strip():
            if "공급자" not in l and "회사명" not in l:
                result["hazard_cls"].append(l.strip())

    # H/P Code
    regex_code = re.compile(r"([HP]\s?\d{3}(?:\s*\+\s*[HP]\s?\d{3})*)")
    all_matches = regex_code.findall(full_text_hp)
    seen = set()
    if "P321" in full_text_hp and "P321" not in all_matches: all_matches.append("P321")
    for code_raw in all_matches:
        code = code_raw.replace(" ", "").upper()
        if code in seen: continue
        seen.add(code)
        if code.startswith("H"): result["h_codes"].append(code)
        elif code.startswith("P"):
            p = code.split("+")[0]
            if p.startswith("P2"): result["p_prev"].append(code)
            elif p.startswith("P3"): result["p_resp"].append(code)
            elif p.startswith("P4"): result["p_stor"].append(code)
            elif p.startswith("P5"): result["p_disp"].append(code)

    # 구성성분
    regex_cas = re.compile(r'\b(\d{2,7}-\d{2}-\d)\b')
    regex_conc = re.compile(r'\b(\d+)\s*~\s*(\d+)\b')
    in_comp = False
    for line in all_lines:
        txt = line['text']
        if "3." in txt and ("성분" in txt or "Composition" in txt): in_comp=True; continue
        if "4." in txt and ("응급" in txt or "First" in txt): in_comp=False; break
        if in_comp:
            if re.search(r'\d+\.\d+', txt): continue
            cas = regex_cas.search(txt)
            conc = regex_conc.search(txt)
            if cas:
                c_val = cas.group(1); cn_val = None
                if conc:
                    s, e = conc.group(1), conc.group(2)
                    if s=="1": s="0"
                    cn_val = f"{s} ~ {e}"
                result["composition_data"].append((c_val, cn_val))

    # 섹션 4~7 (Gap Logic + Fix: Same line content)
    data = {}
    data["B125"] = extract_section_smart(all_lines, "나. 눈", "다. 피부")
    data["B126"] = extract_section_smart(all_lines, "다. 피부", "라. 흡입")
    data["B127"] = extract_section_smart(all_lines, "라. 흡입", "마. 먹었을")
    data["B128"] = extract_section_smart(all_lines, "마. 먹었을", "바. 기타")
    data["B129"] = extract_section_smart(all_lines, "바. 기타", ["5.", "폭발"])
    data["B132"] = extract_section_smart(all_lines, "가. 적절한", "나. 화학물질")
    data["B133"] = extract_section_smart(all_lines, "나. 화학물질", "다. 화재진압")
    data["B134"] = extract_section_smart(all_lines, "다. 화재진압", ["6.", "누출"])
    data["B138"] = extract_section_smart(all_lines, "가. 인체를", "나. 환경을")
    data["B139"] = extract_section_smart(all_lines, "나. 환경을", "다. 정화")
    data["B140"] = extract_section_smart(all_lines, "다. 정화", ["7.", "취급"])
    data["B143"] = extract_section_smart(all_lines, "가. 안전취급", "나. 안전한")
    data["B144"] = extract_section_smart(all_lines, "나. 안전한", ["8.", "노출"])
    
    result["sec4_to_7"] = data
    return result

# --------------------------------------------------------------------------
# [함수] 포맷팅 & 유틸
# --------------------------------------------------------------------------
def get_description_smart(code, code_map):
    clean_code = str(code).replace(" ", "").upper().strip()
    if clean_code in code_map: return code_map[clean_code]
    if "+" in clean_code:
        parts = clean_code.split("+")
        found_texts = []
        for p in parts:
            if p in code_map: found_texts.append(code_map[p])
        if found_texts: return " ".join(found_texts)
    return ""

def safe_write_force(ws, row, col, value, center=False):
    cell = ws.cell(row=row, column=col)
    try: cell.value = value
    except AttributeError:
        try:
            for rng in list(ws.merged_cells.ranges):
                if cell.coordinate in rng:
                    ws.unmerge_cells(str(rng))
                    cell = ws.cell(row=row, column=col)
                    break
            cell.value = value
        except: pass
    if cell.font.name != '굴림': cell.font = FONT_STYLE
    if center: cell.alignment = ALIGN_CENTER
    else: cell.alignment = ALIGN_DATA

def calculate_smart_height_basic(text): 
    if not text: return 19.2
    explicit_lines = str(text).count('\n') + 1
    estimated_width_bytes = 72 
    current_bytes = 0; wrapped_lines = 1
    for char in str(text):
        if char == '\n': current_bytes = 0; wrapped_lines += 1; continue
        if '가' <= char <= '힣': current_bytes += 2
        else: current_bytes += 1
        if current_bytes >= estimated_width_bytes: wrapped_lines += 1; current_bytes = 0 
    final_lines = max(explicit_lines, wrapped_lines)
    if final_lines == 1: return 19.2
    elif final_lines == 2: return 23.3
    else: return 33.0

def format_and_calc_height_sec47(text):
    """
    [수정] 높이 계산 보정 (11줄 -> 120px)
    Wrapping 감도를 높여서(40자 기준) 줄 수를 넉넉하게 계산
    """
    if not text: return "", 19.2
    
    # 마침표 뒤 줄바꿈
    formatted_text = re.sub(r'(?<!\d)\.(?!\d)(?!\n)', '.\n', text)
    lines = [line.strip() for line in formatted_text.split('\n') if line.strip()]
    final_text = "\n".join(lines)
    
    # 높이 계산 (Wrapping 감도: 40자)
    char_limit_per_line = 40
    
    total_visual_lines = 0
    for line in lines:
        line_len = 0
        for ch in line:
            line_len += 2 if '가' <= ch <= '힣' else 1
        
        visual_lines = math.ceil(line_len / (char_limit_per_line * 2)) 
        if visual_lines == 0: visual_lines = 1
        total_visual_lines += visual_lines
    
    if total_visual_lines == 0: total_visual_lines = 1
    
    # (줄 수 * 10) + 10
    height = (total_visual_lines * 10) + 10
    
    return final_text, height

def fill_fixed_range(ws, start_row, end_row, codes, code_map):
    unique_codes = []; seen = set()
    for c in codes:
        clean = c.replace(" ", "").upper().strip()
        if clean not in seen: unique_codes.append(clean); seen.add(clean)
    limit = end_row - start_row + 1
    for i in range(limit):
        current_row = start_row + i
        if i < len(unique_codes):
            code = unique_codes[i]
            desc = get_description_smart(code, code_map)
            ws.row_dimensions[current_row].hidden = False
            final_height = calculate_smart_height_basic(desc)
            ws.row_dimensions[current_row].height = final_height
            safe_write_force(ws, current_row, 2, code, center=False)
            safe_write_force(ws, current_row, 4, desc, center=False)
        else:
            ws.row_dimensions[current_row].hidden = True
            safe_write_force(ws, current_row, 2, "") 
            safe_write_force(ws, current_row, 4, "")

def fill_composition_data(ws, comp_data, cas_to_name_map):
    start_row = 80; end_row = 123; limit = end_row - start_row + 1
    for i in range(limit):
        current_row = start_row + i
        if i < len(comp_data) and comp_data[i][1]:
            cas_no, concentration = comp_data[i]
            clean_cas = cas_no.replace(" ", "").strip()
            chem_name = cas_to_name_map.get(clean_cas, "")
            ws.row_dimensions[current_row].hidden = False
            ws.row_dimensions[current_row].height = 26.7
            safe_write_force(ws, current_row, 1, chem_name, center=True)
            safe_write_force(ws, current_row, 4, cas_no, center=True)
            safe_write_force(ws, current_row, 6, concentration, center=True)
        else:
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
            with st.spinner("동일 라인 내용 복구 및 정밀 보정 중..."):
                
                new_files = []
                new_download_data = {}
                
                code_map = {} 
                cas_name_map = {} 
                
                try:
                    xls = pd.ExcelFile(master_data_file)
                    target_sheet = None
                    for sheet in xls.sheet_names:
                        if "위험" in sheet and "안전" in sheet: target_sheet = sheet; break
                    if not target_sheet:
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
                    
                    sheet_kor = None
                    for sheet in xls.sheet_names:
                        if "국문" in sheet: sheet_kor = sheet; break
                    if sheet_kor:
                        df_kor = pd.read_excel(master_data_file, sheet_name=sheet_kor)
                        for _, row in df_kor.iterrows():
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
                            parsed_data = parse_pdf_final(doc)
                            
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
                            
                            # 3. 유해성 & 신호어
                            if parsed_data["hazard_cls"]:
                                clean_hazard_text = "\n".join([line for line in parsed_data["hazard_cls"] if line.strip()])
                                safe_write_force(dest_ws, 20, 2, clean_hazard_text, center=False)
                                dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

                            signal_final = parsed_data["signal_word"] if parsed_data["signal_word"] else ""
                            safe_write_force(dest_ws, 24, 2, signal_final, center=False) 

                            # 4. H/P 코드
                            fill_fixed_range(dest_ws, 25, 36, parsed_data["h_codes"], code_map)
                            fill_fixed_range(dest_ws, 38, 49, parsed_data["p_prev"], code_map)
                            fill_fixed_range(dest_ws, 50, 63, parsed_data["p_resp"], code_map)
                            fill_fixed_range(dest_ws, 64, 69, parsed_data["p_stor"], code_map)
                            fill_fixed_range(dest_ws, 70, 72, parsed_data["p_disp"], code_map)

                            # 5. 구성성분
                            fill_composition_data(dest_ws, parsed_data["composition_data"], cas_name_map)

                            # 6. 섹션 4~7 데이터 쓰기
                            sec_data = parsed_data["sec4_to_7"]
                            import openpyxl.utils
                            
                            for cell_addr, raw_text in sec_data.items():
                                formatted_txt, row_h = format_and_calc_height_sec47(raw_text)
                                
                                try:
                                    col_str = re.match(r"([A-Z]+)", cell_addr).group(1)
                                    row_num = int(re.search(r"(\d+)", cell_addr).group(1))
                                    col_idx = openpyxl.utils.column_index_from_string(col_str)
                                    
                                    # 초기화
                                    safe_write_force(dest_ws, row_num, col_idx, "")
                                    
                                    if formatted_txt:
                                        # B열 쓰기
                                        safe_write_force(dest_ws, row_num, col_idx, formatted_txt, center=False)
                                        dest_ws.row_dimensions[row_num].height = row_h
                                        
                                        # A열 정렬 (왼쪽+수직중앙)
                                        try:
                                            cell_a = dest_ws.cell(row=row_num, column=1)
                                            cell_a.alignment = ALIGN_TITLE
                                        except: pass

                                except Exception as e:
                                    print(f"Cell write error: {cell_addr} - {e}")

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
                
                if 'df_code' in locals(): del df_code
                if 'df_kor' in locals(): del df_kor
                if 'doc' in locals(): doc.close()
                if 'dest_wb' in locals(): del dest_wb
                if 'output' in locals(): del output
                gc.collect()

                if new_files:
                    st.success("완료! 내용 소실 방지 및 높이 계산 정밀 보정 성공.")
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
