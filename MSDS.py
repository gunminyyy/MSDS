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
st.title("MSDS 양식 변환기 (글자 잘림 방지 & 정밀 정제)")
st.markdown("---")

# --------------------------------------------------------------------------
# [스타일]
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
# [핵심] 시각적 행 클러스터링
# --------------------------------------------------------------------------
def get_clustered_lines(doc):
    all_lines = []
    
    noise_regexs = [
        r'^\s*\d+\s*/\s*\d+\s*$', 
        r'물질안전보건자료', r'Material Safety Data Sheet', 
        r'PAGE', r'Ver\.\s*:?\s*\d+\.?\d*', r'발행일\s*:?.*', 
        r'주식회사\s*고려.*', r'Cff', r'Corea\s*flavors.*', 
        r'제\s*품\s*명\s*:?.*'
    ]
    
    global_y_offset = 0
    
    for page in doc:
        page_h = page.rect.height
        clip_rect = fitz.Rect(0, 60, page.rect.width, page_h - 50)
        
        words = page.get_text("words", clip=clip_rect)
        words.sort(key=lambda w: w[1]) 
        
        rows = []
        if words:
            current_row = [words[0]]
            row_base_y = words[0][1]
            
            for w in words[1:]:
                if abs(w[1] - row_base_y) < 8:
                    current_row.append(w)
                else:
                    current_row.sort(key=lambda x: x[0])
                    rows.append(current_row)
                    current_row = [w]
                    row_base_y = w[1]
            
            if current_row:
                current_row.sort(key=lambda x: x[0])
                rows.append(current_row)
        
        for row in rows:
            line_text = " ".join([w[4] for w in row])
            
            is_noise = False
            for pat in noise_regexs:
                if re.search(pat, line_text, re.IGNORECASE):
                    is_noise = True; break
            
            if not is_noise:
                avg_y = sum([w[1] for w in row]) / len(row)
                all_lines.append({
                    'text': line_text,
                    'global_y0': avg_y + global_y_offset,
                    'global_y1': (sum([w[3] for w in row]) / len(row)) + global_y_offset
                })
        
        global_y_offset += page_h
        
    return all_lines

# --------------------------------------------------------------------------
# [핵심 수정] 섹션 추출 (정규식 기반 안전한 잔여물 제거)
# --------------------------------------------------------------------------
def extract_section_smart(all_lines, start_kw, end_kw):
    start_idx = -1
    end_idx = -1
    
    # 1. 시작점 찾기 (공백 무시)
    clean_start_kw = start_kw.replace(" ", "")
    for i, line in enumerate(all_lines):
        if clean_start_kw in line['text'].replace(" ", ""):
            start_idx = i
            break
    if start_idx == -1: return ""
    
    # 2. 종료점 찾기
    if isinstance(end_kw, str): end_kw = [end_kw]
    clean_end_kws = [k.replace(" ", "") for k in end_kw]
    
    for i in range(start_idx + 1, len(all_lines)):
        line_clean = all_lines[i]['text'].replace(" ", "")
        for cek in clean_end_kws:
            if cek in line_clean:
                end_idx = i; break
        if end_idx != -1: break
    if end_idx == -1: end_idx = len(all_lines)
    
    target_lines_raw = all_lines[start_idx : end_idx]
    if not target_lines_raw: return ""
    
    # 3. 첫 줄 제목 제거
    first_line = target_lines_raw[0].copy()
    txt = first_line['text']
    escaped_kw = re.escape(start_kw)
    pattern_str = escaped_kw.replace(r"\ ", r"\s*")
    
    match = re.search(pattern_str, txt)
    if match:
        content_part = txt[match.end():].strip()
        content_part = re.sub(r"^[:\.\-\s]+", "", content_part)
        first_line['text'] = content_part
    else:
        if start_kw in txt:
            parts = txt.split(start_kw, 1)
            first_line['text'] = parts[1].strip() if len(parts) > 1 else ""
        else:
            first_line['text'] = ""
    
    target_lines = []
    if first_line['text'].strip():
        target_lines.append(first_line)
    target_lines.extend(target_lines_raw[1:])
    
    if not target_lines: return ""
    
    # [4. 잔여물 제거 - 정규식 리스트]
    # 주의: 한 글자(의, 료 등)는 단독으로 쓰일 때만 지우거나, 아예 목록에서 뺌
    garbage_regex_list = [
        # 긴 문구들
        r"^(에\s*)?접촉했을\s*때", r"^(에\s*)?들어갔을\s*때", r"^했을\s*때", # [수정] "했을 때" 추가
        r"^흡입했을\s*때", r"^먹었을\s*때", 
        r"^주의사항", r"^내용물", r"^취급요령", r"^저장방법", r"^보호구", r"^조치사항", r"^제거\s*방법",
        r"^소화제", r"^유해성", r"^로부터\s*생기는", r"^착용할\s*보호구", r"^예방조치",
        r"^방법", r"^경고표지\s*항목", r"^그림문자", r"^화학물질", 
        r"^의사의\s*주의사항", r"^기타\s*의사의\s*주의사항", r"^필요한\s*정보", r"^관한\s*정보",
        r"^보호하기\s*위해\s*필요한\s*조치사항", r"^또는\s*제거\s*방법", 
        r"^시\s*착용할\s*보호구(\s*및\s*예방조치)?", 
        r"^부터\s*생기는(\s*특정\s*유해성)?", r"^\(?부적절한\)?\s*소화제",
        
        # 짧은 단어들 (공백 필수 조건 추가하여 안전하게)
        r"^및\s+", r"^요령\s+", r"^때\s+", r"^항의\s+", 
        r"^또는\s+", r"^시\s+"  # [수정] "시" 뒤에 공백 필수 (시작 vs 시 착용)
        # "의"는 삭제 목록에서 완전 제외 (의료인력 보호)
    ]
    
    cleaned_lines = []
    for line in target_lines:
        txt = line['text'].strip()
        
        # 반복 정제
        for _ in range(3):
            changed = False
            for pat in garbage_regex_list:
                match = re.search(pat, txt)
                if match:
                    txt = txt[match.end():].strip()
                    changed = True
            
            # 특수문자 제거
            txt = re.sub(r"^[:\.\)\s]+", "", txt)
            if not changed: break
        
        if txt:
            line['text'] = txt
            cleaned_lines.append(line)
            
    if not cleaned_lines: return ""

    # [5. 문맥 기반 연결 (조사 + 어미)]
    JOSAS = ['을', '를', '이', '가', '은', '는', '의', '와', '과', '에', '로', '서']
    SPACERS_END = ['고', '며', '여', '해', '나', '면', '니', '등', '및', '또는', '경우', ',', ')']
    SPACERS_START = ['및', '또는', '(', '참고']

    final_text = ""
    if len(cleaned_lines) > 0:
        final_text = cleaned_lines[0]['text']
        
        for i in range(1, len(cleaned_lines)):
            prev = cleaned_lines[i-1]
            curr = cleaned_lines[i]
            
            prev_txt = prev['text'].strip()
            curr_txt = curr['text'].strip()
            
            ends_with_sentence = re.search(r"(\.|시오|음|함|것|임|있음|주의|금지|참조|따르시오|마시오)$", prev_txt)
            starts_with_bullet = re.match(r"^(\-|•|\*|\d+\.|[가-하]\.|\(\d+\))", curr_txt)
            
            if ends_with_sentence or starts_with_bullet:
                final_text += "\n" + curr_txt
            else:
                last_char = prev_txt[-1] if prev_txt else ""
                first_char = curr_txt[0] if curr_txt else ""
                
                is_last_hangul = 0xAC00 <= ord(last_char) <= 0xD7A3
                is_first_hangul = 0xAC00 <= ord(first_char) <= 0xD7A3
                
                gap = curr['global_y0'] - prev['global_y1']
                
                if gap < 3.0: 
                    if is_last_hangul and is_first_hangul:
                        need_space = False
                        if last_char in JOSAS: need_space = True
                        elif last_char in SPACERS_END: need_space = True
                        elif any(curr_txt.startswith(x) for x in SPACERS_START): need_space = True
                        
                        if need_space: final_text += " " + curr_txt
                        else: final_text += curr_txt
                    else:
                        final_text += " " + curr_txt
                else:
                    final_text += "\n" + curr_txt
                
    return final_text

# --------------------------------------------------------------------------
# [함수] 메인 파서
# --------------------------------------------------------------------------
def parse_pdf_final(doc):
    all_lines = get_clustered_lines(doc)
    
    result = {
        "hazard_cls": [], "signal_word": "", "h_codes": [], 
        "p_prev": [], "p_resp": [], "p_stor": [], "p_disp": [],
        "composition_data": [], "sec4_to_7": {} 
    }

    limit_y = 999999
    for line in all_lines:
        if "3. 구성성분" in line['text'] or "3. 성분" in line['text']:
            limit_y = line['global_y0']; break
            
    full_text_hp = "\n".join([l['text'] for l in all_lines if l['global_y0'] < limit_y])
    
    for line in full_text_hp.split('\n'):
        if "신호어" in line:
            val = line.replace("신호어", "").replace(":", "").strip()
            if val in ["위험", "경고"]: result["signal_word"] = val
        elif line.strip() in ["위험", "경고"] and not result["signal_word"]:
            result["signal_word"] = line.strip()
    
    lines_hp = full_text_hp.split('\n')
    state = 0
    for l in lines_hp:
        l_ns = l.replace(" ", "")
        if "가.유해성" in l_ns and "분류" in l_ns: state=1; continue
        if "나.예방조치" in l_ns: state=0; continue
        if state==1 and l.strip():
            if "공급자" not in l and "회사명" not in l:
                result["hazard_cls"].append(l.strip())

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
    if not text: return "", 19.2
    
    formatted_text = re.sub(r'(?<!\d)\.(?!\d)(?!\n)', '.\n', text)
    lines = [line.strip() for line in formatted_text.split('\n') if line.strip()]
    final_text = "\n".join(lines)
    
    char_limit_per_line = 45
    
    total_visual_lines = 0
    for line in lines:
        line_len = 0
        for ch in line:
            line_len += 2 if '가' <= ch <= '힣' else 1.1 
        
        visual_lines = math.ceil(line_len / (char_limit_per_line * 2)) 
        if visual_lines == 0: visual_lines = 1
        total_visual_lines += visual_lines
    
    if total_visual_lines == 0: total_visual_lines = 1
    
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
            with st.spinner("최종 정밀 보정 및 변환 중..."):
                
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

                            for row in dest_ws.iter_rows():
                                for cell in row:
                                    if isinstance(cell, MergedCell): continue
                                    if cell.data_type == 'f' and "ingredients" in str(cell.value):
                                        cell.value = ""

                            safe_write_force(dest_ws, 7, 2, product_name_input, center=True)
                            safe_write_force(dest_ws, 10, 2, product_name_input, center=True)
                            
                            if parsed_data["hazard_cls"]:
                                clean_hazard_text = "\n".join([line for line in parsed_data["hazard_cls"] if line.strip()])
                                safe_write_force(dest_ws, 20, 2, clean_hazard_text, center=False)
                                dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

                            signal_final = parsed_data["signal_word"] if parsed_data["signal_word"] else ""
                            safe_write_force(dest_ws, 24, 2, signal_final, center=False) 

                            fill_fixed_range(dest_ws, 25, 36, parsed_data["h_codes"], code_map)
                            fill_fixed_range(dest_ws, 38, 49, parsed_data["p_prev"], code_map)
                            fill_fixed_range(dest_ws, 50, 63, parsed_data["p_resp"], code_map)
                            fill_fixed_range(dest_ws, 64, 69, parsed_data["p_stor"], code_map)
                            fill_fixed_range(dest_ws, 70, 72, parsed_data["p_disp"], code_map)

                            fill_composition_data(dest_ws, parsed_data["composition_data"], cas_name_map)

                            sec_data = parsed_data["sec4_to_7"]
                            import openpyxl.utils
                            
                            for cell_addr, raw_text in sec_data.items():
                                formatted_txt, row_h = format_and_calc_height_sec47(raw_text)
                                try:
                                    col_str = re.match(r"([A-Z]+)", cell_addr).group(1)
                                    row_num = int(re.search(r"(\d+)", cell_addr).group(1))
                                    col_idx = openpyxl.utils.column_index_from_string(col_str)
                                    
                                    safe_write_force(dest_ws, row_num, col_idx, "")
                                    
                                    if formatted_txt:
                                        safe_write_force(dest_ws, row_num, col_idx, formatted_txt, center=False)
                                        dest_ws.row_dimensions[row_num].height = row_h
                                        try:
                                            cell_a = dest_ws.cell(row=row_num, column=1)
                                            if cell_a.value:
                                                cell_a.value = str(cell_a.value).strip()
                                            cell_a.alignment = ALIGN_TITLE
                                        except: pass
                                except Exception as e: pass

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
                    st.success("완료! 글자 잘림 방지 및 정밀 정제 적용.")
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
