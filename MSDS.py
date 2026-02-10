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
from datetime import datetime

# 1. 페이지 설정
st.set_page_config(page_title="MSDS 스마트 변환기", layout="wide")
st.title("MSDS 양식 변환기 (HP(K) 정렬/구성성분 패치)")
st.markdown("---")

# --------------------------------------------------------------------------
# [스타일]
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
# [핵심] 섹션 추출
# --------------------------------------------------------------------------
def extract_section_smart(all_lines, start_kw, end_kw, mode="CFF(K)"):
    start_idx = -1
    end_idx = -1
    
    clean_start_kw = start_kw.replace(" ", "")
    for i, line in enumerate(all_lines):
        if clean_start_kw in line['text'].replace(" ", ""):
            start_idx = i
            break
    if start_idx == -1: return ""
    
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
    
    if mode == "HP(K)":
        garbage_heads = [
            "에 접촉했을 때", "에 들어갔을 때", "들어갔을 때", "접촉했을 때", "했을 때", 
            "흡입했을 때", "먹었을 때", "주의사항", "내용물", 
            "취급요령", "저장방법", "보호구", "조치사항", "제거 방법",
            "소화제", "유해성", "로부터 생기는", "착용할 보호구", "예방조치",
            "방법", "경고표지 항목", "그림문자", "화학물질", 
            "의사의 주의사항", "기타 의사의 주의사항", "필요한 정보", "관한 정보",
            "보호하기 위해 필요한 조치사항", "또는 제거 방법", 
            "시 착용할 보호구 및 예방조치", "시 착용할 보호구",
            "부터 생기는 특정 유해성", "사의 주의사항", "(부적절한) 소화제",
            "및", "요령", "때", "항의", "색상", "인화점", "비중", "굴절률",
            "에 의한 규제", "의한 규제", "- 색",
            "(및 부적절한) 소화제", "특정 유해성", 
            "보호하기 위해 필요한 조치 사항 및 보호구", "저장 방법"
        ]
    else: 
        garbage_heads = [
            "에 접촉했을 때", "에 들어갔을 때", "들어갔을 때", "접촉했을 때", "했을 때", 
            "흡입했을 때", "먹었을 때", "주의사항", "내용물", 
            "취급요령", "저장방법", "보호구", "조치사항", "제거 방법",
            "소화제", "유해성", "로부터 생기는", "착용할 보호구", "예방조치",
            "방법", "경고표지 항목", "그림문자", "화학물질", 
            "의사의 주의사항", "기타 의사의 주의사항", "필요한 정보", "관한 정보",
            "보호하기 위해 필요한 조치사항", "또는 제거 방법", 
            "시 착용할 보호구 및 예방조치", "시 착용할 보호구",
            "부터 생기는 특정 유해성", "사의 주의사항", "(부적절한) 소화제",
            "및", "요령", "때", "항의", "색상", "인화점", "비중", "굴절률",
            "에 의한 규제", "의한 규제"
        ]
    
    sensitive_garbage_regex = [r"^시\s+", r"^또는\s+", r"^의\s+"]

    cleaned_lines = []
    for line in target_lines:
        txt = line['text'].strip()
        
        if mode == "HP(K)":
            txt = txt.lstrip("-").strip()
        
        for _ in range(3):
            changed = False
            for gb in garbage_heads:
                if txt.replace(" ","").startswith(gb.replace(" ","")):
                     p = re.compile(r"^" + re.escape(gb).replace(r"\ ", r"\s*") + r"[\s\.:]*")
                     m = p.match(txt)
                     if m:
                         txt = txt[m.end():].strip()
                         changed = True
                     elif txt.startswith(gb):
                         txt = txt[len(gb):].strip()
                         changed = True
            
            for pat in sensitive_garbage_regex:
                m = re.search(pat, txt)
                if m:
                    txt = txt[m.end():].strip()
                    changed = True

            txt = re.sub(r"^[:\.\)\s]+", "", txt)
            if not changed: break
        
        if txt:
            if mode == "HP(K)":
                txt = txt.lstrip("-").strip()
            line['text'] = txt
            cleaned_lines.append(line)
            
    if not cleaned_lines: return ""

    JOSAS = ['을', '를', '이', '가', '은', '는', '의', '와', '과', '에', '로', '서']
    SPACERS_END = ['고', '며', '여', '해', '나', '면', '니', '등', '및', '또는', '경우', ',', ')', '속']
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

def parse_sec8_hp_content(text):
    if not text: return "자료없음"
    
    chunks = text.split("-")
    valid_lines = []
    
    for chunk in chunks:
        clean_chunk = chunk.strip()
        if not clean_chunk: continue
        
        if ":" in clean_chunk:
            parts = clean_chunk.split(":", 1)
            name_part = parts[0].strip()
            value_part = parts[1].strip()
            
            if "해당없음" in value_part: continue 
            
            name_part = name_part.replace("[", "").replace("]", "").strip()
            value_part = value_part.replace("[", "").replace("]", "").strip()
            
            final_line = f"{name_part} : {value_part}"
            valid_lines.append(final_line)
        else:
            if "해당없음" not in clean_chunk:
                clean_chunk = clean_chunk.replace("[", "").replace("]", "").strip()
                valid_lines.append(clean_chunk)
            
    if not valid_lines: return "자료없음"
    return "\n".join(valid_lines)

# --------------------------------------------------------------------------
# [함수] 메인 파서 (Dual Mode)
# --------------------------------------------------------------------------
def parse_pdf_final(doc, mode="CFF(K)"):
    all_lines = get_clustered_lines(doc)
    
    if mode == "CFF(K)":
        for i in range(len(all_lines)):
            if "적정선적명" in all_lines[i]['text']:
                target_line = all_lines[i]
                if i > 0:
                    prev_line = all_lines[i-1]
                    if abs(prev_line['global_y0'] - target_line['global_y0']) < 20:
                        if "적정선적명" not in prev_line['text'] and "유엔번호" not in prev_line['text']:
                            all_lines[i]['text'] = target_line['text'] + " " + prev_line['text']
                            all_lines[i-1]['text'] = ""
    
    result = {
        "hazard_cls": [], "signal_word": "", "h_codes": [], 
        "p_prev": [], "p_resp": [], "p_stor": [], "p_disp": [],
        "composition_data": [], "sec4_to_7": {}, "sec8": {}, "sec9": {}, "sec14": {}, "sec15": {}
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
    
    if mode == "HP(K)":
        lines_hp = full_text_hp.split('\n')
        state = 0
        for l in lines_hp:
            if "가. 유해성" in l: state=1; continue
            if "나. 예방조치" in l: state=0; continue
            if state==1 and l.strip():
                if "공급자" not in l and "회사명" not in l:
                    clean_l = l.replace("-", "").strip()
                    if clean_l: result["hazard_cls"].append(clean_l)
    else:
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

    # [수정] CAS 정규식 (공백 허용) + 함유량 조건 완화
    regex_cas = re.compile(r'\b(\d{2,7}\s*-\s*\d{2}\s*-\s*\d)\b')
    regex_conc = re.compile(r'\b(\d+)\s*~\s*(\d+)\b')
    in_comp = False
    for line in all_lines:
        txt = line['text']
        if "3." in txt and ("성분" in txt or "Composition" in txt): in_comp=True; continue
        if "4." in txt and ("응급" in txt or "First" in txt): in_comp=False; break
        if in_comp:
            if re.search(r'^\d+\.\d+', txt): continue 
            cas = regex_cas.search(txt)
            conc = regex_conc.search(txt)
            if cas:
                c_val = cas.group(1).replace(" ", "") 
                # [수정] 함유량 없어도 CAS 있으면 저장
                cn_val = ""
                if conc:
                    s, e = conc.group(1), conc.group(2)
                    if s=="1": s="0"
                    cn_val = f"{s} ~ {e}"
                result["composition_data"].append((c_val, cn_val))

    # 섹션 4~7
    data = {}
    if mode == "HP(K)":
        data["B125"] = extract_section_smart(all_lines, "가. 눈에", "나. 피부", mode)
        data["B126"] = extract_section_smart(all_lines, "나. 피부", "다. 흡입", mode)
        data["B127"] = extract_section_smart(all_lines, "다. 흡입", "라. 먹었을", mode)
        data["B128"] = extract_section_smart(all_lines, "라. 먹었을", "마. 기타", mode)
        data["B129"] = extract_section_smart(all_lines, "마. 기타", ["5.", "폭발"], mode)
        data["B132"] = extract_section_smart(all_lines, "가. 적절한", "나. 화학물질", mode)
        data["B133"] = extract_section_smart(all_lines, "나. 화학물질", "다. 화재진압", mode)
        data["B134"] = extract_section_smart(all_lines, "다. 화재진압", ["6.", "누출"], mode)
    else: 
        data["B125"] = extract_section_smart(all_lines, "나. 눈", "다. 피부", mode)
        data["B126"] = extract_section_smart(all_lines, "다. 피부", "라. 흡입", mode)
        data["B127"] = extract_section_smart(all_lines, "라. 흡입", "마. 먹었을", mode)
        data["B128"] = extract_section_smart(all_lines, "마. 먹었을", "바. 기타", mode)
        data["B129"] = extract_section_smart(all_lines, "바. 기타", ["5.", "폭발"], mode)
        data["B132"] = extract_section_smart(all_lines, "가. 적절한", "나. 화학물질", mode)
        data["B133"] = extract_section_smart(all_lines, "나. 화학물질", "다. 화재진압", mode)
        data["B134"] = extract_section_smart(all_lines, "다. 화재진압", ["6.", "누출"], mode)
    
    data["B138"] = extract_section_smart(all_lines, "가. 인체를", "나. 환경을", mode)
    data["B139"] = extract_section_smart(all_lines, "나. 환경을", "다. 정화", mode)
    data["B140"] = extract_section_smart(all_lines, "다. 정화", ["7.", "취급"], mode)
    data["B143"] = extract_section_smart(all_lines, "가. 안전취급", "나. 안전한", mode)
    data["B144"] = extract_section_smart(all_lines, "나. 안전한", ["8.", "노출"], mode)
    
    result["sec4_to_7"] = data

    sec8_lines = []
    start_8 = -1; end_8 = -1
    for i, line in enumerate(all_lines):
        if "8. 노출방지" in line['text']: start_8 = i
        if "9. 물리화학" in line['text']: end_8 = i; break
    if start_8 != -1:
        if end_8 == -1: end_8 = len(all_lines)
        sec8_lines = all_lines[start_8:end_8]
    
    if mode == "HP(K)":
        b148_raw = extract_section_smart(sec8_lines, "국내노출기준", "ACGIH노출기준", mode)
        b150_raw = extract_section_smart(sec8_lines, "ACGIH노출기준", "생물학적", mode)
        b148_raw = parse_sec8_hp_content(b148_raw)
        b150_raw = parse_sec8_hp_content(b150_raw)
    else:
        b148_raw = extract_section_smart(sec8_lines, "국내규정", "ACGIH", mode)
        b150_raw = extract_section_smart(sec8_lines, "ACGIH", "생물학적", mode)
        
    result["sec8"] = {"B148": b148_raw, "B150": b150_raw}

    sec9_lines = []
    start_9 = -1; end_9 = -1
    for i, line in enumerate(all_lines):
        if "9. 물리화학" in line['text']: start_9 = i
        if "10. 안정성" in line['text']: end_9 = i; break
    if start_9 != -1:
        if end_9 == -1: end_9 = len(all_lines)
        sec9_lines = all_lines[start_9:end_9]
        
    if mode == "HP(K)":
        result["sec9"] = {
            "B163": extract_section_smart(sec9_lines, "- 색", "나. 냄새", mode),
            "B169": extract_section_smart(sec9_lines, "인화점", "아. 증발속도", mode),
            "B176": extract_section_smart(sec9_lines, "비중", "거. n-옥탄올", mode),
            "B182": extract_section_smart(sec9_lines, "굴절률", ["10. 안정성", "10. 화학적"], mode)
        }
    else:
        result["sec9"] = {
            "B163": extract_section_smart(sec9_lines, "색상", "나. 냄새", mode),
            "B169": extract_section_smart(sec9_lines, "인화점", "아. 증발속도", mode),
            "B176": extract_section_smart(sec9_lines, "비중", "거. n-옥탄올", mode),
            "B182": extract_section_smart(sec9_lines, "굴절률", ["10. 안정성", "10. 화학적"], mode)
        }

    sec14_lines = []
    start_14 = -1; end_14 = -1
    for i, line in enumerate(all_lines):
        if "14. 운송에" in line['text']: start_14 = i
        if "15. 법적규제" in line['text']: end_14 = i; break
    if start_14 != -1:
        if end_14 == -1: end_14 = len(all_lines)
        sec14_lines = all_lines[start_14:end_14]
    
    if mode == "HP(K)":
        un_no = extract_section_smart(sec14_lines, "유엔번호", "나. 유엔", mode)
        ship_name = extract_section_smart(sec14_lines, "유엔 적정 선적명", ["다. 운송에서의", "다.운송에서의"], mode)
    else:
        un_no = extract_section_smart(sec14_lines, "유엔번호", "나. 적정선적명", mode)
        ship_name = extract_section_smart(sec14_lines, "적정선적명", ["다. 운송에서의", "다.운송에서의"], mode)
        
    result["sec14"] = {"UN": un_no, "NAME": ship_name}

    sec15_lines = []
    start_15 = -1; end_15 = -1
    for i, line in enumerate(all_lines):
        if "15. 법적규제" in line['text']: start_15 = i
        if "16. 그 밖의" in line['text']: end_15 = i; break
    if start_15 != -1:
        if end_15 == -1: end_15 = len(all_lines)
        sec15_lines = all_lines[start_15:end_15]
    
    danger_act = extract_section_smart(sec15_lines, "위험물안전관리법", "마. 폐기물", mode)
    result["sec15"] = {"DANGER": danger_act}

    return result

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
    else: cell.alignment = ALIGN_LEFT

def calculate_smart_height_basic(text): 
    if not text: return 19.2
    explicit_lines = str(text).count('\n') + 1
    final_lines = max(explicit_lines, 1)
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
        # [수정] 함유량 없어도 표시 (comp_data[i][1] 조건 제거)
        if i < len(comp_data):
            cas_no, concentration = comp_data[i]
            clean_cas = cas_no.replace(" ", "").strip()
            chem_name = cas_to_name_map.get(clean_cas, "")
            ws.row_dimensions[current_row].hidden = False
            ws.row_dimensions[current_row].height = 26.
