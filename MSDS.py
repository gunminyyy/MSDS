import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
import io
import re
import gc
import os
import fitz  # PyMuPDF

# 1. 페이지 설정
st.set_page_config(page_title="MSDS 스마트 변환기", layout="wide")
st.title("MSDS 양식 변환기 (스타일/매핑/레이아웃 완벽 보정)")
st.markdown("---")

# --------------------------------------------------------------------------
# [스타일 정의] 굴림 8pt & 테두리
# --------------------------------------------------------------------------
# 굴림 8pt 폰트
FONT_STYLE = Font(name='굴림', size=8)

# 얇은 실선 테두리
BORDER_STYLE = Border(
    left=Side(style='thin'), 
    right=Side(style='thin'), 
    top=Side(style='thin'), 
    bottom=Side(style='thin')
)

# 정렬 (코드용: 가운데, 내용용: 왼쪽+줄바꿈)
ALIGN_CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
ALIGN_LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)

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
        src_arr = np.array(src_norm, dtype='int16') # numpy 필요
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
# [함수] PDF 텍스트 파싱 (노이즈 제거 & 구역 추출)
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

    ZONE_NONE = 0; ZONE_HAZARD_CLS = 1; ZONE_LABEL_INFO = 2
    current_zone = ZONE_NONE
    
    # 서브존 (P코드)
    SUB_NONE=0; SUB_PREV=1; SUB_RESP=2; SUB_STOR=3; SUB_DISP=4
    current_sub = SUB_NONE
    
    regex_code = re.compile(r"([HP]\d{3}(?:\s*\+\s*[HP]\d{3})*)")
    BLACKLIST_HAZARD = ["공급자정보", "회사명", "주소", "긴급전화번호", "권고용도", "사용상의제한"]

    for line in clean_lines:
        line_ns = line.replace(" ", "")
        
        # 구역 전환
        if "가.유해성" in line_ns and "분류" in line_ns:
            current_zone = ZONE_HAZARD_CLS; continue
        if "나.예방조치" in line_ns:
            current_zone = ZONE_LABEL_INFO; current_sub = SUB_NONE; continue
        if "3.구성성분" in line_ns or "다.기타" in line_ns:
            current_zone = ZONE_NONE; break

        # 데이터 수집
        if current_zone == ZONE_HAZARD_CLS:
            is_bad = False
            for bl in BLACKLIST_HAZARD:
                if bl in line_ns: is_bad = True; break
            if not is_bad:
                result["hazard_cls"].append(line)
                # H코드 보험용 추출
                codes = regex_code.findall(line)
                for c in codes:
                    if c.startswith("H"): result["h_codes"].append(c)

        elif current_zone == ZONE_LABEL_INFO:
            if "신호어" in line_ns:
                val = line.replace("신호어", "").replace(":", "").strip()
                if val: result["signal_word"] = val
            
            # 서브존 전환 (엄격)
            if line_ns.startswith("예방") and len(line_ns) < 15: current_sub = SUB_PREV
            elif line_ns.startswith("대응") and len(line_ns) < 15: current_sub = SUB_RESP
            elif line_ns.startswith("저장") and len(line_ns) < 15: current_sub = SUB_STOR
            elif line_ns.startswith("폐기") and len(line_ns) < 15: current_sub = SUB_DISP

            codes = regex_code.findall(line)
            for c in codes:
                if c.startswith("H"): result["h_codes"].append(c)
                elif c.startswith("P"):
                    if current_sub == SUB_PREV: result["p_prev"].append(c)
                    elif current_sub == SUB_RESP: result["p_resp"].append(c)
                    elif current_sub == SUB_STOR: result["p_stor"].append(c)
                    elif current_sub == SUB_DISP: result["p_disp"].append(c)

    return result

# --------------------------------------------------------------------------
# [함수] 중앙 데이터 매핑 (분할 검색 & 공백 무시)
# --------------------------------------------------------------------------
def get_description_smart(code, code_map):
    # 1. 완전 정규화 (공백 제거, 대문자)
    clean_code = code.replace(" ", "").upper().strip()
    
    # 2. 통째로 검색
    if clean_code in code_map:
        return code_map[clean_code]
    
    # 3. 분할 검색 (P301+P310 -> P301, P310 각각 찾기)
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
# [함수] 셀 스타일 적용 (굴림 8pt, 테두리)
# --------------------------------------------------------------------------
def apply_style(cell, is_code=False):
    cell.font = FONT_STYLE
    cell.border = BORDER_STYLE
    if is_code:
        cell.alignment = ALIGN_CENTER
    else:
        cell.alignment = ALIGN_LEFT

# --------------------------------------------------------------------------
# [함수] 안전 쓰기 (병합 해제 + 스타일 적용)
# --------------------------------------------------------------------------
def safe_write_styled(ws, row, col, value, is_code=False):
    cell = ws.cell(row=row, column=col)
    
    # 병합 해제 시도
    try:
        if isinstance(cell, MergedCell):
            for rng in ws.merged_cells.ranges:
                if cell.coordinate in rng:
                    ws.unmerge_cells(str(rng))
                    break
            cell = ws.cell(row=row, column=col)
    except: pass

    cell.value = value
    apply_style(cell, is_code) # 스타일 적용

# --------------------------------------------------------------------------
# [함수] 스마트 행 쓰기 (행 추가 시 스타일 복구)
# --------------------------------------------------------------------------
def write_ghs_data_final(ws, parsed_data, code_map):
    
    # 1. 앵커 찾기
    anchors = {"H": -1, "PREV": -1, "RESP": -1, "STOR": -1, "DISP": -1}
    for r in range(1, 150):
        val = str(ws.cell(row=r, column=2).value).replace(" ", "")
        if "유해·위험문구" in val: anchors["H"] = r
        elif val == "예방": anchors["PREV"] = r
        elif val == "대응": anchors["RESP"] = r
        elif val == "저장": anchors["STOR"] = r
        elif val == "폐기": anchors["DISP"] = r
    
    # 기본값
    if anchors["H"] == -1: anchors["H"] = 24
    if anchors["PREV"] == -1: anchors["PREV"] = 31
    if anchors["RESP"] == -1: anchors["RESP"] = 41
    if anchors["STOR"] == -1: anchors["STOR"] = 49
    if anchors["DISP"] == -1: anchors["DISP"] = 52

    current_offset = 0
    
    # 섹션 정의
    sections = [
        ("H", parsed_data["h_codes"], "PREV"),
        ("PREV", parsed_data["p_prev"], "RESP"),
        ("RESP", parsed_data["p_resp"], "STOR"),
        ("STOR", parsed_data["p_stor"], "DISP"),
        ("DISP", parsed_data["p_disp"], "END")
    ]
    
    for section_name, codes, next_section_name in sections:
        
        start_row = anchors[section_name] + current_offset + 1
        
        if next_section_name == "END":
            next_header_row = start_row + 1 
        else:
            next_header_row = anchors[next_section_name] + current_offset
            
        available_space = next_header_row - start_row
        
        # 정규화된 유니크 코드 리스트
        unique_codes = []
        for c in codes:
            clean = c.replace(" ", "").upper().strip()
            if clean not in unique_codes: unique_codes.append(clean)
        
        needed_rows = len(unique_codes)
        
        # 행 부족 시 추가
        if needed_rows > available_space:
            rows_to_add = needed_rows - available_space
            ws.insert_rows(next_header_row, amount=rows_to_add)
            current_offset += rows_to_add
            available_space += rows_to_add
        
        # 데이터 쓰기 및 스타일링
        curr = start_row
        for i, code in enumerate(unique_codes):
            ws.row_dimensions[curr].hidden = False
            ws.row_dimensions[curr].height = 19 # 높이 고정
            
            # B열 (코드)
            safe_write_styled(ws, curr, 2, code, is_code=True)
            
            # D열 (내용)
            desc = get_description_smart(code, code_map)
            safe_write_styled(ws, curr, 4, desc, is_code=False)
            
            # [중요] C열, E열 등 주변 셀에도 테두리 그려주기 (끊김 방지)
            for col_idx in range(2, 5): # B, C, D
                cell = ws.cell(row=curr, column=col_idx)
                cell.border = BORDER_STYLE
                cell.font = FONT_STYLE

            curr += 1
            
        # 빈 공간 처리 (숨김 & 내용 삭제)
        limit_row = start_row + available_space
        for r in range(curr, limit_row):
            safe_write_styled(ws, r, 2, "", True)
            safe_write_styled(ws, r, 4, "", False)
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

import numpy as np # 누락 방지

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
            with st.spinner("스타일 보정 및 데이터 매핑 중..."):
                
                new_files = []
                new_download_data = {}
                
                # 중앙 데이터 로드 (CODE, K 컬럼)
                try: 
                    # 엑셀의 '위험 안전문구' 시트 (보통 첫번째)
                    df_master = pd.read_excel(master_data_file, sheet_name=0)
                    code_map = {}
                    # A열(CODE), B열(K) 가정
                    for idx, row in df_master.iterrows():
                        if pd.notna(row.iloc[0]):
                            # Key 공백제거
                            code_key = str(row.iloc[0]).replace(" ", "").upper().strip()
                            desc_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                            code_map[code_key] = desc_val
                except: 
                    df_master = pd.DataFrame()
                    code_map = {}

                for uploaded_file in uploaded_files:
                    if option == "CFF(K)":
                        try:
                            # 1. PDF 파싱
                            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                            parsed_data = parse_pdf_ghs_logic(doc)
                            
                            # 2. 템플릿 준비
                            template_file.seek(0)
                            dest_wb = load_workbook(io.BytesIO(template_file.read()))
                            dest_ws = dest_wb.active

                            # 3. 데이터 동기화
                            target_sheet = '위험 안전문구'
                            if target_sheet in dest_wb.sheetnames: del dest_wb[target_sheet]
                            data_ws = dest_wb.create_sheet(target_sheet)
                            for r in dataframe_to_rows(df_master, index=False, header=True): data_ws.append(r)

                            # 수식 청소
                            for row in dest_ws.iter_rows():
                                for cell in row:
                                    if isinstance(cell, MergedCell): continue
                                    if cell.data_type == 'f' and "ingredients" in str(cell.value):
                                        cell.value = ""

                            # 4. 제품명 & 기본 정보 (스타일 적용)
                            safe_write_styled(dest_ws, 7, 2, product_name_input, is_code=True)
                            safe_write_styled(dest_ws, 10, 2, product_name_input, is_code=True)
                            
                            if parsed_data["hazard_cls"]:
                                b20_text = "\n".join(parsed_data["hazard_cls"])
                                safe_write_styled(dest_ws, 20, 2, b20_text, is_code=False)
                                dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

                            if parsed_data["signal_word"]:
                                safe_write_styled(dest_ws, 24, 2, parsed_data["signal_word"], is_code=True)

                            # 5. [핵심] 스마트 행 쓰기 (굴림 8pt, 테두리, 행추가)
                            write_ghs_data_final(dest_ws, parsed_data, code_map)

                            # 6. 이미지 정렬
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
                    st.success("완료! 스타일(굴림8pt)과 레이아웃이 완벽하게 보정되었습니다.")
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
