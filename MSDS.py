import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
import io
import re
import gc
import fitz  # PyMuPDF
import os
import numpy as np

# 1. 페이지 설정
st.set_page_config(page_title="MSDS 스마트 변환기", layout="wide")
st.title("MSDS 양식 변환기 (위치 밀림 자동보정 & 데이터 매핑 강화)")
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
# [함수] PDF 파싱 (구역별 정밀 추출)
# --------------------------------------------------------------------------
def parse_pdf_ghs_final(doc):
    # 1. 노이즈 제거된 텍스트 라인 추출
    clean_lines = []
    NOISE_KEYWORDS = [
        "물질안전보건자료", "MSDS", "Material Safety Data Sheet",
        "Corea flavors", "주식회사 고려", "HAIR CARE", "Ver.", "발행일", "개정일",
        "제 품 명", "GHS", "페이지", "PAGE", "---"
    ]

    for page in doc:
        # sort=True로 시각적 순서 정렬
        blocks = page.get_text("blocks", sort=True)
        for b in blocks:
            text = b[4]
            lines = text.split('\n')
            for line in lines:
                line_str = line.strip()
                if not line_str: continue
                # 노이즈 필터링
                is_noise = False
                for kw in NOISE_KEYWORDS:
                    if kw.replace(" ", "") in line_str.replace(" ", ""):
                        is_noise = True; break
                if not is_noise: clean_lines.append(line_str)

    result = {
        "hazard_cls": [], "signal_word": "", "h_codes": [],
        "p_prev": [], "p_resp": [], "p_stor": [], "p_disp": []
    }

    # 2. 구역(Zone) 상태 머신
    ZONE_NONE = 0
    ZONE_HAZARD_CLS = 1    # B20
    ZONE_LABEL_INFO = 2    # 라벨 정보 구간
    
    current_zone = ZONE_NONE
    
    # 서브존 (P코드)
    SUB_NONE = 0
    SUB_PREV = 1; SUB_RESP = 2; SUB_STOR = 3; SUB_DISP = 4

    current_sub = SUB_NONE
    
    regex_code = re.compile(r"([HP]\d{3}(?:\s*\+\s*[HP]\d{3})*)")
    
    # B20 수집 시 제외할 단어
    BLACKLIST_HAZARD = ["공급자정보", "회사명", "주소", "긴급전화번호", "권고용도", "사용상의제한"]

    for line in clean_lines:
        line_ns = line.replace(" ", "")
        
        # [메인 구역 전환]
        if "가.유해성" in line_ns and "분류" in line_ns:
            current_zone = ZONE_HAZARD_CLS; continue
        if "나.예방조치" in line_ns:
            current_zone = ZONE_LABEL_INFO; current_sub = SUB_NONE; continue
        if "3.구성성분" in line_ns or "다.기타" in line_ns:
            current_zone = ZONE_NONE; break

        # [데이터 수집]
        if current_zone == ZONE_HAZARD_CLS:
            # 1번 섹션 내용 혼입 방지
            is_bad = False
            for bl in BLACKLIST_HAZARD:
                if bl in line_ns: is_bad = True; break
            if not is_bad:
                result["hazard_cls"].append(line)
                # 혹시 모를 H코드
                codes = regex_code.findall(line)
                for c in codes:
                    if c.startswith("H"): result["h_codes"].append(c)

        elif current_zone == ZONE_LABEL_INFO:
            if "신호어" in line_ns:
                val = line.replace("신호어", "").replace(":", "").strip()
                if val: result["signal_word"] = val
            
            # 서브존 전환 (줄 시작 단어로 엄격 구분)
            # 글자수 제한: "화재 예방을 위해" 같은 문장 방지
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
# [함수] 중앙 데이터 매핑 (분할 검색 기능 추가)
# --------------------------------------------------------------------------
def get_description(code, code_map):
    """
    코드를 받아서 설명을 반환. 
    1. 원본 그대로 검색
    2. 공백 제거 후 검색
    3. '+'로 쪼개서 각각 검색 후 합침 (복합 코드 대응)
    """
    # 1. 기본 정규화
    clean_code = code.replace(" ", "").upper().strip()
    
    # 맵핑 시도 1: 통째로 찾기
    if clean_code in code_map:
        return code_map[clean_code]
    
    # 맵핑 시도 2: +로 연결된 경우 쪼개서 찾기 (P301+P310 -> P301 내용 + P310 내용)
    if "+" in clean_code:
        parts = clean_code.split("+")
        descriptions = []
        for part in parts:
            desc = code_map.get(part, "") # 없으면 빈칸
            if desc: descriptions.append(desc)
        
        if descriptions:
            return " ".join(descriptions) # 찾은 내용들을 이어 붙임
            
    return "" # 정말 없으면 빈 문자열

# --------------------------------------------------------------------------
# [함수] 안전 쓰기 (병합 해제)
# --------------------------------------------------------------------------
def safe_write_force(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    try:
        # 병합된 셀이면 해제
        if isinstance(cell, MergedCell):
            for rng in ws.merged_cells.ranges:
                if cell.coordinate in rng:
                    ws.unmerge_cells(str(rng))
                    break
            cell = ws.cell(row=row, column=col) # 다시 조회
        cell.value = value
    except:
        pass

# --------------------------------------------------------------------------
# [함수] 스마트 행 관리 및 쓰기 (위치 밀림 보정)
# --------------------------------------------------------------------------
def write_ghs_data_smart(ws, parsed_data, code_map):
    
    # 1. 고정 앵커(Anchors) 위치 찾기 (템플릿 기준)
    # 템플릿의 초기 위치를 찾아둡니다.
    anchors = {
        "H": -1, "PREV": -1, "RESP": -1, "STOR": -1, "DISP": -1
    }
    
    # 전체 스캔하여 헤더 위치 파악
    for r in range(1, 150):
        val = str(ws.cell(row=r, column=2).value).replace(" ", "")
        if "유해·위험문구" in val: anchors["H"] = r
        elif val == "예방": anchors["PREV"] = r
        elif val == "대응": anchors["RESP"] = r
        elif val == "저장": anchors["STOR"] = r
        elif val == "폐기": anchors["DISP"] = r
    
    # 혹시 못 찾았을 경우를 대비한 기본값 (템플릿 구조 가정)
    if anchors["H"] == -1: anchors["H"] = 24  # 예: 24행 헤더 -> 25행부터 데이터
    if anchors["PREV"] == -1: anchors["PREV"] = 31
    if anchors["RESP"] == -1: anchors["RESP"] = 41
    if anchors["STOR"] == -1: anchors["STOR"] = 49
    if anchors["DISP"] == -1: anchors["DISP"] = 52

    # 2. 섹션별 처리 함수 (Offset 관리)
    # current_offset: 행이 추가됨에 따라 아래쪽 섹션들이 얼마나 밀려야 하는지 추적
    current_offset = 0
    
    # 처리 순서: H -> 예방 -> 대응 -> 저장 -> 폐기
    sections = [
        ("H", parsed_data["h_codes"], "PREV"),
        ("PREV", parsed_data["p_prev"], "RESP"),
        ("RESP", parsed_data["p_resp"], "STOR"),
        ("STOR", parsed_data["p_stor"], "DISP"),
        ("DISP", parsed_data["p_disp"], "END")
    ]
    
    for section_name, codes, next_section_name in sections:
        
        # 현재 섹션의 시작 행 (원래 위치 + 지금까지 밀린 offset)
        start_row = anchors[section_name] + current_offset + 1
        
        # 다음 섹션의 헤더 위치 (범위 계산용)
        if next_section_name == "END":
            # 폐기의 경우 다음 섹션이 없으므로 적당히 1행으로 간주하거나 현재 남은 칸
            next_header_row = start_row + 1 # 최소 1칸
        else:
            next_header_row = anchors[next_section_name] + current_offset
            
        available_space = next_header_row - start_row
        
        # 중복 제거 및 정규화
        unique_codes = []
        for c in codes:
            clean = c.replace(" ", "").upper().strip()
            if clean not in unique_codes: unique_codes.append(clean) # 여기선 원본이 아니라 정규화된 것 저장
        
        needed_rows = len(unique_codes)
        
        # 행 부족 시 삽입
        if needed_rows > available_space:
            rows_to_add = needed_rows - available_space
            # 다음 헤더 위치 직전에 삽입하여 공간 확보
            ws.insert_rows(next_header_row, amount=rows_to_add)
            current_offset += rows_to_add # 오프셋 누적
            available_space += rows_to_add # 가용 공간 늘어남
        
        # 데이터 쓰기
        curr = start_row
        for i, code in enumerate(unique_codes):
            # 행 속성 설정
            ws.row_dimensions[curr].hidden = False
            ws.row_dimensions[curr].height = 19
            
            # 코드 입력
            safe_write_force(ws, curr, 2, code)
            
            # 설명 매핑 (핵심: 여기서 분할 검색 사용)
            desc = get_description(code, code_map)
            safe_write_force(ws, curr, 4, desc)
            
            curr += 1
            
        # 남은 빈 공간 처리 (숨김 & 내용 삭제)
        # 데이터를 쓴 곳(curr)부터 다음 헤더(start_row + available_space)까지
        limit_row = start_row + available_space
        for r in range(curr, limit_row):
            safe_write_force(ws, r, 2, "")
            safe_write_force(ws, r, 4, "")
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
            with st.spinner("데이터 분석 및 매핑 중..."):
                
                new_files = []
                new_download_data = {}
                
                # 중앙 데이터 로드 (정규화 필수)
                try: 
                    df_master = pd.read_excel(master_data_file, sheet_name=0)
                    code_map = {}
                    for idx, row in df_master.iterrows():
                        if pd.notna(row.iloc[0]):
                            # [핵심] 키 정규화 (공백제거, 대문자)
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
                            parsed_data = parse_pdf_ghs_final(doc)
                            
                            # 2. 템플릿 준비
                            template_file.seek(0)
                            dest_wb = load_workbook(io.BytesIO(template_file.read()))
                            dest_ws = dest_wb.active

                            # 3. 데이터 동기화 및 수식 초기화
                            target_sheet = '위험 안전문구'
                            if target_sheet in dest_wb.sheetnames: del dest_wb[target_sheet]
                            data_ws = dest_wb.create_sheet(target_sheet)
                            for r in dataframe_to_rows(df_master, index=False, header=True): data_ws.append(r)

                            # 수식 청소 (병합 셀 건너뛰기)
                            for row in dest_ws.iter_rows():
                                for cell in row:
                                    if isinstance(cell, MergedCell): continue
                                    if cell.data_type == 'f':
                                        f_str = str(cell.value)
                                        if "ingredients" in f_str:
                                            cell.value = "" # 외부 참조 수식 제거

                            # 4. 단순 데이터 입력
                            safe_write_force(dest_ws, 7, 2, product_name_input)
                            safe_write_force(dest_ws, 10, 2, product_name_input)
                            
                            if parsed_data["hazard_cls"]:
                                b20_text = "\n".join(parsed_data["hazard_cls"])
                                safe_write_force(dest_ws, 20, 2, b20_text)
                                dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

                            if parsed_data["signal_word"]:
                                safe_write_force(dest_ws, 24, 2, parsed_data["signal_word"])
                                dest_ws['B24'].alignment = Alignment(horizontal='center', vertical='center')

                            # 5. [핵심] 스마트 행 쓰기 (위치 보정 + 데이터 매핑)
                            write_ghs_data_smart(dest_ws, parsed_data, code_map)

                            # 6. 이미지 처리
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
                    st.success("완료! 중앙 데이터 매핑 및 행 밀림 현상이 완벽하게 해결되었습니다.")
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
