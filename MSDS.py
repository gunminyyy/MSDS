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
st.title("MSDS 양식 변환기 (PDF 정밀 파싱 - 좌표 정렬 및 노이즈 차단)")
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
# [함수] PDF 텍스트 정밀 파싱 (좌표 정렬 + 강력 필터링)
# --------------------------------------------------------------------------
def parse_pdf_ghs_logic(doc):
    
    # 1. 노이즈 제거 및 [좌표 순서 정렬]된 줄 리스트 생성
    clean_lines = []
    
    # 무시할 헤더/푸터 노이즈 키워드
    NOISE_KEYWORDS = [
        "물질안전보건자료", "MSDS", "Material Safety Data Sheet",
        "Corea flavors", "주식회사 고려", "HAIR CARE", "Ver.", "발행일", "개정일",
        "제 품 명", "GHS", "페이지", "PAGE", "---"
    ]

    for page in doc:
        # [핵심 수정] sort=True를 사용하여 시각적 순서(좌->우, 위->아래)로 텍스트 정렬
        blocks = page.get_text("blocks", sort=True)
        
        for b in blocks:
            # 블록 내 텍스트 (b[4])를 줄바꿈으로 나눔
            block_text = b[4]
            lines = block_text.split('\n')
            
            for line in lines:
                line_str = line.strip()
                if not line_str: continue
                
                # 노이즈 필터링
                is_noise = False
                for kw in NOISE_KEYWORDS:
                    if kw.replace(" ", "") in line_str.replace(" ", ""):
                        is_noise = True
                        break
                
                if not is_noise:
                    clean_lines.append(line_str)

    # 2. 데이터 저장소
    result = {
        "hazard_cls": [],       # B20
        "signal_word": "",      # B24
        "h_codes": [],          # B25:30
        "p_prev": [],           # B32:41
        "p_resp": [],           # B42:49
        "p_stor": [],           # B50:52
        "p_disp": []            # B53
    }

    # 3. 구역(Zone) 플래그 및 상태 머신
    ZONE_NONE = 0
    ZONE_HAZARD_CLS = 1    # 가. 유해성 분류
    ZONE_LABEL_INFO = 2    # 나. 예방조치
    
    # P코드 하위 구역
    SUBZONE_PREV = 11      # 예방
    SUBZONE_RESP = 12      # 대응
    SUBZONE_STOR = 13      # 저장
    SUBZONE_DISP = 14      # 폐기

    current_zone = ZONE_NONE
    current_subzone = None
    
    # 복합 P코드 정규식: P300, P300+P310, P300 + P310
    regex_code = re.compile(r"([HP]\d{3}(?:\s*\+\s*[HP]\d{3})*)")

    # 유해성 분류 수집 시 절대 들어오면 안 되는 금지어 (섹션 1 내용)
    BLACKLIST_IN_HAZARD = ["공급자정보", "회사명", "주소", "긴급전화번호", "권고용도", "사용상의제한"]

    for line in clean_lines:
        line_ns = line.replace(" ", "") # 공백제거 비교용
        
        # --- [1] 메인 구역 전환 확인 ---
        
        # "가. 유해성·위험성 분류" 감지 -> ZONE_HAZARD_CLS 진입
        if "가.유해성" in line_ns and "분류" in line_ns:
            current_zone = ZONE_HAZARD_CLS
            continue # 제목 줄 스킵

        # "나. 예방조치문구...항목" 감지 -> ZONE_LABEL_INFO 진입 (유해성 분류 종료)
        if "나.예방조치" in line_ns:
            current_zone = ZONE_LABEL_INFO
            current_subzone = None
            continue

        # "3. 구성성분" 감지 -> 종료
        if "3.구성성분" in line_ns or "다.기타" in line_ns:
            current_zone = ZONE_NONE
            break

        # --- [2] 구역별 데이터 수집 ---

        # [ZONE 1] 유해성 분류 내용 (B20)
        if current_zone == ZONE_HAZARD_CLS:
            # [필터] 섹션 1의 내용이 섞여 들어오는지 확인 (공급자 정보 등)
            is_blacklisted = False
            for bl in BLACKLIST_IN_HAZARD:
                if bl in line_ns:
                    is_blacklisted = True
                    break
            
            if not is_blacklisted:
                result["hazard_cls"].append(line)
                # 여기에도 H코드가 섞여 있을 수 있으니 추출
                codes = regex_code.findall(line)
                for c in codes:
                    if c.startswith("H"): result["h_codes"].append(c)

        # [ZONE 2] 라벨 정보 (신호어, H코드, P코드)
        elif current_zone == ZONE_LABEL_INFO:
            
            # 신호어
            if "신호어" in line_ns:
                val = line.replace("신호어", "").replace(":", "").strip()
                if val: result["signal_word"] = val
            
            # P코드 서브존 전환 (엄격: 줄 시작이 키워드일 때만)
            if line_ns.startswith("예방") and len(line_ns) < 10:
                current_subzone = SUBZONE_PREV
            elif line_ns.startswith("대응"):
                current_subzone = SUBZONE_RESP
            elif line_ns.startswith("저장"):
                current_subzone = SUBZONE_STOR
            elif line_ns.startswith("폐기"):
                current_subzone = SUBZONE_DISP

            # 코드 추출
            codes = regex_code.findall(line)
            for c in codes:
                if c.startswith("H"):
                    # H코드는 ZONE_LABEL_INFO 내 어디서든 수집 (유해위험문구 섹션)
                    result["h_codes"].append(c)
                
                elif c.startswith("P"):
                    if current_subzone == SUBZONE_PREV:
                        result["p_prev"].append(c)
                    elif current_subzone == SUBZONE_RESP:
                        result["p_resp"].append(c)
                    elif current_subzone == SUBZONE_STOR:
                        result["p_stor"].append(c)
                    elif current_subzone == SUBZONE_DISP:
                        result["p_disp"].append(c)

    return result

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
            with st.spinner("PDF 정밀 분석 및 데이터 매핑 중..."):
                
                new_files = []
                new_download_data = {}
                
                # 중앙 데이터 로드
                try: 
                    df_master = pd.read_excel(master_data_file, sheet_name=0)
                    code_map = {}
                    for idx, row in df_master.iterrows():
                        if pd.notna(row.iloc[0]):
                            # 공백 제거, 대문자 변환하여 키 생성
                            code_key = str(row.iloc[0]).replace(" ", "").strip().upper()
                            desc_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                            code_map[code_key] = desc_val
                except: 
                    df_master = pd.DataFrame()
                    code_map = {}

                for uploaded_file in uploaded_files:
                    if option == "CFF(K)":
                        try:
                            # 1. PDF 로드 및 파싱 (정렬 모드)
                            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                            parsed_data = parse_pdf_ghs_logic(doc)
                            
                            # 2. 양식 파일 준비
                            template_file.seek(0)
                            dest_wb = load_workbook(io.BytesIO(template_file.read()))
                            dest_ws = dest_wb.active

                            # [데이터 동기화]
                            target_sheet = '위험 안전문구'
                            if target_sheet in dest_wb.sheetnames: del dest_wb[target_sheet]
                            data_ws = dest_wb.create_sheet(target_sheet)
                            for r in dataframe_to_rows(df_master, index=False, header=True): data_ws.append(r)

                            # 수식 청소
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
                            
                            # ---------------------------------------------------
                            # [데이터 입력]
                            # ---------------------------------------------------
                            
                            # [B20] 유해성 분류 (줄바꿈 유지)
                            if parsed_data["hazard_cls"]:
                                b20_text = "\n".join(parsed_data["hazard_cls"])
                                dest_ws['B20'] = b20_text
                                dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

                            # [B24] 신호어
                            if parsed_data["signal_word"]:
                                dest_ws['B24'] = parsed_data["signal_word"]
                                dest_ws['B24'].alignment = Alignment(horizontal='center', vertical='center')

                            # [공통 함수] 코드 입력 및 행 숨김/해제
                            def fill_rows_precise(code_list, start_row, end_row):
                                # 1. 중복 제거 및 Key 정규화
                                unique_codes = []
                                for c in code_list:
                                    clean_c = c.replace(" ", "").strip().upper()
                                    if clean_c not in unique_codes:
                                        unique_codes.append(clean_c)
                                
                                # 2. 해당 범위 행 전체 숨김 해제
                                for r in range(start_row, end_row + 1):
                                    dest_ws.row_dimensions[r].hidden = False
                                
                                # 3. 데이터 입력
                                curr = start_row
                                for code in unique_codes:
                                    if curr > end_row: break
                                    
                                    # B열: 코드 (PDF 원본이 아니라 정규화된 코드 입력 - 깔끔하게)
                                    dest_ws.cell(row=curr, column=2).value = code
                                    
                                    # D열: 중앙 데이터 매핑
                                    matched_desc = code_map.get(code, "") 
                                    dest_ws.cell(row=curr, column=4).value = matched_desc
                                    
                                    curr += 1
                                
                                # 4. 남은 빈 행 다시 숨김
                                for r in range(start_row, end_row + 1):
                                    cell_val = dest_ws.cell(row=r, column=2).value
                                    if cell_val is None or str(cell_val).strip() == "":
                                        dest_ws.row_dimensions[r].hidden = True

                            # [B25~B30] H코드
                            fill_rows_precise(parsed_data["h_codes"], 25, 30)

                            # [B32~B41] 예방
                            fill_rows_precise(parsed_data["p_prev"], 32, 41)

                            # [B42~B49] 대응
                            fill_rows_precise(parsed_data["p_resp"], 42, 49)

                            # [B50~B52] 저장
                            fill_rows_precise(parsed_data["p_stor"], 50, 52)

                            # [B53] 폐기
                            fill_rows_precise(parsed_data["p_disp"], 53, 53)

                            # ---------------------------------------------------
                            # [기존 기능] 이미지 정렬 (유지)
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
                    st.success("완료! PDF 데이터 매핑 오류가 완벽히 수정되었습니다.")
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
