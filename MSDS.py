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
st.title("MSDS 양식 변환기 (PDF 정밀 파싱)")
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
# [신규 함수] PDF 섹션 파싱 (한 줄씩 읽기)
# --------------------------------------------------------------------------
def parse_pdf_ghs_section(doc):
    """
    PDF를 줄 단위로 읽어서 유해성 분류(B20)와 P코드 섹션(예방,대응,저장,폐기)을 분리함
    """
    full_text_lines = []
    for page in doc:
        # 줄 단위 리스트로 가져오기
        text = page.get_text("text")
        lines = text.split('\n')
        full_text_lines.extend(lines)

    # 데이터 저장소
    extracted_data = {
        "hazard_classification": [], # B20 내용
        "prevention": [],
        "response": [],
        "storage": [],
        "disposal": [],
        "h_codes": [] # 전체 H코드
    }

    # 상태 플래그
    mode = None  # None -> 'hazard_cls' -> 'label_elements'
    p_section = None # 'prevention', 'response', 'storage', 'disposal'
    
    # 키워드 정리 (공백 제거 후 비교용)
    KEY_HAZARD_START = "유해성·위험성분류" # 가. 유해성...
    KEY_LABEL_START = "예방조치문구" # 나. ... 항목 (또는 경고표지)
    KEY_COMP_START = "3.구성성분" # 다음 챕터
    
    # P코드 섹션 키워드
    KEY_PREV = "예방"
    KEY_RESP = "대응"
    KEY_STOR = "저장"
    KEY_DISP = "폐기"

    for line in full_text_lines:
        clean_line = line.strip()
        if not clean_line: continue
        
        line_nospace = clean_line.replace(" ", "")
        
        # 1. 유해성 분류 (B20) 시작 감지
        if KEY_HAZARD_START in line_nospace and "가." in line_nospace:
            mode = 'hazard_cls'
            continue # 제목 줄은 포함 안 함
        
        # 2. 예방조치문구 (경고표지 항목) 시작 감지 -> B20 종료
        if KEY_LABEL_START in line_nospace:
            mode = 'label_elements'
            p_section = None # 아직 소제목 안 나옴
            continue
        
        # 3. 섹션 3 시작 -> 종료
        if KEY_COMP_START in line_nospace:
            break

        # --- 모드별 동작 ---
        
        # [A] 유해성 분류 내용 수집
        if mode == 'hazard_cls':
            # 내용에 H코드 등이 섞여 있을 수 있음
            extracted_data["hazard_classification"].append(clean_line)
            # 여기서 H코드 추출
            h_found = re.findall(r"H\d{3}", clean_line)
            extracted_data["h_codes"].extend(h_found)

        # [B] 예방조치문구 내용 수집 (P코드 섹션 감지)
        elif mode == 'label_elements':
            # 소제목 감지 (줄의 시작 부분이 키워드일 때)
            # 주의: "예방조치"라는 단어가 문장에 들어갈 수도 있으므로, 짧은 키워드 매칭 시 주의
            
            # 섹션 전환 로직 (우선순위: 폐기 > 저장 > 대응 > 예방)
            if clean_line.startswith(KEY_DISP):
                p_section = 'disposal'
            elif clean_line.startswith(KEY_STOR):
                p_section = 'storage'
            elif clean_line.startswith(KEY_RESP):
                p_section = 'response'
            elif clean_line.startswith(KEY_PREV):
                p_section = 'prevention'
            
            # 현재 섹션에 내용 담기 (제목 줄 포함 여부는 내용에 따라 다르나, 코드는 보통 제목 줄에 없음)
            if p_section:
                # P코드 추출 (P300+P310 같은 복합 코드 지원)
                # 정규식: P숫자3개 + (플러스 + P숫자3개)가 0번 이상 반복
                p_codes = re.findall(r"P\d{3}(?:\s*\+\s*P\d{3})*", clean_line)
                if p_codes:
                    extracted_data[p_section].extend(p_codes)

    return extracted_data

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
            with st.spinner("PDF 정밀 분석 중..."):
                
                new_files = []
                new_download_data = {}
                
                # 중앙 데이터 로드 (매핑용)
                try: 
                    df_master = pd.read_excel(master_data_file, sheet_name=0)
                    code_map = {}
                    for idx, row in df_master.iterrows():
                        # 공백 제거 및 문자열 변환
                        code_val = str(row.iloc[0]).replace(" ", "").strip()
                        desc_val = str(row.iloc[1]).strip()
                        code_map[code_val] = desc_val
                except: 
                    df_master = pd.DataFrame()
                    code_map = {}

                for uploaded_file in uploaded_files:
                    if option == "CFF(K)":
                        try:
                            # 1. PDF 로드 및 파싱
                            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                            parsed_data = parse_pdf_ghs_section(doc)
                            
                            # 2. 양식 파일 준비
                            template_file.seek(0)
                            dest_wb = load_workbook(io.BytesIO(template_file.read()))
                            dest_ws = dest_wb.active

                            # [데이터 동기화 & 수식 수정]
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
                            
                            # ---------------------------------------------------
                            # [데이터 입력] 파싱된 데이터 넣기
                            # ---------------------------------------------------
                            
                            # [A] 유해성 분류 (B20)
                            # 리스트를 줄바꿈 문자로 합침
                            b20_text = "\n".join(parsed_data["hazard_classification"])
                            dest_ws['B20'] = b20_text
                            dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

                            # [B] H코드 (B25 ~ B30)
                            # 중복 제거 및 순서 유지
                            unique_h = sorted(list(set(parsed_data["h_codes"])))
                            
                            curr = 25
                            for code in unique_h:
                                if curr > 30: break
                                clean_code = code.replace(" ", "").strip()
                                dest_ws.cell(row=curr, column=2).value = clean_code
                                dest_ws.cell(row=curr, column=4).value = code_map.get(clean_code, "")
                                curr += 1
                            
                            # 빈 행 숨기기
                            for r in range(25, 31):
                                if not dest_ws.cell(row=r, column=2).value:
                                    dest_ws.row_dimensions[r].hidden = True
                                else:
                                    dest_ws.row_dimensions[r].hidden = False

                            # [C] P코드 입력 함수
                            def fill_section_codes(p_code_list, start_row, end_row):
                                # 중복 제거
                                unique_p = []
                                for p in p_code_list:
                                    # 공백 정규화 (P300 + P310 -> P300+P310)
                                    norm_p = p.replace(" ", "")
                                    if norm_p not in unique_p: unique_p.append(norm_p)
                                
                                # 숨김 취소
                                for r in range(start_row, end_row + 1):
                                    dest_ws.row_dimensions[r].hidden = False
                                
                                curr = start_row
                                for p_code in unique_p:
                                    if curr > end_row: break
                                    dest_ws.cell(row=curr, column=2).value = p_code
                                    dest_ws.cell(row=curr, column=4).value = code_map.get(p_code, "") # 매칭
                                    curr += 1
                                
                                # 빈 행 숨기기
                                for r in range(start_row, end_row + 1):
                                    if not dest_ws.cell(row=r, column=2).value:
                                        dest_ws.row_dimensions[r].hidden = True

                            # 섹션별 적용
                            fill_section_codes(parsed_data["prevention"], 32, 41)
                            fill_section_codes(parsed_data["response"], 42, 49)
                            fill_section_codes(parsed_data["storage"], 50, 52)
                            fill_section_codes(parsed_data["disposal"], 53, 53)

                            # ---------------------------------------------------
                            # [기존 기능] 이미지 정렬 (로직 유지)
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
                    st.success("완료! PDF 분석 및 변환이 끝났습니다.")
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
