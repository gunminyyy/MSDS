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
st.title("MSDS 양식 변환기 (PDF 지원 정밀 모드)")
st.markdown("---")

# --------------------------------------------------------------------------
# [함수] 이미지 정규화
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

# [함수] 리소스 경로 찾기
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

# [함수] 이미지 비교 매칭
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

# [함수] 파일명에서 숫자 추출
def extract_number(filename):
    nums = re.findall(r'\d+', filename)
    return int(nums[0]) if nums else 999

# 2. 파일 업로드
with st.expander("📂 필수 파일 업로드", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        master_data_file = st.file_uploader("1. 중앙 데이터 (master_data.xlsx)", type="xlsx")
        loaded_refs, folder_exists = get_reference_images()
        if folder_exists and loaded_refs:
            st.success(f"✅ 기준 그림 {len(loaded_refs)}개 로드됨 (폴더: reference_imgs)")
        elif not folder_exists:
            st.warning("⚠️ 'reference_imgs' 폴더를 만들고 그림 파일들을 넣어주세요.")

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
            with st.spinner("PDF 분석 및 데이터 변환 중..."):
                
                new_files = []
                new_download_data = {}
                
                # 중앙 데이터 읽기 (문자열 변환 강화)
                try: 
                    df_master = pd.read_excel(master_data_file, sheet_name=0)
                    code_map = {}
                    for idx, row in df_master.iterrows():
                        # [수정] 공백 제거 및 문자열 강제 변환
                        code_val = str(row.iloc[0]).strip()
                        desc_val = str(row.iloc[1]).strip()
                        code_map[code_val] = desc_val
                except: 
                    df_master = pd.DataFrame()
                    code_map = {}

                for uploaded_file in uploaded_files:
                    if option == "CFF(K)":
                        try:
                            # 1. PDF 로드
                            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                            
                            # 2. 양식 파일 준비
                            template_file.seek(0)
                            dest_wb = load_workbook(io.BytesIO(template_file.read()))
                            dest_ws = dest_wb.active

                            # ---------------------------------------------------
                            # [데이터 동기화 & 수식 수정]
                            # ---------------------------------------------------
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

                            # 제품명 입력
                            dest_ws['B7'] = product_name_input
                            dest_ws['B10'] = product_name_input
                            
                            # ---------------------------------------------------
                            # [PDF 텍스트 분석 및 추출]
                            # ---------------------------------------------------
                            full_text = ""
                            # 줄바꿈 유지를 위해 "text" 옵션 사용
                            for page in doc:
                                full_text += page.get_text("text") + "\n"

                            # [A] 유해성 본문 (B20) - 헤더 제외 및 줄바꿈 유지 수정
                            # "가. 유해성...분류" 헤더 다음 내용부터 "나. 예방..." 전까지
                            # 정규식: 헤더(group1) + 내용(group2) + 다음헤더
                            pattern_b20 = re.search(r"(가\.\s*유해성.*?분류\s*\n)(.*?)(나\.\s*예방조치)", full_text, re.DOTALL)
                            
                            b20_text = ""
                            if pattern_b20:
                                # group(2)가 실제 내용입니다. strip()으로 앞뒤 공백만 제거
                                raw_content = pattern_b20.group(2).strip()
                                b20_text = raw_content[:1000] # 길이 제한
                            
                            if b20_text:
                                dest_ws['B20'] = b20_text
                                dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

                            # [B] H코드 추출 (B25 ~ B30)
                            # B20에서 추출한 유해성 본문 내에서 H코드 검색 (정확도 향상)
                            extracted_h_codes = []
                            if b20_text:
                                found_h_codes = re.findall(r"H\d{3}", b20_text)
                                for code in found_h_codes:
                                    if code not in extracted_h_codes: extracted_h_codes.append(code)
                            
                            # B25 입력 및 D열 매칭
                            current_target_row = 25
                            for code in extracted_h_codes:
                                if current_target_row > 30: break
                                # 코드 입력 (문자열로 변환하여 공백제거)
                                clean_code = str(code).strip()
                                dest_ws.cell(row=current_target_row, column=2).value = clean_code
                                
                                # 매칭 확인
                                matched_desc = code_map.get(clean_code, "")
                                dest_ws.cell(row=current_target_row, column=4).value = matched_desc
                                
                                current_target_row += 1
                            
                            # B25~B30 숨김 처리
                            for r in range(25, 31):
                                if not dest_ws.cell(row=r, column=2).value:
                                    dest_ws.row_dimensions[r].hidden = True
                                else:
                                    dest_ws.row_dimensions[r].hidden = False

                            # ---------------------------------------------------
                            # [신규] P코드 섹션별 정밀 추출 (순서 뒤섞임 방지)
                            # ---------------------------------------------------
                            
                            # 전체 텍스트에서 "나. 예방조치...항목" 부터 "3. 구성성분" 전까지 추출
                            section_2_block_match = re.search(r"나\.\s*예방조치.*?항목\s*\n(.*?)(3\.\s*구성성분|다\.\s*기타)", full_text, re.DOTALL)
                            section_2_text = section_2_block_match.group(1) if section_2_block_match else ""

                            # 섹션별 텍스트 나누기 (예방 -> 대응 -> 저장 -> 폐기 순서 보장)
                            # 정규식으로 각 키워드의 위치(인덱스)를 찾습니다.
                            # 주의: PDF 줄바꿈으로 인해 "예 방", "대 응" 등으로 띄어쓰기가 있을 수 있음
                            
                            # 1. 예방 ~ 대응 사이
                            match_prev = re.search(r"(예\s*방)(.*?)(대\s*응)", section_2_text, re.DOTALL)
                            txt_prevention = match_prev.group(2) if match_prev else ""
                            
                            # 2. 대응 ~ 저장 사이
                            match_resp = re.search(r"(대\s*응)(.*?)(저\s*장)", section_2_text, re.DOTALL)
                            txt_response = match_resp.group(2) if match_resp else ""
                            
                            # 3. 저장 ~ 폐기 사이
                            match_stor = re.search(r"(저\s*장)(.*?)(폐\s*기)", section_2_text, re.DOTALL)
                            txt_storage = match_stor.group(2) if match_stor else ""
                            
                            # 4. 폐기 ~ 끝까지
                            match_disp = re.search(r"(폐\s*기)(.*)", section_2_text, re.DOTALL)
                            txt_disposal = match_disp.group(2) if match_disp else ""

                            # 공통 함수: P코드 추출 및 셀 입력 (D열 매칭 포함)
                            def fill_p_codes(target_text, start_row, end_row):
                                # P코드 정규식 (P300+P310 형태 포함)
                                p_codes = re.findall(r"P\d{3}(?:\+P\d{3})*", target_text)
                                unique_p = []
                                for p in p_codes:
                                    if p not in unique_p: unique_p.append(p)
                                
                                # 우선 해당 범위 숨김 취소
                                for r in range(start_row, end_row + 1):
                                    dest_ws.row_dimensions[r].hidden = False
                                
                                curr = start_row
                                for p_code in unique_p:
                                    if curr > end_row: break
                                    
                                    clean_p = str(p_code).strip()
                                    dest_ws.cell(row=curr, column=2).value = clean_p
                                    
                                    # D열 매칭 (중앙 데이터)
                                    # P코드는 +로 연결된 경우가 있으므로, 없으면 각각 찾아서 합치거나 그대로 둠
                                    if clean_p in code_map:
                                        dest_ws.cell(row=curr, column=4).value = code_map[clean_p]
                                    else:
                                        # 매칭 실패 시 (복합 코드 등) -> 일단 빈칸 (또는 수동 확인 필요)
                                        # 복합코드(P300+P310)인 경우 중앙 데이터에 해당 키가 없으면 안 나옵니다.
                                        # 중앙 데이터에 "P300+P310" 키가 있거나, 아니면 코드를 쪼개서 찾아야 합니다.
                                        # 여기서는 일단 1:1 매칭 시도
                                        dest_ws.cell(row=curr, column=4).value = code_map.get(clean_p, "")

                                    curr += 1
                                
                                # 값이 안 들어간 나머지 행 숨기기
                                for r in range(start_row, end_row + 1):
                                    if not dest_ws.cell(row=r, column=2).value:
                                        dest_ws.row_dimensions[r].hidden = True

                            # 각 섹션별 적용
                            fill_p_codes(txt_prevention, 32, 41)
                            fill_p_codes(txt_response, 42, 49)
                            fill_p_codes(txt_storage, 50, 52)
                            fill_p_codes(txt_disposal, 53, 53)

                            # ---------------------------------------------------
                            # [기존 기능] PDF 이미지 추출 및 정렬
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
                                if key not in unique_images:
                                    unique_images[key] = img
                            
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

                            # 저장
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
