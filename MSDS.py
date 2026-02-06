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
st.title("MSDS 양식 변환기 (PDF 정밀 파싱 - 최종 완벽 수정)")
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
# [핵심 함수] PDF 텍스트 정밀 파싱 (State Machine)
# --------------------------------------------------------------------------
def parse_pdf_state_machine(doc):
    # 1. 노이즈 제거 및 줄 단위 리스트 생성
    clean_lines = []
    
    # PDF 헤더/푸터 등 무시할 키워드 (공백 제거 후 비교용)
    NOISE_KEYWORDS = [
        "물질안전보건자료", "MSDS", "MaterialSafetyDataSheet",
        "Coreaflavors", "주식회사고려", "HAIRCARE", "Ver.", "발행일", "개정일",
        "제품명:", "GHS", "페이지", "PAGE"
    ]

    for page in doc:
        text = page.get_text("text")
        lines = text.split('\n')
        for line in lines:
            line_str = line.strip()
            if not line_str: continue
            
            # 노이즈 필터링
            line_check = line_str.replace(" ", "")
            is_noise = False
            for kw in NOISE_KEYWORDS:
                if kw.replace(" ", "") in line_check:
                    is_noise = True
                    break
            
            # 특정 상황 제외: "가. 제품명" 같은 항목은 노이즈 키워드가 있어도 살려야 할 수 있음
            # 하지만 여기선 2장(유해성) 데이터를 뽑는 게 목적이므로 과감히 날려도 됨.
            if not is_noise:
                clean_lines.append(line_str)

    # 2. 상태 머신을 이용한 데이터 추출
    result = {
        "hazard_cls": [],
        "signal_word": "",
        "h_codes": [],
        "p_prev": [],
        "p_resp": [],
        "p_stor": [],
        "p_disp": []
    }

    # 상태 정의
    STATE_INIT = 0
    STATE_HAZARD_CLS = 1    # 가. 유해성 분류
    STATE_LABEL_START = 2   # 나. 예방조치... (대기)
    STATE_PREV = 3          # 예방
    STATE_RESP = 4          # 대응
    STATE_STOR = 5          # 저장
    STATE_DISP = 6          # 폐기
    STATE_END = 99

    current_state = STATE_INIT
    
    # 정규식
    # P코드: P300, P300+P310 (공백 허용)
    regex_code = re.compile(r"([HP]\d{3}(?:\s*\+\s*[HP]\d{3})*)")

    for line in clean_lines:
        line_ns = line.replace(" ", "") # 공백 제거 문자열
        
        # [상태 전환 로직]
        
        # 1. 유해성 분류 시작
        if "가.유해성" in line_ns and "분류" in line_ns:
            current_state = STATE_HAZARD_CLS
            continue # 제목 줄 스킵

        # 2. 경고표지 항목 (유해성 분류 끝)
        if "나.예방조치" in line_ns:
            current_state = STATE_LABEL_START
            continue
        
        # 3. 구성성분 (모든 추출 종료)
        if "3.구성성분" in line_ns or "다.기타" in line_ns:
            current_state = STATE_END
            break

        # [데이터 수집 로직]
        
        # [B20] 유해성 분류 내용
        if current_state == STATE_HAZARD_CLS:
            result["hazard_cls"].append(line)
            # 분류 내용 안에서도 H코드가 있을 수 있음 (추출)
            found = regex_code.findall(line)
            for c in found:
                if c.startswith("H"): result["h_codes"].append(c)

        # [신호어, H코드, P코드]
        elif current_state >= STATE_LABEL_START:
            
            # 신호어 찾기 (어디서든)
            if "신호어" in line_ns:
                val = line.replace("신호어", "").replace(":", "").strip()
                if val: result["signal_word"] = val
            
            # 유해위험문구 (H코드)
            if "유해" in line_ns and "위험문구" in line_ns:
                # 이 줄부터 H코드 찾기 시작 (상태 변경은 안함, 문맥상 찾음)
                pass
            
            # P코드 섹션 감지 (순서대로 상태 변경)
            if line_ns.startswith("예방"):
                current_state = STATE_PREV
                # 제목 줄에 코드가 있을 수도 있으므로 continue 하지 않고 아래 로직 수행
            elif line_ns.startswith("대응"):
                current_state = STATE_RESP
            elif line_ns.startswith("저장"):
                current_state = STATE_STOR
            elif line_ns.startswith("폐기"):
                current_state = STATE_DISP

            # 코드 추출 수행
            codes = regex_code.findall(line)
            for code in codes:
                if code.startswith("H"):
                    result["h_codes"].append(code)
                elif code.startswith("P"):
                    if current_state == STATE_PREV:
                        result["p_prev"].append(code)
                    elif current_state == STATE_RESP:
                        result["p_resp"].append(code)
                    elif current_state == STATE_STOR:
                        result["p_stor"].append(code)
                    elif current_state == STATE_DISP:
                        result["p_disp"].append(code)

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
                
                # 중앙 데이터 로드 (매핑용 Dictionary)
                try: 
                    df_master = pd.read_excel(master_data_file, sheet_name=0)
                    code_map = {}
                    for idx, row in df_master.iterrows():
                        # Key 생성 시 공백 완전 제거 (P300 + P310 -> P300+P310)
                        if pd.notna(row.iloc[0]):
                            code_key = str(row.iloc[0]).replace(" ", "").strip()
                            desc_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                            code_map[code_key] = desc_val
                except: 
                    df_master = pd.DataFrame()
                    code_map = {}

                for uploaded_file in uploaded_files:
                    if option == "CFF(K)":
                        try:
                            # 1. PDF 로드 및 파싱 (State Machine)
                            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                            parsed_data = parse_pdf_state_machine(doc)
                            
                            # 2. 양식 파일 준비
                            template_file.seek(0)
                            dest_wb = load_workbook(io.BytesIO(template_file.read()))
                            dest_ws = dest_wb.active

                            # [데이터 동기화]
                            target_sheet = '위험 안전문구'
                            if target_sheet in dest_wb.sheetnames: del dest_wb[target_sheet]
                            data_ws = dest_wb.create_sheet(target_sheet)
                            for r in dataframe_to_rows(df_master, index=False, header=True): data_ws.append(r)

                            # 수식 경로 청소
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
                            
                            # [B20] 유해성 분류 (리스트 -> 문자열)
                            if parsed_data["hazard_cls"]:
                                b20_text = "\n".join(parsed_data["hazard_cls"])
                                dest_ws['B20'] = b20_text
                                dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

                            # [B24] 신호어
                            if parsed_data["signal_word"]:
                                dest_ws['B24'] = parsed_data["signal_word"]
                                dest_ws['B24'].alignment = Alignment(horizontal='center', vertical='center')

                            # [공통 함수] 코드 입력 및 행 숨김/해제 (완전 재작성)
                            def fill_rows_precise(code_list, start_row, end_row):
                                # 1. 중복 제거 및 Key 정규화
                                unique_codes = []
                                for c in code_list:
                                    # PDF 추출 코드 정규화 (공백 제거)
                                    clean_c = c.replace(" ", "").strip()
                                    if clean_c not in unique_codes:
                                        unique_codes.append(clean_c)
                                
                                # 2. [중요] 해당 범위 행 전체 숨김 해제 (Unhide All First)
                                for r in range(start_row, end_row + 1):
                                    dest_ws.row_dimensions[r].hidden = False
                                
                                # 3. 데이터 입력
                                curr = start_row
                                for code in unique_codes:
                                    if curr > end_row: break
                                    
                                    # B열: 원본 코드 (가독성을 위해 원본에 가까운 형태나 정규화된 형태 넣기)
                                    # 매칭을 위해선 공백 제거된 code를 사용하지만, 출력은 깔끔하게
                                    dest_ws.cell(row=curr, column=2).value = code
                                    
                                    # D열: 중앙 데이터 매핑 (수식 덮어쓰기)
                                    # code_map 키도 공백이 제거된 상태이므로 매칭됨
                                    matched_desc = code_map.get(code, "")
                                    dest_ws.cell(row=curr, column=4).value = matched_desc
                                    
                                    curr += 1
                                
                                # 4. [중요] 데이터가 없는 행만 다시 숨김 (Hide Empty)
                                for r in range(start_row, end_row + 1):
                                    cell_val = dest_ws.cell(row=r, column=2).value
                                    # 값이 없거나 공백만 있는 경우 숨김
                                    if cell_val is None or str(cell_val).strip() == "":
                                        dest_ws.row_dimensions[r].hidden = True

                            # [B25~B30] H코드
                            fill_rows_precise(parsed_data["h_codes"], 25, 30)

                            # [B32~B41] 예방 (P_PREV)
                            fill_rows_precise(parsed_data["p_prev"], 32, 41)

                            # [B42~B49] 대응 (P_RESP)
                            fill_rows_precise(parsed_data["p_resp"], 42, 49)

                            # [B50~B52] 저장 (P_STOR)
                            fill_rows_precise(parsed_data["p_stor"], 50, 52)

                            # [B53] 폐기 (P_DISP)
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
                    st.success("완료! PDF 데이터가 정확하게 매핑되었습니다.")
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
