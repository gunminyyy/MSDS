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
st.title("MSDS 양식 변환기 (PDF 정밀 파싱 - 최종 교정)")
st.markdown("---")

# --------------------------------------------------------------------------
# [함수] 이미지 처리 (기존 유지)
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
# [신규 함수] PDF 섹션 정밀 파싱 (노이즈 필터링 및 섹션 분리)
# --------------------------------------------------------------------------
def parse_pdf_ghs_logic(doc):
    full_text_lines = []
    for page in doc:
        text = page.get_text("text")
        lines = text.split('\n')
        full_text_lines.extend(lines)

    # 결과 저장소
    result = {
        "hazard_cls": [],       # B20
        "signal_word": "",      # B24
        "h_codes": [],          # B25:30
        "p_prev": [],           # B32:41 (예방)
        "p_resp": [],           # B42:49 (대응)
        "p_stor": [],           # B50:52 (저장)
        "p_disp": []            # B53 (폐기)
    }

    # 상태 관리
    current_section = None # 'HAZARD_CLS', 'H_CODE', 'P_PREV', 'P_RESP', 'P_STOR', 'P_DISP'
    
    # 노이즈 필터 (헤더/푸터 등 무시할 단어들)
    NOISE_KEYWORDS = [
        "물질안전보건자료", "MSDS", "Material Safety Data Sheet",
        "Corea flavors", "주식회사 고려", "HAIR CARE", "Ver.", "발행일",
        "제 품 명", "개정일"
    ]

    for line in full_text_lines:
        clean_line = line.strip()
        if not clean_line: continue

        # 1. 노이즈 제거 (반복되는 헤더 무시)
        is_noise = False
        for kw in NOISE_KEYWORDS:
            if kw in clean_line:
                is_noise = True
                break
        if is_noise: continue

        # 공백 제거 버전 (키워드 매칭용)
        line_nospace = clean_line.replace(" ", "")

        # ------------------- 섹션 감지 및 전환 -------------------

        # [B20] 유해성 분류 시작
        if "가.유해성" in line_nospace and "분류" in line_nospace:
            current_section = "HAZARD_CLS"
            continue # 제목 줄은 저장 안 함

        # [B24] 신호어 (어디에 있든 찾아서 저장)
        if "신호어" in line_nospace:
            # "신호어 : 위험" 형태 처리
            parts = clean_line.split(":")
            if len(parts) > 1:
                result["signal_word"] = parts[-1].strip()
            else:
                # 같은 줄에 없고 다음 줄에 있을 수도 있지만, 보통 같은 줄에 있음
                # "신호어 위험" 처럼 공백으로 구분된 경우
                result["signal_word"] = clean_line.replace("신호어", "").strip()
            continue

        # [H코드] 유해 위험 문구 시작 -> B20 수집 종료
        if "유해" in line_nospace and "위험문구" in line_nospace:
            current_section = "H_CODE"
            continue

        # [P코드] 예방조치문구 시작 (큰 제목)
        if "예방조치문구" in line_nospace:
            # 아직 세부 섹션(예방, 대응...)을 모르므로 대기 상태
            current_section = "WAITING_P"
            continue

        # 나. 예방조치...항목 -> B20 종료 조건 (혹시 위에서 못 잡았을 경우)
        if "나.예방조치" in line_nospace and "항목" in line_nospace:
            if current_section == "HAZARD_CLS":
                current_section = "WAITING_P"
            continue

        # P코드 세부 섹션 감지 (예방, 대응, 저장, 폐기)
        # 주의: 문장 속에 '예방'이 들어갈 수 있으므로, 줄의 시작이거나 명확한 헤더일 때만
        if line_nospace.startswith("예방"):
            current_section = "P_PREV"
            continue
        elif line_nospace.startswith("대응"):
            current_section = "P_RESP"
            continue
        elif line_nospace.startswith("저장"):
            current_section = "P_STOR"
            continue
        elif line_nospace.startswith("폐기"):
            current_section = "P_DISP"
            continue

        # 3. 구성성분 (섹션 종료)
        if "3.구성성분" in line_nospace or "다.기타" in line_nospace:
            current_section = "DONE"
            break

        # ------------------- 데이터 수집 -------------------

        if current_section == "HAZARD_CLS":
            # 가. 제목 줄은 이미 건너뛰었으므로 내용만 담김
            result["hazard_cls"].append(clean_line)

        elif current_section == "H_CODE":
            # H코드 추출 (H300)
            codes = re.findall(r"H\d{3}", clean_line)
            result["h_codes"].extend(codes)

        elif current_section in ["P_PREV", "P_RESP", "P_STOR", "P_DISP"]:
            # P코드 추출 (복합 코드 P300+P310 지원)
            # 정규식 설명: P숫자3개로 시작하고, (+P숫자3개)가 0번 이상 반복되는 패턴
            codes = re.findall(r"P\d{3}(?:\+P\d{3})*", clean_line)
            
            if current_section == "P_PREV":
                result["p_prev"].extend(codes)
            elif current_section == "P_RESP":
                result["p_resp"].extend(codes)
            elif current_section == "P_STOR":
                result["p_stor"].extend(codes)
            elif current_section == "P_DISP":
                result["p_disp"].extend(codes)

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
            with st.spinner("PDF 정밀 분석 및 변환 중..."):
                
                new_files = []
                new_download_data = {}
                
                # 중앙 데이터 로드 (매핑용) - 공백 제거하여 Key 생성
                try: 
                    df_master = pd.read_excel(master_data_file, sheet_name=0)
                    code_map = {}
                    for idx, row in df_master.iterrows():
                        # 코드의 공백 제거 (P300 + P310 -> P300+P310)
                        code_val = str(row.iloc[0]).replace(" ", "").strip()
                        desc_val = str(row.iloc[1]).strip()
                        code_map[code_val] = desc_val
                except: 
                    df_master = pd.DataFrame()
                    code_map = {}

                for uploaded_file in uploaded_files:
                    if option == "CFF(K)":
                        try:
                            # 1. PDF 로드 및 파싱 (새 로직)
                            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                            parsed_data = parse_pdf_ghs_logic(doc)
                            
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
                            
                            # [B20] 유해성 분류
                            # 리스트 내용을 줄바꿈으로 연결
                            if parsed_data["hazard_cls"]:
                                b20_text = "\n".join(parsed_data["hazard_cls"])
                                dest_ws['B20'] = b20_text
                                dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

                            # [B24] 신호어
                            if parsed_data["signal_word"]:
                                dest_ws['B24'] = parsed_data["signal_word"]
                                dest_ws['B24'].alignment = Alignment(horizontal='center', vertical='center')

                            # [공통 함수] 코드 입력 및 행 숨김/해제 처리
                            def fill_rows(code_list, start_row, end_row):
                                # 중복 제거 (순서 유지)
                                unique_codes = []
                                for c in code_list:
                                    # 공백 제거 정규화
                                    norm_c = c.replace(" ", "").strip()
                                    if norm_c not in unique_codes: unique_codes.append(norm_c)
                                
                                # 1. 범위 내 모든 행 숨김 취소 (초기화)
                                for r in range(start_row, end_row + 1):
                                    dest_ws.row_dimensions[r].hidden = False
                                
                                # 2. 데이터 입력
                                curr = start_row
                                for code in unique_codes:
                                    if curr > end_row: break # 칸 부족하면 멈춤
                                    
                                    # B열: 코드
                                    dest_ws.cell(row=curr, column=2).value = code
                                    # D열: 내용 매칭
                                    matched_text = code_map.get(code, "") 
                                    dest_ws.cell(row=curr, column=4).value = matched_text
                                    
                                    curr += 1
                                
                                # 3. 데이터 없는 행 다시 숨김 처리
                                for r in range(start_row, end_row + 1):
                                    val = dest_ws.cell(row=r, column=2).value
                                    if val is None or str(val).strip() == "":
                                        dest_ws.row_dimensions[r].hidden = True

                            # [B25~B30] H코드
                            fill_rows(parsed_data["h_codes"], 25, 30)

                            # [B32~B41] 예방 (P_PREV)
                            fill_rows(parsed_data["p_prev"], 32, 41)

                            # [B42~B49] 대응 (P_RESP)
                            fill_rows(parsed_data["p_resp"], 42, 49)

                            # [B50~B52] 저장 (P_STOR)
                            fill_rows(parsed_data["p_stor"], 50, 52)

                            # [B53] 폐기 (P_DISP)
                            fill_rows(parsed_data["p_disp"], 53, 53)

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
                    st.success("완료! PDF 정밀 변환이 끝났습니다.")
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
