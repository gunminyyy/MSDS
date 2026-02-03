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

# 1. 페이지 설정
st.set_page_config(page_title="MSDS 스마트 변환기", layout="wide")
st.title("MSDS 양식 변환기 (정밀 인식 모드)")
st.markdown("---")

# --------------------------------------------------------------------------
# [함수] 이미지 정규화 (투명 배경 제거 -> 흰색 배경으로 통일)
# --------------------------------------------------------------------------
def normalize_image(pil_img):
    """이미지를 32x32 크기의 흑백(Grayscale)으로 변환하되, 투명한 부분은 흰색으로 채움"""
    try:
        # RGBA(투명도 포함)라면 흰색 배경을 깔아줌
        if pil_img.mode in ('RGBA', 'LA') or (pil_img.mode == 'P' and 'transparency' in pil_img.info):
            # 흰색 배경 캔버스 생성
            background = PILImage.new('RGB', pil_img.size, (255, 255, 255))
            # 투명도가 있는 이미지를 위에 덮어씌움 (마스크 사용)
            if pil_img.mode == 'P':
                pil_img = pil_img.convert('RGBA')
            background.paste(pil_img, mask=pil_img.split()[3]) # 3번 채널이 Alpha
            pil_img = background
        else:
            pil_img = pil_img.convert('RGB')
            
        # 32x32로 리사이징하고 흑백 변환
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

# [함수] 이미지 비교 매칭 (개선됨)
def find_best_match_name(src_img, ref_images):
    best_score = float('inf')
    best_name = None
    
    try:
        # 원본 이미지 정규화 (흰배경+흑백)
        src_norm = normalize_image(src_img)
        src_arr = np.array(src_norm, dtype=np.int16)
        
        for name, ref_img in ref_images.items():
            # 기준 이미지 정규화
            ref_norm = normalize_image(ref_img)
            ref_arr = np.array(ref_norm, dtype=np.int16)
            
            # 차이 계산
            diff = np.mean(np.abs(src_arr - ref_arr))
            
            if diff < best_score:
                best_score = diff
                best_name = name
        
        # 임계값: 0(완벽일치) ~ 255(완전반대). 50 이하면 꽤 비슷한 그림
        if best_score < 65: 
            return best_name
        else: 
            return None
    except: return None

# [함수] 파일명에서 숫자 추출 (정렬용)
def extract_number(filename):
    # "1.tif" -> 1, "10.png" -> 10 변환 (숫자가 없으면 999)
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
    uploaded_files = st.file_uploader("원본 데이터(엑셀)", type=["xlsx"], accept_multiple_files=True)

with col_center:
    st.write("") ; st.write("") ; st.write("")
    
    if st.button("▶ 변환 시작", use_container_width=True):
        if uploaded_files and master_data_file and template_file:
            with st.spinner("그림 분석 및 정밀 정렬 중..."):
                
                new_files = []
                new_download_data = {}
                
                # 중앙 데이터 읽기
                try: df_master = pd.read_excel(master_data_file, sheet_name=0)
                except: df_master = pd.DataFrame()

                for uploaded_file in uploaded_files:
                    if option == "CFF(K)":
                        try:
                            # 1. 파일 로드
                            src_wb = load_workbook(uploaded_file, data_only=True)
                            src_ws = src_wb.active
                            
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

                            # 제품명 및 텍스트 복사
                            dest_ws['B7'] = product_name_input
                            dest_ws['B10'] = product_name_input
                            
                            start_row = 0; end_row = 0
                            for i, row in enumerate(src_ws.iter_rows(values_only=True), 1):
                                row_text = " ".join([str(c) for c in row if c])
                                if "2. 유해성" in row_text and "위험성" in row_text: start_row = i
                                if "나. 예방조치문구를 포함한 경고표지 항목" in row_text: end_row = i; break
                            
                            if start_row > 0 and end_row > 0:
                                texts = []
                                for r in range(start_row + 1, end_row):
                                    val = src_ws.cell(row=r, column=4).value
                                    if val: texts.append(str(val).strip())
                                dest_ws['B20'] = "\n".join(texts)
                                dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

                            # ---------------------------------------------------
                            # [핵심 수정] 그림 정밀 인식 및 정렬
                            # ---------------------------------------------------
                            
                            # 1. 기존 그림 삭제
                            target_anchor_row = 22
                            if hasattr(dest_ws, '_images'):
                                preserved_imgs = []
                                for img in dest_ws._images:
                                    try:
                                        if not (target_anchor_row - 2 <= img.anchor._from.row <= target_anchor_row + 2):
                                            preserved_imgs.append(img)
                                    except: preserved_imgs.append(img)
                                dest_ws._images = preserved_imgs

                            # 2. 원본 그림 수집 & 매칭
                            img_row = 0
                            for i, row in enumerate(src_ws.iter_rows(values_only=True), 1):
                                if "그림문자" in str(row[0]): img_row = i; break
                            
                            collected_pil_images = []
                            matched_names = [] # 디버깅용: 어떤 파일로 인식됐는지 기록
                            
                            if img_row > 0 and hasattr(src_ws, '_images'):
                                for img in src_ws._images:
                                    if hasattr(img, 'anchor'):
                                        r = img.anchor._from.row
                                        if img_row - 2 <= r <= img_row + 1:
                                            if hasattr(img, '_data'):
                                                pil_img = PILImage.open(io.BytesIO(img._data()))
                                                
                                                # [인식]
                                                matched_name = None
                                                if loaded_refs:
                                                    matched_name = find_best_match_name(pil_img, loaded_refs)
                                                
                                                if matched_name:
                                                    matched_names.append(matched_name)
                                                    # 정렬 키: 파일명에서 숫자 추출 (예: '2.tif' -> 2)
                                                    sort_key = extract_number(matched_name)
                                                    collected_pil_images.append((sort_key, pil_img))
                                                else:
                                                    # 인식 실패 시 9999번으로 맨 뒤로 보냄
                                                    matched_names.append("인식실패")
                                                    collected_pil_images.append((9999, pil_img))
                            
                            # 3. 정렬 (숫자 오름차순: 1 -> 2 -> 3...)
                            collected_pil_images.sort(key=lambda x: x[0])
                            sorted_imgs = [item[1] for item in collected_pil_images]
                            
                            # 화면에 인식 결과 표시 (디버깅)
                            if matched_names:
                                st.info(f"🔍 인식된 그림 목록: {', '.join(matched_names)}")
                            
                            # 4. 그림 합치기 (Stitching)
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
                                    x_pos = (idx * unit_size) + padding_left
                                    y_pos = padding_top
                                    merged_img.paste(p_img_resized, (x_pos, y_pos))
                                
                                img_byte_arr = io.BytesIO()
                                merged_img.save(img_byte_arr, format='PNG') 
                                img_byte_arr.seek(0)
                                
                                final_xl_img = XLImage(img_byte_arr)
                                dest_ws.add_image(final_xl_img, 'B23')

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
                if 'src_wb' in locals(): del src_wb
                if 'dest_wb' in locals(): del dest_wb
                if 'output' in locals(): del output
                gc.collect()

                if new_files:
                    st.success("완료! 그림들이 번호 순서대로 정렬되었습니다.")
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
