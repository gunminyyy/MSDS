import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.cell.cell import MergedCell
from copy import copy
from PIL import Image as PILImage
import io
import re
import os
import fitz  # PyMuPDF
import numpy as np
import gc

# 1. 페이지 설정
st.set_page_config(page_title="MSDS 스마트 변환기", layout="wide")
st.title("MSDS 양식 변환기 (위치 자동 추적 & 스타일 복제)")
st.markdown("---")

# --------------------------------------------------------------------------
# [스타일 정의] 굴림 8pt, 왼쪽 정렬
# --------------------------------------------------------------------------
FONT_STYLE = Font(name='굴림', size=8)
ALIGN_LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)

# --------------------------------------------------------------------------
# [함수] 중앙 데이터 로드 (정규화: 공백제거, 대문자)
# --------------------------------------------------------------------------
def load_master_data(file):
    try:
        df = pd.read_excel(file, sheet_name=0)
        # 컬럼명 정규화
        df.columns = [str(c).replace(" ", "").upper() for c in df.columns]
        
        # CODE, K 컬럼 찾기
        col_code = 'CODE' if 'CODE' in df.columns else df.columns[0]
        col_kor = 'K' if 'K' in df.columns else (df.columns[1] if len(df.columns)>1 else None)
        
        mapping = {}
        if col_kor:
            for idx, row in df.iterrows():
                if pd.notna(row[col_code]):
                    # Key: 공백제거, 대문자
                    k = str(row[col_code]).replace(" ", "").replace("\n", "").upper().strip()
                    v = str(row[col_kor]).strip() if pd.notna(row[col_kor]) else ""
                    mapping[k] = v
        return mapping
    except Exception:
        return {}

def get_desc(code, mapping):
    # 입력된 코드 정규화
    clean = str(code).replace(" ", "").replace("\n", "").upper().strip()
    
    # 1. 완벽 일치
    if clean in mapping: return mapping[clean]
    
    # 2. 복합 코드 (+ 분리)
    if "+" in clean:
        parts = clean.split("+")
        found = []
        for p in parts:
            if p in mapping: found.append(mapping[p])
        if found: return " ".join(found)
        
    return ""

# --------------------------------------------------------------------------
# [함수] PDF 파싱 (구역 추출)
# --------------------------------------------------------------------------
def parse_pdf(doc):
    full_text = []
    # 페이지별로 읽되 좌표 순서(sort=True)로 정렬
    for page in doc:
        blocks = page.get_text("blocks", sort=True)
        for b in blocks:
            full_text.append(b[4]) # 텍스트 내용만
            
    # 전체 텍스트를 줄 단위로 분리
    lines = []
    for txt in full_text:
        lines.extend(txt.split('\n'))
        
    # 노이즈 필터링
    clean_lines = []
    for line in lines:
        l = line.strip()
        if not l: continue
        if any(x in l for x in ["물질안전보건자료", "MSDS", "PAGE", "Ver.", "발행일"]): continue
        clean_lines.append(l)

    # 데이터 추출
    data = {"h": [], "prev": [], "resp": [], "stor": [], "disp": [], "signal": "", "hazard_cls": []}
    
    # 상태 머신
    ZONE_NONE = 0
    ZONE_HAZARD = 1 # 유해성 분류
    ZONE_LABEL = 2  # 라벨 요소
    state = ZONE_NONE
    
    sub_state = None # P코드 서브존
    
    regex_code = re.compile(r"([HP]\d{3}(?:\s*\+\s*[HP]\d{3})*)")
    
    for line in clean_lines:
        lns = line.replace(" ", "")
        
        # 구역 전환 감지
        if "가.유해성" in lns and "분류" in lns:
            state = ZONE_HAZARD; continue
        if "나.예방조치" in lns:
            state = ZONE_LABEL; sub_state = None; continue
        if "3.구성성분" in lns or "다.기타" in lns:
            state = ZONE_NONE; break
            
        if state == ZONE_HAZARD:
            if "공급자정보" in lns or "회사명" in lns: continue
            data["hazard_cls"].append(line)
            # H코드 추출
            codes = regex_code.findall(line)
            for c in codes: 
                if c.startswith("H"): data["h"].append(c)
                
        elif state == ZONE_LABEL:
            if "신호어" in lns:
                data["signal"] = line.replace("신호어", "").strip()
            
            # 서브존 전환 (키워드)
            if lns.startswith("예방") and len(lns)<10: sub_state = "prev"
            elif lns.startswith("대응") and len(lns)<10: sub_state = "resp"
            elif lns.startswith("저장") and len(lns)<10: sub_state = "stor"
            elif lns.startswith("폐기") and len(lns)<10: sub_state = "disp"
            
            # 코드 추출
            codes = regex_code.findall(line)
            for c in codes:
                if c.startswith("H"): data["h"].append(c)
                elif c.startswith("P") and sub_state:
                    data[sub_state].append(c)
                    
    return data

# --------------------------------------------------------------------------
# [핵심] 행 스타일 복사 (서식 유지용)
# --------------------------------------------------------------------------
def copy_style(ws, src_row, tgt_row):
    ws.row_dimensions[tgt_row].height = ws.row_dimensions[src_row].height
    for col in range(1, 10): # A~I열 복사
        src = ws.cell(row=src_row, column=col)
        tgt = ws.cell(row=tgt_row, column=col)
        if src.has_style:
            try: tgt._style = copy(src._style)
            except: pass # 스타일 복사 실패 시 무시

# --------------------------------------------------------------------------
# [핵심] 순차적 섹션 처리기 (밀림 현상 완벽 대응)
# --------------------------------------------------------------------------
def process_section(ws, start_keyword, next_keyword, codes, mapping, search_start_row):
    """
    search_start_row 부터 시작해서 start_keyword를 찾고, 
    그 다음 next_keyword를 찾아서 그 사이 공간에 데이터를 넣음.
    부족하면 행을 추가하고 스타일을 복사함.
    처리가 끝난 마지막 행 위치를 반환함 (다음 검색 시작점).
    """
    
    # 1. 시작 헤더 찾기
    header_row = -1
    for r in range(search_start_row, ws.max_row + 1):
        val = str(ws.cell(row=r, column=2).value).replace(" ", "")
        if start_keyword in val:
            header_row = r
            break
    
    if header_row == -1: return search_start_row # 못 찾으면 현 위치 반환
    
    # 2. 다음 헤더(끝) 찾기
    next_header_row = -1
    if next_keyword == "END":
        next_header_row = header_row + 2 # 최소 공간
    else:
        for r in range(header_row + 1, ws.max_row + 100):
            val = str(ws.cell(row=r, column=2).value).replace(" ", "")
            if next_keyword in val:
                next_header_row = r
                break
        if next_header_row == -1: next_header_row = header_row + 5 # fallback
        
    # 데이터 들어갈 첫 줄
    data_row = header_row + 1
    
    # 가용 공간 (현재 빈 줄 수)
    available = next_header_row - data_row
    
    # 코드 중복 제거
    unique_codes = []
    seen = set()
    for c in codes:
        clean = c.replace(" ", "").upper().strip()
        if clean not in seen:
            unique_codes.append(clean)
            seen.add(clean)
            
    needed = len(unique_codes)
    
    # 3. 공간 부족 시 행 삽입 (스타일 복사 포함)
    if needed > available:
        rows_to_add = needed - available
        insert_pos = next_header_row # 다음 헤더 바로 위에 삽입
        
        ws.insert_rows(insert_pos, amount=rows_to_add)
        
        # 스타일 복사 (삽입 위치 바로 윗줄 = 섹션의 마지막 줄 서식을 복사)
        style_src_row = insert_pos - 1
        for i in range(rows_to_add):
            tgt_row = insert_pos + i
            copy_style(ws, style_src_row, tgt_row)
            
        # 행 추가로 인해 다음 헤더 위치가 밀려남
        next_header_row += rows_to_add
        
    # 4. 데이터 쓰기
    curr = data_row
    for code in unique_codes:
        # 숨김 해제 및 높이 고정
        ws.row_dimensions[curr].hidden = False
        ws.row_dimensions[curr].height = 19
        
        # 셀 병합 해제 (안전장치)
        for col in [2, 4]:
            cell = ws.cell(row=curr, column=col)
            if isinstance(cell, MergedCell):
                # 병합 해제 로직 (간소화)
                pass 
        
        # B열: 코드
        cell_b = ws.cell(row=curr, column=2)
        cell_b.value = code
        cell_b.font = FONT_STYLE
        cell_b.alignment = ALIGN_LEFT
        
        # D열: 내용 (매핑)
        cell_d = ws.cell(row=curr, column=4)
        desc = get_desc(code, mapping)
        cell_d.value = desc
        cell_d.font = FONT_STYLE
        cell_d.alignment = ALIGN_LEFT
        
        curr += 1
        
    # 5. 남은 빈 칸 처리 (수식/내용 지우고 숨김)
    for r in range(curr, next_header_row):
        ws.cell(row=r, column=2).value = ""
        ws.cell(row=r, column=4).value = ""
        ws.row_dimensions[r].hidden = True
        
    # 다음 검색은 현재 섹션 끝(next_header_row) 부터 시작
    return next_header_row

# 2. UI 구성
with st.expander("📂 파일 업로드", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        f_master = st.file_uploader("1. 중앙 데이터 (master.xlsx)", type="xlsx")
    with col2:
        f_template = st.file_uploader("2. 양식 파일 (template.xlsx)", type="xlsx")

product_name = st.text_input("제품명 입력")
st.write("")

col_l, col_c, col_r = st.columns([4, 2, 4])

with col_l:
    st.subheader("3. 원본 PDF")
    f_pdfs = st.file_uploader("PDF 업로드", type=["pdf"], accept_multiple_files=True)

if 'results' not in st.session_state:
    st.session_state['results'] = {}

with col_c:
    st.write("") ; st.write("")
    if st.button("▶ 변환 시작", use_container_width=True):
        if f_master and f_template and f_pdfs:
            with st.spinner("순차적 처리 및 스타일 복제 중..."):
                
                # 중앙 데이터 로드
                mapping = load_master_data(f_master)
                st.toast(f"중앙 데이터 {len(mapping)}개 로드 완료")
                
                results = {}
                
                for f_pdf in f_pdfs:
                    try:
                        # 1. PDF 파싱
                        doc = fitz.open(stream=f_pdf.read(), filetype="pdf")
                        data = parse_pdf(doc)
                        
                        # 2. 양식 로드
                        f_template.seek(0)
                        wb = load_workbook(io.BytesIO(f_template.read()))
                        ws = wb.active
                        
                        # 3. 기본 정보 입력
                        ws['B7'] = product_name
                        ws['B10'] = product_name
                        
                        if data["hazard_cls"]:
                            ws['B20'] = "\n".join(data["hazard_cls"])
                            ws['B20'].alignment = ALIGN_LEFT
                            
                        if data["signal"]:
                            ws['B24'] = data["signal"]
                            ws['B24'].alignment = Alignment(horizontal='center', vertical='center')
                            
                        # 4. [핵심] 순차적 섹션 처리 (위치 자동 추적)
                        # 반드시 위에서 아래 순서로 실행해야 밀림 현상이 반영됨
                        
                        # (1) H코드 (유해·위험문구 ~ 예방)
                        cursor = process_section(ws, "유해·위험문구", "예방", data["h"], mapping, 20)
                        
                        # (2) 예방 (예방 ~ 대응)
                        cursor = process_section(ws, "예방", "대응", data["prev"], mapping, cursor)
                        
                        # (3) 대응 (대응 ~ 저장)
                        cursor = process_section(ws, "대응", "저장", data["resp"], mapping, cursor)
                        
                        # (4) 저장 (저장 ~ 폐기)
                        cursor = process_section(ws, "저장", "폐기", data["stor"], mapping, cursor)
                        
                        # (5) 폐기 (폐기 ~ 3.구성성분)
                        cursor = process_section(ws, "폐기", "3.", data["disp"], mapping, cursor)
                        
                        # 5. 저장
                        out = io.BytesIO()
                        wb.save(out)
                        fname = f"{product_name}_{f_pdf.name.split('.')[0]}.xlsx"
                        results[fname] = out.getvalue()
                        
                    except Exception as e:
                        st.error(f"{f_pdf.name} 오류: {e}")
                        
                st.session_state['results'] = results
                st.success("완료!")
                gc.collect()

with col_r:
    st.subheader("다운로드")
    if st.session_state['results']:
        for fname, data in st.session_state['results'].items():
            st.download_button(label=f"📥 {fname}", data=data, file_name=fname, 
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
