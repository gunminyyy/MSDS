import streamlit as st
import io
from openpyxl import load_workbook

st.title("🚑 파일 손상 원인 찾기")

uploaded_file = st.file_uploader("문제가 되는 양식 파일을 올려주세요", type="xlsx")

if uploaded_file:
    # 테스트 1: 그냥 그대로 돌려주기 (Byte Copy)
    st.subheader("테스트 1: 단순 복사 (이게 안 열리면 업로드/다운로드 문제)")
    st.download_button(
        label="1. 원본 그대로 다운로드",
        data=uploaded_file.getvalue(),
        file_name="test_original.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # 테스트 2: Openpyxl 거쳐서 저장하기
    st.subheader("테스트 2: 라이브러리 통과 (이게 안 열리면 호환성 문제)")
    if st.button("2. 라이브러리로 읽고 다시 저장하기"):
        try:
            # 포인터 초기화
            uploaded_file.seek(0)
            wb = load_workbook(uploaded_file)
            
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.download_button(
                label="결과 다운로드",
                data=output.getvalue(),
                file_name="test_openpyxl.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"라이브러리가 파일을 읽지도 못했습니다: {e}")

    st.info("💡 팁: '테스트 1'은 되는데 '테스트 2'가 안 된다면, 엑셀 파일을 열어서 [다른 이름으로 저장] 후 다시 시도하세요.")
