import streamlit as st
import io
from excel_converter import convert_excel  # 파일명이 반드시 excel_converter.py 여야 함

st.set_page_config(page_title="Ermore Excel Converter", layout="centered")

st.title("📦 Ermore Excel Converter")
st.write("엑셀 파일을 업로드하면 자동으로 토탈피킹 엑셀을 생성합니다.")

uploaded_file = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    if st.button("엑셀 변환 시작"):
        with st.spinner("엑셀 처리 중입니다..."):
            try:
                # 1. 업로드된 파일을 메모리 버퍼로 읽기
                input_buffer = io.BytesIO(uploaded_file.getvalue())
                output_buffer = io.BytesIO()

                # 2. 변환 함수 호출 (함수 내부에서 Pandas가 이 버퍼를 읽도록 수정 필요)
                # 만약 기존 함수(convert_excel)를 그대로 쓰려면 임시 파일 방식 유지

                # 기존 함수가 파일 경로를 받으므로, 안전하게 처리하려면:
                import tempfile
                import os

                with tempfile.TemporaryDirectory() as tmpdir:
                    input_path = os.path.join(tmpdir, "input.xlsx")
                    output_path = os.path.join(tmpdir, "output.xlsx")

                    with open(input_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    convert_excel(input_path, output_path)

                    with open(output_path, "rb") as f:
                        processed_data = f.read()

                st.success("완료되었습니다!")
                st.download_button(
                    label="📥 변환된 엑셀 다운로드",
                    data=processed_data,
                    file_name=f"작업완료_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")