import streamlit as st
import tempfile
import os
from excel_converter import convert_excel

st.set_page_config(page_title="Ermore Excel Converter", layout="centered")

st.title("📦 Ermore Excel Converter")
st.write("엑셀 파일을 업로드하면 자동으로 토탈피킹 엑셀을 생성합니다.")

uploaded_file = st.file_uploader(
    "엑셀 파일 업로드 (.xlsx)",
    type=["xlsx"]
)

if uploaded_file is not None:
    if st.button("엑셀 변환 시작"):
        with st.spinner("엑셀 처리 중입니다..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                input_path = os.path.join(tmpdir, uploaded_file.name)
                output_path = os.path.join(tmpdir, f"작업완료_{uploaded_file.name}")

                # 업로드 파일 저장
                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # 변환 실행 (⭐ 여기서만 호출)
                convert_excel(input_path, output_path)

                # 다운로드 버튼
                with open(output_path, "rb") as f:
                    st.success("완료되었습니다!")
                    st.download_button(
                        label="📥 변환된 엑셀 다운로드",
                        data=f,
                        file_name=f"작업완료_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )