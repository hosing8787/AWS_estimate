import streamlit as st
import os
import json
from openpyxl import load_workbook
import generate_quote
from generate_quote import parse_export

st.set_page_config(page_title="AWS Quote Converter", layout="centered")

st.title("☁️ AWS Estimate to SKR Quote Template")
st.markdown("AWS 계산기에서 내보낸 CSV를 업로드하면, 지정된 단일 탭 형식의 견적서 파일로 즉시 자동 변환합니다.")

uploaded_file = st.file_uploader("AWS Estimate CSV 업로드 (.csv)", type=['csv'])

if uploaded_file is not None:
    st.success("파일 업로드 완료! 변환을 시작합니다.")
    
    with st.spinner("CSV 데이터를 분석하고 변환하는 중..."):
        try:
            # 1. Save uploaded file to inputs/
            os.makedirs("inputs", exist_ok=True)
            csv_path = os.path.join("inputs", uploaded_file.name)
            with open(csv_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
                
            # 2. Convert
            inventory = parse_export(csv_path)
            project_data = {
                "project_info": {
                    "customer_name": "고객사 (자동생성)",
                    "project_name": "AWS Cloud 인프라 환경 구축",
                    "fx_rate": 1405.5,
                    "vat_included": False
                },
                "inventory_items": inventory
            }
            with open("sample_data.json", "w", encoding="utf-8") as f:
                json.dump(project_data, f, ensure_ascii=False, indent=2)

            output_filename = "outputs/Converted_Quote_Result.xlsx"
            os.makedirs("outputs", exist_ok=True)
            
            generate_quote.run_generation(output_filename=output_filename)
            
            st.success("변환이 성공적으로 완료되었습니다!")
            
            # 3. Download Button
            with open(output_filename, "rb") as f:
                st.download_button(
                    label="📥 변환된 견적서 다운로드",
                    data=f,
                    file_name="Converted_AWS_Quote.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        except Exception as e:
            st.error(f"변환 중 오류가 발생했습니다: {e}")
