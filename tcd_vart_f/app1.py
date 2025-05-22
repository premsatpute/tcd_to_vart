import streamlit as st
import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side
import io
from preprocessor import load_and_preprocess_tcd,generate_vart_sheet,extract_steps

# Step 1: Load and preprocess the TCD sheet


st.title("TCD to VART Converter")
st.write("Convert your TCD to VART in one click")

uploaded_file = st.file_uploader("Upload your TCD Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.success("✅ File uploaded successfully!")
    try:
        df = pd.read_excel(uploaded_file)
        processed_df = load_and_preprocess_tcd(df.copy())
        processed_df["Steps"] = processed_df.apply(extract_steps, axis=1)
        wb, vart_data = generate_vart_sheet(processed_df)

        # Display preview
        st.subheader("Output Sheet Preview")
        preview_df = pd.DataFrame(vart_data)
        st.dataframe(preview_df)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.download_button(
            label="Download VART Sheet",
            data=output,
            file_name="VART_Sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"An error occurred: {e}")