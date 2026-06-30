import streamlit as st
import tempfile
import os
from pathlib import Path
from convert import convert

st.set_page_config(page_title="Azure Cost Converter", page_icon="☁️", layout="centered")

st.title("☁️ Azure Cost Estimation Converter")
st.write("Upload your Azure Calculator export to mathematically deduct licenses and fetch RI pricing.")

currency = st.selectbox("Select Target Currency:", ["INR", "USD", "EUR", "GBP", "AUD"])
uploaded_file = st.file_uploader("Upload Azure Export (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    if st.button("Convert File", type="primary"):
        # Use tempfile to prevent race conditions in multi-user Web App environments
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / f"input_{uploaded_file.file_id}.xlsx"
            output_path = Path(temp_dir) / f"Processed_Estimate_{currency}.xlsx"
            
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            with st.spinner("Querying Azure Retail Pricing API & processing..."):
                try:
                    # Pass a fresh session cache to prevent memory leaks
                    convert(str(input_path), str(output_path), currency)
                    
                    with open(output_path, "rb") as f:
                        file_data = f.read()
                        
                    st.success("Conversion Complete!")
                    st.download_button(
                        label="📥 Download Processed Estimate",
                        data=file_data,
                        file_name=f"Processed_Estimate_{currency}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except ValueError as ve:
                    st.error(f"Validation Error: {ve}")
                except Exception as e:
                    st.error(f"An unexpected error occurred during processing: {str(e)}")
