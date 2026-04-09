import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import os
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="Grok-K2.5 BOQ Extractor", page_icon="📊", layout="centered")
st.title("🚀 Grok-K2.5 BOQ Extractor Web")
st.markdown("**Solo Purpose:** Extract complete BOQ from pipe support drawing PDFs → Excel")
st.caption("K2.5 Thinking Mode | Made for your KECSA / Q24250 drawings")

uploaded_files = st.file_uploader("Upload your drawing PDF(s)", type="pdf", accept_multiple_files=True)

if st.button("🔥 Extract BOQ Now", type="primary", use_container_width=True):
    if not uploaded_files:
        st.error("Please upload at least one PDF")
        st.stop()

    all_data = []
    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, uploaded_file in enumerate(uploaded_files):
        status_text.text(f"Processing: {uploaded_file.name} ({idx+1}/{len(uploaded_files)})")
        
        # Save temporarily
        with open(uploaded_file.name, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        doc = fitz.open(uploaded_file.name)
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            text = page.get_text("text")
            
            # Extract Drawing No
            drawing_match = re.search(r'DRAWING No\s+([Q0-9A-Z\- ]+)', text, re.I)
            drawing_no = drawing_match.group(1).strip() if drawing_match else f"Unknown_Page_{page_num+1}"
            
            # Extract Mark No
            mark_match = re.search(r'SUPPORT MARK No\s+([A-Z0-9\-]+)', text, re.I)
            mark_no = mark_match.group(1).strip() if mark_match else ""
            
            # Extract the item table
            tables = page.find_tables()
            if tables:
                for table in tables:
                    df_table = table.to_pandas()
                    if len(df_table.columns) >= 4 and any("Item" in str(col) for col in df_table.columns):
                        for _, row in df_table.iterrows():
                            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "":
                                continue
                            all_data.append({
                                "Drawing No": drawing_no,
                                "Mark No": mark_no,
                                "Item No": str(row.iloc[0]),
                                "Quantity": str(row.iloc[1]) if len(row) > 1 else "1",
                                "Part No / Fig No": str(row.iloc[2]) if len(row) > 2 else "",
                                "Description": str(row.iloc[3]) if len(row) > 3 else "",
                                "Material": str(row.iloc[4]) if len(row) > 4 else "",
                                "Tube Length / Notes": re.search(r'Tube Length.*|Note:.*|Serial No.*|1\.SS PLATE.*', text, re.I).group(0) if re.search(r'Tube Length.*|Note:.*|Serial No.*|1\.SS PLATE.*', text, re.I) else "",
                                "Source Page": page_num + 1,
                                "Source File": uploaded_file.name
