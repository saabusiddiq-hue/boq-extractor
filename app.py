import streamlit as st
import pandas as pd
import io
import re
import shutil
from datetime import datetime
from typing import List, Dict, Optional

# Setup
TESSERACT_PATH = shutil.which("tesseract")
OCR_AVAILABLE = False
if TESSERACT_PATH:
    try:
        import pytesseract
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
        OCR_AVAILABLE = True
    except ImportError:
        pass

import pdfplumber
import fitz
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="BOQ Extractor Pro", page_icon="📋", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #0d1117; }
    .upload-bar { background: #161b22; border: 1px solid #30363d; border-radius: 8px; padding: 1rem; margin-bottom: 1rem; }
    .left-panel { background: #161b22; border: 2px solid #30363d; border-radius: 8px; padding: 1rem; }
    .right-panel { background: #161b22; border: 1px solid #30363d; border-radius: 8px; padding: 1rem; min-height: 650px; }
    .success-box { background: rgba(35, 134, 54, 0.1); border: 1px solid #238636; border-radius: 6px; padding: 0.75rem; margin: 0.5rem 0; }
    .warning-box { background: rgba(248, 81, 73, 0.1); border: 1px solid #f85149; border-radius: 6px; padding: 0.75rem; margin: 0.5rem 0; }
</style>
""", unsafe_allow_html=True)

class MaterialSeparator:
    def __init__(self, custom_patterns=None):
        self.custom_patterns = custom_patterns or []
        self.material_patterns = [
            # Compound first
            (r"Graphite\s+Bronze", "Graphite Bronze"),
            (r"Bronze\s+Graphite", "Graphite Bronze"),
            (r"Per\s+MSS\-SP\d+", "Per MSS-SP"),
            (r"MSS\-SP\d+", "MSS-SP"),
            (r"A194\s+GR\.?2H", "A194 GR.2H"),
            (r"A193\s+GR\.?B7", "A193 GR.B7"),
            (r"A240\s+SS316L?", "A240 SS316"),
            (r"SS316L?", "SS316"),
            (r"SS304L?", "SS304"),
            (r"A516", "A516"),
            (r"A240", "A240"),
            (r"A194", "A194"),
            (r"A193", "A193"),
            (r"A105", "A105"),
            (r"A36\b", "A36"),
            (r"Stainless\s+Steel", "Stainless Steel"),
            (r"Carbon\s+Steel", "Carbon Steel"),
            (r"Cast\s+Iron", "Cast Iron"),
            (r"Graphite", "Graphite"),
            (r"Bronze", "Bronze"),
            (r"PTFE", "PTFE"),
        ]

    def separate(self, desc: str) -> tuple:
        if not desc:
            return "", ""
        full = desc.strip()
        for pattern, name in self.custom_patterns + self.material_patterns:
            matches = list(re.finditer(pattern, full, re.IGNORECASE))
            if matches:
                last = matches[-1]
                material = name
                clean = full[:last.start()].strip(" -:/\t")
                clean = re.sub(r"\s+Per\s*$", "", clean, flags=re.IGNORECASE).strip()
                clean = re.sub(r"\s*[-:]\s*$", "", clean).strip()
                return clean, material
        return full, ""

def extract_table_with_pdfplumber(pdf_bytes: bytes, crop: tuple, max_items: int):
    items = []
    sep = MaterialSeparator()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            if crop:
                w, h = page.width, page.height
                page = page.crop((w*crop[0]/100, h*crop[1]/100, w*crop[2]/100, h*crop[3]/100))

            # Try to extract tables
            tables = page.extract_tables()
            for table in tables:
                if not table or len(table) < 2:
                    continue
                # Assume first row is header
                headers = [str(h).strip().lower() if h else "" for h in table[0]]
                for row in table[1:]:  # skip header
                    if len(row) < 4:
                        continue
                    try:
                        item_str = str(row[0]).strip()
                        if not item_str.isdigit():
                            continue
                        item_no = int(item_str)

                        qty = int(str(row[1]).strip() or 1)
                        fig_no = str(row[2]).strip() if len(row) > 2 else ""
                        description = str(row[3]).strip() if len(row) > 3 else ""
                        material = str(row[4]).strip() if len(row) > 4 else ""

                        # If Material column is blank, try to extract from Description
                        if not material or material.lower() in ["", "nan", "none"]:
                            description, material = sep.separate(description)

                        if item_no <= max_items:
                            items.append({
                                "Item": item_no,
                                "Qty": qty,
                                "Fig No": fig_no,
                                "Description": description,
                                "Material": material,
                                "Page": page_num
                            })
                    except:
                        continue
    return items

# Preview function (same as before)
@st.cache_data
def render_preview(pdf_bytes, page_num, crop, zoom=2.0):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page = doc[page_num-1]
        if crop and len(crop)==4:
            x1,y1,x2,y2 = [c/100 * (page.rect.width if i%2==0 else page.rect.height) for i,c in enumerate(crop)]
            shape = page.new_shape()
            shape.draw_rect(fitz.Rect(x1,y1,x2,y2))
            shape.finish(color=(1,0,0), fill=(1,0,0), fill_opacity=0.1, width=2)
            shape.commit()
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        return pix.tobytes("png")
    except:
        return None

def create_excel(df):
    # Same as previous version - omitted for brevity, copy from v30 if needed
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ"
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_font = Font(bold=True, size=10)
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    for col_num, header in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.fill = yellow_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    for row_num, row in enumerate(df.values, 2):
        for col_num, value in enumerate(row, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    for idx in range(1, len(df.columns) + 1):
        col_letter = get_column_letter(idx)
        ws.column_dimensions[col_letter].width = 20

    ws.freeze_panes = "A2"
    wb.save(output)
    return output.getvalue()

# ====================== UI ======================
if "target_bytes" not in st.session_state:
    st.session_state.target_bytes = None
if "data" not in st.session_state:
    st.session_state.data = None

st.title("📋 BOQ Extractor Pro v31 - Grid Aware + Material Separation")

with st.sidebar:
    st.subheader("Settings")
    max_items = st.slider("Max Items", 10, 500, 100)
    st.caption("Now uses table grid detection when possible")

# Main UI
col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("Crop Settings")
    x = st.slider("Left %", 0, 100, 5)
    y = st.slider("Top %", 0, 100, 10)
    z = st.slider("Right %", 0, 100, 95)
    w = st.slider("Bottom %", 0, 100, 65)
    crop = (x, y, z, w)

    if st.session_state.target_bytes:
        img = render_preview(st.session_state.target_bytes, 1, crop, 2.0)
        if img:
            st.image(img, use_column_width=True)

with col2:
    uploaded = st.file_uploader("Upload BOQ PDF", type=["pdf"])
    if uploaded:
        st.session_state.target_bytes = uploaded.read()
        st.success("PDF Loaded")

    if st.session_state.target_bytes and st.button("Extract BOQ", type="primary", use_container_width=True):
        items = extract_table_with_pdfplumber(st.session_state.target_bytes, crop, max_items)
        if items:
            st.session_state.data = pd.DataFrame(items)
            st.success(f"Extracted {len(items)} rows successfully!")
        else:
            st.error("No table found. Try adjusting crop.")

    if st.session_state.data is not None:
        df = st.session_state.data
        edited = st.data_editor(df, use_container_width=True, num_rows="dynamic", hide_index=True)
        st.session_state.data = edited

        col_e, col_c = st.columns(2)
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        with col_e:
            st.download_button("Download Excel", create_excel(edited), f"BOQ_{ts}.xlsx", use_container_width=True)
        with col_c:
            st.download_button("Download CSV", edited.to_csv(index=False).encode(), f"BOQ_{ts}.csv", "text/csv", use_container_width=True)
