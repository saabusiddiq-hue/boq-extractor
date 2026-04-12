import streamlit as st
import pandas as pd
import io
import re
import shutil
from datetime import datetime
from typing import List, Dict, Tuple

# Setup Libraries
import pdfplumber
import fitz
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="BOQ Extractor Pro v32", page_icon="📋", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #0d1117; }
    .left-panel { background: #161b22; border: 2px solid #30363d; border-radius: 8px; padding: 1rem; }
    .right-panel { background: #161b22; border: 1px solid #30363d; border-radius: 8px; padding: 1rem; min-height: 650px; }
    .success-box { background: rgba(35, 134, 54, 0.1); border: 1px solid #238636; border-radius: 6px; padding: 0.75rem; }
</style>
""", unsafe_allow_html=True)

class MaterialSeparator:
    def __init__(self):
        self.patterns = [
            (r"Graphite\s+Bronze", "Graphite Bronze"),
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

    def separate(self, text: str) -> Tuple[str, str]:
        if not text:
            return "", ""
        for pattern, mat in self.patterns:
            matches = list(re.finditer(pattern, text, re.IGNORECASE))
            if matches:
                last = matches[-1]
                material = mat
                desc = text[:last.start()].strip(" -:/\t")
                desc = re.sub(r"\s+Per\s*$", "", desc, flags=re.IGNORECASE).strip()
                return desc, material
        return text.strip(), ""

# ====================== EXTRACTION ENGINE ======================
def auto_extract_boq(pdf_bytes: bytes, crop: tuple, max_items: int, use_auto_grid: bool):
    items = []
    separator = MaterialSeparator()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            if crop:
                w, h = page.width, page.height
                page = page.crop((w * crop[0]/100, h * crop[1]/100, w * crop[2]/100, h * crop[3]/100))

            if use_auto_grid:
                # Auto table detection (works even without lines)
                tables = page.extract_tables(table_settings={
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "explicit_vertical_lines": [],
                    "explicit_horizontal_lines": [],
                    "snap_tolerance": 3,
                    "join_tolerance": 3,
                })
            else:
                tables = page.extract_tables()

            for table in tables:
                if not table or len(table) < 2:
                    continue
                for row in table[1:]:   # Skip header row
                    if len(row) < 3:
                        continue
                    try:
                        item_str = str(row[0]).strip()
                        if not item_str or not item_str[0].isdigit():
                            continue
                        item_no = int(item_str.split()[0])

                        qty = int(str(row[1]).strip() or 1)
                        fig_no = str(row[2]).strip() if len(row) > 2 else ""
                        desc_raw = str(row[3]).strip() if len(row) > 3 else ""
                        mat_raw = str(row[4]).strip() if len(row) > 4 else ""

                        # Smart Material Handling
                        if mat_raw and mat_raw.lower() not in ["", "nan", "none", "-"]:
                            material = mat_raw
                            description = desc_raw
                        else:
                            description, material = separator.separate(desc_raw)

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

            # Fallback: Old line-by-line method if no table found
            if not items and not use_auto_grid:
                text = page.extract_text() or ""
                for line in text.split("\n"):
                    line = line.strip()
                    if not line or not line[0].isdigit():
                        continue
                    # You can keep your old extract_line logic here if needed
                    pass

    return items

# Preview Function
@st.cache_data
def render_preview(pdf_bytes, page_num, crop, zoom=2.0):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page = doc[page_num - 1]
        if crop and len(crop) == 4:
            x1 = page.rect.width * crop[0] / 100
            y1 = page.rect.height * crop[1] / 100
            x2 = page.rect.width * crop[2] / 100
            y2 = page.rect.height * crop[3] / 100
            shape = page.new_shape()
            shape.draw_rect(fitz.Rect(x1, y1, x2, y2))
            shape.finish(color=(1, 0, 0), fill=(1, 0, 0), fill_opacity=0.1)
            shape.commit()
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        return pix.tobytes("png")
    except:
        return None

def create_excel(df):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ"

    yellow = PatternFill("solid", fgColor="FFFF00")
    border = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))

    for c, header in enumerate(df.columns, 1):
        cell = ws.cell(1, c, header)
        cell.fill = yellow
        cell.font = Font(bold=True)
        cell.border = border

    for r, row in enumerate(df.values, 2):
        for c, val in enumerate(row, 1):
            cell = ws.cell(r, c, val)
            cell.border = border

    ws.freeze_panes = "A2"
    wb.save(output)
    return output.getvalue()

# ====================== STREAMLIT UI ======================
if "target_bytes" not in st.session_state:
    st.session_state.target_bytes = None
if "data" not in st.session_state:
    st.session_state.data = None

st.title("📋 BOQ Extractor Pro v32 - Auto Grid Detection")

# Sidebar
with st.sidebar:
    st.subheader("Extraction Options")
    use_auto_grid = st.checkbox("✅ Enable Auto Column Detection (Recommended when no lines)", value=True)
    max_items = st.slider("Maximum Items", 10, 500, 100)

# Main Area
left, right = st.columns([1.1, 2])

with left:
    st.subheader("Crop Area")
    x = st.slider("Left (%)", 0, 100, 5)
    y = st.slider("Top (%)", 0, 100, 12)
    z = st.slider("Right (%)", 0, 100, 95)
    w = st.slider("Bottom (%)", 0, 100, 70)
    crop = (x, y, z, w)

    if st.session_state.target_bytes:
        preview = render_preview(st.session_state.target_bytes, 1, crop)
        if preview:
            st.image(preview, use_column_width=True)

with right:
    uploaded_file = st.file_uploader("Upload BOQ PDF", type="pdf")
    if uploaded_file:
        st.session_state.target_bytes = uploaded_file.read()

    if st.session_state.target_bytes:
        if st.button("🚀 EXTRACT BOQ NOW", type="primary", use_container_width=True):
            with st.spinner("Detecting table and extracting data..."):
                data_list = auto_extract_boq(
                    st.session_state.target_bytes, 
                    crop, 
                    max_items, 
                    use_auto_grid
                )
                
                if data_list:
                    st.session_state.data = pd.DataFrame(data_list)
                    st.success(f"✅ Successfully extracted {len(data_list)} items using {'Auto Grid' if use_auto_grid else 'Line'} mode")
                else:
                    st.warning("No data extracted. Try toggling Auto Column Detection or adjust crop.")

    if st.session_state.data is not None and not st.session_state.data.empty:
        df = st.session_state.data
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic", hide_index=True)

        st.divider()
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("📥 Download Excel", create_excel(edited_df), f"BOQ_{ts}.xlsx", use_container_width=True)
        with c2:
            st.download_button("📄 Download CSV", edited_df.to_csv(index=False).encode(), f"BOQ_{ts}.csv", "text/csv", use_container_width=True)
