import streamlit as st
import pandas as pd
import io
import re
import shutil
from datetime import datetime
from typing import List, Dict, Tuple

import pdfplumber
import fitz
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="BOQ Extractor Pro v33", page_icon="📋", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #0d1117; }
    .left-panel { background: #161b22; border: 2px solid #30363d; border-radius: 8px; padding: 1rem; }
    .right-panel { background: #161b22; border: 1px solid #30363d; border-radius: 8px; padding: 1rem; min-height: 700px; }
    .success-box { background: rgba(35, 134, 54, 0.1); border: 1px solid #238636; border-radius: 6px; padding: 0.75rem; }
    .warning-box { background: rgba(248, 81, 73, 0.1); border: 1px solid #f85149; border-radius: 6px; padding: 0.75rem; }
</style>
""", unsafe_allow_html=True)

class MaterialSeparator:
    def __init__(self):
        self.patterns = [
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
                desc = re.sub(r"\s*[-:]\s*$", "", desc).strip()
                return desc, material
        return text.strip(), ""

def extract_no_line_table(pdf_bytes: bytes, crop: tuple, max_items: int):
    items = []
    separator = MaterialSeparator()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            if crop:
                w, h = page.width, page.height
                page = page.crop((w*crop[0]/100, h*crop[1]/100, w*crop[2]/100, h*crop[3]/100))

            # Get all text words with positions
            words = page.extract_words(x_tolerance=3, y_tolerance=3)

            # Group words into lines
            lines = {}
            for word in words:
                y = round(word['top'], 0)
                if y not in lines:
                    lines[y] = []
                lines[y].append(word)

            # Process each line
            for y_pos in sorted(lines.keys()):
                line_words = sorted(lines[y_pos], key=lambda w: w['x0'])
                line_text = " ".join(w['text'] for w in line_words)

                if not line_text or not line_text[0].isdigit():
                    continue

                # Simple split by assuming columns based on position or spaces
                parts = re.split(r'\s{2,}', line_text.strip())  # split on 2+ spaces

                try:
                    item_no = int(parts[0].split()[0])
                    if item_no > max_items:
                        continue

                    qty = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 1
                    fig_no = parts[2] if len(parts) > 2 else ""
                    
                    # Remaining text = Description + Material
                    remaining = " ".join(parts[3:]) if len(parts) > 3 else ""
                    
                    description, material = separator.separate(remaining)

                    items.append({
                        "Item": item_no,
                        "Qty": qty,
                        "Fig No": fig_no,
                        "Description": description,
                        "Material": material,
                        "Page": page_num,
                        "Raw Line": line_text   # for debugging
                    })
                except:
                    continue

    return items

# Preview
@st.cache_data
def render_preview(pdf_bytes, page_num, crop, zoom=2.0):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page = doc[page_num-1]
        if crop and len(crop)==4:
            x1 = page.rect.width * crop[0]/100
            y1 = page.rect.height * crop[1]/100
            x2 = page.rect.width * crop[2]/100
            y2 = page.rect.height * crop[3]/100
            shape = page.new_shape()
            shape.draw_rect(fitz.Rect(x1, y1, x2, y2))
            shape.finish(color=(1,0,0), fill=(1,0,0), fill_opacity=0.15)
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

    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin"))

    for col_num, header in enumerate(df.columns, 1):
        cell = ws.cell(1, col_num, header)
        cell.fill = yellow
        cell.font = Font(bold=True)
        cell.border = border

    for r, row in enumerate(df.values, 2):
        for c, val in enumerate(row, 1):
            cell = ws.cell(r, c, val)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    ws.freeze_panes = "A2"
    wb.save(output)
    return output.getvalue()

# ====================== UI ======================
if "target_bytes" not in st.session_state:
    st.session_state.target_bytes = None
if "data" not in st.session_state:
    st.session_state.data = None

st.title("📋 BOQ Extractor Pro v33 - No Lines Mode")

with st.sidebar:
    st.subheader("Options")
    show_debug = st.checkbox("Show Raw Lines (Debug)", value=False)
    max_items = st.slider("Max Items", 10, 500, 100)

left, right = st.columns([1.2, 2.8])

with left:
    st.subheader("Crop the Table Area")
    x = st.slider("Left %", 0, 100, 5)
    y = st.slider("Top %", 0, 100, 10)
    z = st.slider("Right %", 0, 100, 98)
    w = st.slider("Bottom %", 0, 100, 75)
    crop = (x, y, z, w)

    if st.session_state.target_bytes:
        img = render_preview(st.session_state.target_bytes, 1, crop, 2.0)
        if img:
            st.image(img, use_column_width=True)

with right:
    uploaded = st.file_uploader("Upload your BOQ PDF", type=["pdf"])
    if uploaded:
        st.session_state.target_bytes = uploaded.read()
        st.success("PDF Loaded")

    if st.session_state.target_bytes:
        if st.button("🚀 EXTRACT NOW (No Lines Mode)", type="primary", use_container_width=True):
            with st.spinner("Processing..."):
                data = extract_no_line_table(st.session_state.target_bytes, crop, max_items)
                if data:
                    st.session_state.data = pd.DataFrame(data)
                    st.success(f"✅ Extracted {len(data)} items")
                else:
                    st.error("Nothing extracted. Try adjusting crop area larger.")

    if st.session_state.data is not None and not st.session_state.data.empty:
        df = st.session_state.data

        if not show_debug:
            df = df.drop(columns=['Raw Line'], errors='ignore')

        edited = st.data_editor(df, use_container_width=True, num_rows="dynamic", hide_index=True)

        st.divider()
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("📥 Excel", create_excel(edited), f"BOQ_{ts}.xlsx", use_container_width=True)
        with c2:
            st.download_button("📄 CSV", edited.to_csv(index=False).encode(), f"BOQ_{ts}.csv", "text/csv", use_container_width=True)
