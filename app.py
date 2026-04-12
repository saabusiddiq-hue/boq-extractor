import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
import pdfplumber
import fitz
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="BOQ Extractor Pro v34", page_icon="📋", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #0d1117; }
    .left-panel { background: #161b22; border: 2px solid #30363d; border-radius: 8px; padding: 1rem; }
    .right-panel { background: #161b22; border: 1px solid #30363d; border-radius: 8px; padding: 1rem; min-height: 700px; }
    .success-box { background: rgba(35, 134, 54, 0.1); border: 1px solid #238636; border-radius: 6px; padding: 0.75rem; }
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
            (r"Graphite", "Graphite"),
            (r"Bronze", "Bronze"),
            (r"PTFE", "PTFE"),
        ]

    def separate(self, text: str):
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

def extract_no_line_boq(pdf_bytes: bytes, crop: tuple, max_items: int):
    items = []
    sep = MaterialSeparator()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            if crop:
                w, h = page.width, page.height
                page = page.crop((w*crop[0]/100, h*crop[1]/100, w*crop[2]/100, h*crop[3]/100))

            # Extract words with positions
            words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)

            # Group into lines by y-position
            line_groups = {}
            for w in words:
                y = round(w['top'])
                if y not in line_groups:
                    line_groups[y] = []
                line_groups[y].append(w)

            for y in sorted(line_groups.keys()):
                line_words = sorted(line_groups[y], key=lambda x: x['x0'])
                full_line = " ".join(w['text'] for w in line_words).strip()

                if not full_line or not full_line[0].isdigit():
                    continue

                # Split roughly by multiple spaces or try to guess columns
                # This is the key part for no-line tables
                parts = re.split(r'\s{3,}', full_line)   # split on 3 or more spaces

                try:
                    item_part = parts[0].strip()
                    item_no = int(re.search(r'\d+', item_part).group()) if re.search(r'\d+', item_part) else 0

                    if item_no == 0 or item_no > max_items:
                        continue

                    qty = 1
                    fig_no = ""
                    remaining = ""

                    if len(parts) > 1:
                        qty_str = parts[1].strip()
                        if qty_str.isdigit():
                            qty = int(qty_str)
                            fig_no = parts[2].strip() if len(parts) > 2 else ""
                            remaining = " ".join(parts[3:]) if len(parts) > 3 else ""
                        else:
                            fig_no = parts[1].strip()
                            remaining = " ".join(parts[2:]) if len(parts) > 2 else ""

                    description, material = sep.separate(remaining)

                    items.append({
                        "Item": item_no,
                        "Qty": qty,
                        "Fig No": fig_no,
                        "Description": description,
                        "Material": material,
                        "Page": page_num,
                        "Raw": full_line[:100]   # debug only
                    })
                except:
                    continue
    return items

# Preview function
@st.cache_data
def render_preview(pdf_bytes, page_num=1, crop=None, zoom=2.0):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page = doc[page_num-1]
        if crop and len(crop) == 4:
            x1 = page.rect.width * crop[0]/100
            y1 = page.rect.height * crop[1]/100
            x2 = page.rect.width * crop[2]/100
            y2 = page.rect.height * crop[3]/100
            shape = page.new_shape()
            shape.draw_rect(fitz.Rect(x1, y1, x2, y2))
            shape.finish(color=(1,0,0), fill=(1,0,0), fill_opacity=0.12)
            shape.commit()
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        return pix.tobytes("png")
    except Exception:
        return None

def create_excel(df):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ"

    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    border = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))

    for col_num, header in enumerate(df.columns, 1):
        cell = ws.cell(1, col_num, header)
        cell.fill = yellow
        cell.font = Font(bold=True)
        cell.border = border

    for r_idx, row in enumerate(df.values, 2):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(r_idx, c_idx, value)
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

st.title("📋 BOQ Extractor Pro v34 - Simplified No-Line Mode")

with st.sidebar:
    st.subheader("Settings")
    max_items = st.slider("Max Items to Extract", 10, 500, 100)
    show_raw = st.checkbox("Show Raw Text (Debug)", value=True)

left, right = st.columns([1.1, 2.9])

with left:
    st.subheader("Crop Table Area Precisely")
    x = st.slider("Left %", 0, 100, 8)
    y = st.slider("Top %", 0, 100, 15)
    z = st.slider("Right %", 0, 100, 95)
    w = st.slider("Bottom %", 0, 100, 72)
    crop = (x, y, z, w)

    if st.session_state.target_bytes:
        preview_img = render_preview(st.session_state.target_bytes, 1, crop, 2.0)
        if preview_img:
            st.image(preview_img, use_column_width=True)

with right:
    uploaded = st.file_uploader("Upload BOQ PDF", type=["pdf"])
    if uploaded:
        st.session_state.target_bytes = uploaded.read()
        st.success("PDF uploaded")

    if st.session_state.target_bytes:
        if st.button("🚀 EXTRACT BOQ", type="primary", use_container_width=True):
            with st.spinner("Extracting using position-based method..."):
                extracted = extract_no_line_boq(st.session_state.target_bytes, crop, max_items)
                if extracted:
                    st.session_state.data = pd.DataFrame(extracted)
                    st.success(f"✅ Extracted {len(extracted)} rows")
                else:
                    st.error("No rows extracted. Please make the crop tighter around the table or try a different page.")

    if st.session_state.data is not None and not st.session_state.data.empty:
        df_display = st.session_state.data.copy()
        if not show_raw:
            df_display = df_display.drop(columns=['Raw'], errors='ignore')

        edited_df = st.data_editor(df_display, use_container_width=True, num_rows="dynamic", hide_index=True)

        st.divider()
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("📥 Download Excel", create_excel(edited_df), f"BOQ_{ts}.xlsx", use_container_width=True)
        with col2:
            st.download_button("📄 Download CSV", edited_df.to_csv(index=False).encode("utf-8"), f"BOQ_{ts}.csv", "text/csv", use_container_width=True)
