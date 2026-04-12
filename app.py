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

st.set_page_config(page_title="BOQ Extractor Pro v35", page_icon="📋", layout="wide")

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
            (r"Graphite", "Graphite"),
            (r"Bronze", "Bronze"),
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

def extract_boq_v35(pdf_bytes: bytes, crop: tuple, max_items: int, use_text_strategy: bool):
    items = []
    sep = MaterialSeparator()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            if crop:
                w, h = page.width, page.height
                page = page.crop((w*crop[0]/100, h*crop[1]/100, w*crop[2]/100, h*crop[3]/100))

            table_settings = {
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
                "snap_tolerance": 4,
                "join_tolerance": 4,
                "edge_min_length": 0,
            } if use_text_strategy else {}

            tables = page.extract_tables(table_settings=table_settings)

            for table in tables:
                if not table or len(table) < 2:
                    continue
                for row in table[1:]:  # skip header
                    if len(row) < 3:
                        continue
                    try:
                        item_str = str(row[0]).strip()
                        if not re.search(r'^\d+', item_str):
                            continue
                        item_no = int(re.search(r'^\d+', item_str).group())

                        if item_no > max_items:
                            continue

                        qty = int(str(row[1]).strip() or 1)
                        fig_no = str(row[2]).strip() if len(row) > 2 else ""
                        desc_raw = str(row[3]).strip() if len(row) > 3 else ""
                        mat_raw = str(row[4]).strip() if len(row) > 4 else ""

                        if mat_raw and mat_raw.lower() not in ["", "nan", "-", "none"]:
                            material = mat_raw
                            description = desc_raw
                        else:
                            description, material = sep.separate(desc_raw)

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

# Preview
@st.cache_data
def render_preview(pdf_bytes, crop=None, zoom=2.0):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page = doc[0]
        if crop and len(crop) == 4:
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
    border = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))

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

st.title("📋 BOQ Extractor Pro v35 - No Lines Table Mode")

with st.sidebar:
    st.subheader("Settings")
    use_text_strategy = st.checkbox("Use Text Strategy (Best for No Lines)", value=True)
    max_items = st.slider("Max Items", 10, 500, 100)

left, right = st.columns([1.2, 2.8])

with left:
    st.subheader("Crop the Table Area Tightly")
    x = st.slider("Left %", 0, 100, 5)
    y = st.slider("Top %", 0, 100, 15)
    z = st.slider("Right %", 0, 100, 95)
    w = st.slider("Bottom %", 0, 100, 70)
    crop = (x, y, z, w)

    if st.session_state.target_bytes:
        img = render_preview(st.session_state.target_bytes, crop)
        if img:
            st.image(img, use_column_width=True)

with right:
    uploaded = st.file_uploader("Upload BOQ PDF", type=["pdf"])
    if uploaded:
        st.session_state.target_bytes = uploaded.read()
        st.success("PDF Loaded")

    if st.session_state.target_bytes:
        if st.button("🚀 EXTRACT BOQ", type="primary", use_container_width=True):
            with st.spinner("Trying text-based table extraction..."):
                data = extract_boq_v35(st.session_state.target_bytes, crop, max_items, use_text_strategy)
                if data:
                    st.session_state.data = pd.DataFrame(data)
                    st.success(f"✅ Extracted {len(data)} items!")
                else:
                    st.warning("No items found. Try making crop tighter or toggle Text Strategy OFF.")

    if st.session_state.data is not None and not st.session_state.data.empty:
        edited = st.data_editor(st.session_state.data, use_container_width=True, num_rows="dynamic", hide_index=True)

        st.divider()
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("📥 Excel", create_excel(edited), f"BOQ_{ts}.xlsx", use_container_width=True)
        with c2:
            st.download_button("📄 CSV", edited.to_csv(index=False).encode(), f"BOQ_{ts}.csv", "text/csv", use_container_width=True)
