"""
BOQ Extractor Pro v20.0 - CLEAN EDITION
Simple material extraction with basic review
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import shutil
from datetime import datetime
from typing import Optional, Tuple, List, Dict, Any

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

try:
    import pdfplumber
except ImportError:
    st.error("❌ pdfplumber not found")
    st.stop()

try:
    import fitz
except ImportError:
    st.error("❌ PyMuPDF not found")
    st.stop()

from PIL import Image

try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("❌ openpyxl not found")
    st.stop()

st.set_page_config(page_title="BOQ Extractor Pro", page_icon="📋", layout="wide")

# CSS
st.markdown("""
<style>
    .stApp { background-color: #0d1117; }
    .section-header { 
        background: #21262d; 
        padding: 0.75rem 1rem; 
        border-radius: 8px; 
        margin: 1rem 0 0.5rem 0; 
        border-left: 3px solid #1f6feb; 
        color: #58a6ff !important; 
        font-weight: 600;
    }
    .material-box {
        background: #161b22;
        border: 1px solid #30363d;
        border-radius: 6px;
        padding: 1rem;
        margin: 0.5rem 0;
    }
    .highlight {
        background: rgba(248, 81, 73, 0.2);
        color: #f85149;
        padding: 2px 6px;
        border-radius: 4px;
    }
</style>
""", unsafe_allow_html=True)

# Material keywords
MATERIALS = [
    "Graphite bronze", "Graphite", "Bronze", "PTFE", 
    "SS304", "SS316", "SS316L", "Stainless Steel",
    "A36", "A105", "A193", "A194", "A240", "A516",
    "Carbon Steel", "Cast Iron", "CI", "MSS-SP"
]

def extract_material(text: str) -> Tuple[str, str]:
    """Extract material from text, return (clean_desc, material)"""
    if not text:
        return "", ""

    text_clean = text.strip()
    found_material = ""

    # Check for materials at start (the problem case)
    for mat in sorted(MATERIALS, key=len, reverse=True):  # Longest first
        pattern = rf"^({re.escape(mat)})\s*[-,:/]?\s*(.*)"
        match = re.match(pattern, text_clean, re.IGNORECASE)
        if match:
            found_material = match.group(1)
            text_clean = match.group(2).strip()
            break

    # If no material at start, check at end
    if not found_material:
        for mat in sorted(MATERIALS, key=len, reverse=True):
            pattern = rf"(.*?)\s*[-,:/]?\s*({re.escape(mat)})$"
            match = re.match(pattern, text_clean, re.IGNORECASE)
            if match:
                found_material = match.group(2)
                text_clean = match.group(1).strip()
                break

    return text_clean, found_material

class SimpleDetector:
    def __init__(self, max_items=15):
        self.max_items = max_items

    def parse(self, text: str) -> List[Dict]:
        items = []
        for line in text.strip().split("\n"):
            line = line.strip()
            if not line or line.upper().startswith(("ITEM", "NO.", "QTY")):
                continue

            parts = line.split()
            if len(parts) < 3 or not parts[0].isdigit():
                continue

            item_no = int(parts[0])
            if item_no > self.max_items:
                break

            # Find qty
            qty = 1
            remaining = parts[1:]
            if remaining and remaining[0].isdigit():
                qty = int(remaining[0])
                remaining = remaining[1:]

            # Find part no (pattern like V-25-BM1, 3x240x240, etc)
            part_no = ""
            desc_start = 0
            for i, part in enumerate(remaining[:4]):
                if re.match(r"^[A-Z0-9][-A-Z0-9x]+$", part, re.IGNORECASE):
                    part_no = part
                    desc_start = i + 1
                    break

            # Rest is description (may contain material)
            raw_desc = " ".join(remaining[desc_start:]) if desc_start < len(remaining) else ""

            # Extract material from description
            clean_desc, material = extract_material(raw_desc)

            items.append({
                "Item No": item_no,
                "Quantity": qty,
                "Part No": part_no,
                "Description": clean_desc,
                "Material": material,
                "_raw": raw_desc  # For review
            })

        return items

def render_pdf(pdf_bytes: bytes, page_num: int, crop: Optional[Tuple], zoom: float = 2.0):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if page_num > len(doc):
            return None, "Invalid page"

        page = doc[page_num - 1]

        if crop:
            x1, y1, x2, y2 = crop
            rect = page.rect
            x1, y1 = rect.width * x1/100, rect.height * y1/100
            x2, y2 = rect.width * x2/100, rect.height * y2/100

            shape = page.new_shape()
            shape.draw_rect(fitz.Rect(x1, y1, x2, y2))
            shape.finish(color=(0.97, 0.32, 0.29), fill=(0.97, 0.32, 0.29), fill_opacity=0.15, width=3)
            shape.commit()
            page.insert_text(fitz.Point(x1 + 5, max(y1 - 5, 10)), "EXTRACTION AREA", 
                           fontsize=14, color=(0.97, 0.32, 0.29))

        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        return pix.tobytes("png"), None
    except Exception as e:
        return None, str(e)

def extract_text(pdf_bytes: bytes, page_num: int, crop: Optional[Tuple]):
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            page = pdf.pages[page_num - 1]
            if crop:
                w, h = page.width, page.height
                page = page.crop((w*crop[0]/100, h*crop[1]/100, w*crop[2]/100, h*crop[3]/100))
            return page.extract_text() or ""
    except:
        return ""

def process_pdf(pdf_bytes: bytes, crop: Optional[Tuple], detector: SimpleDetector):
    items = []
    logs = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num in range(1, len(pdf.pages) + 1):
            page = pdf.pages[page_num - 1]
            if crop:
                w, h = page.width, page.height
                page = page.crop((w*crop[0]/100, h*crop[1]/100, w*crop[2]/100, h*crop[3]/100))

            text = page.extract_text() or ""
            page_items = detector.parse(text)
            for item in page_items:
                item["Page"] = page_num
            items.extend(page_items)
            logs.append(f"Page {page_num}: {len(page_items)} items")

    return items, logs

def create_excel(df: pd.DataFrame, yellow_header: bool = True):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ"

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_font = Font(bold=True, size=10)
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    cols = ["Drawing No", "Mark No", "Item No", "Quantity", "Part No", "Description", "Material", "Page"]
    cols = [c for c in cols if c in df.columns] + [c for c in df.columns if c not in cols and not c.startswith("_")]
    df = df[cols]

    for col_num, header in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        if yellow_header:
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
        max_len = max([len(str(ws.cell(row=r, column=idx).value or "")) for r in range(1, len(df)+2)])
        ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

    ws.freeze_panes = "A2"
    wb.save(output)
    return output.getvalue()

# Initialize
if "data" not in st.session_state:
    st.session_state.data = None
if "reviewed" not in st.session_state:
    st.session_state.reviewed = False

st.title("📋 BOQ Extractor Pro v20.0")
st.caption("Clean Edition - Material Review Built-in")

# Sidebar
with st.sidebar:
    st.header("⚙️ Settings")
    max_items = st.number_input("Max Items/Page", 1, 50, 15)

    st.divider()
    st.write("📐 Crop Region")
    x1 = st.slider("Left %", 0, 100, 5)
    y1 = st.slider("Top %", 0, 100, 15)
    x2 = st.slider("Right %", 0, 100, 95)
    y2 = st.slider("Bottom %", 0, 100, 60)
    if x2 <= x1: x2 = min(x1 + 10, 100)
    if y2 <= y1: y2 = min(y1 + 10, 100)
    crop = (x1, y1, x2, y2)

    zoom = st.select_slider("Zoom", [1.0, 1.5, 2.0, 2.5, 3.0], 2.0, format_func=lambda x: f"{x}x")
    yellow_header = st.checkbox("Yellow Headers", True)

    if st.button("🗑️ Clear", type="secondary"):
        st.session_state.data = None
        st.session_state.reviewed = False
        st.rerun()

# Main
uploaded = st.file_uploader("📄 Upload PDF", type=["pdf"])

if uploaded:
    pdf_bytes = uploaded.read()
    st.success(f"✓ {uploaded.name}")

    # Preview
    img, err = render_pdf(pdf_bytes, 1, crop, zoom)
    if img:
        st.image(img, caption="Crop Preview", use_column_width=True)

    if st.button("⚡ Extract & Review Materials", type="primary", use_container_width=True):
        with st.spinner("Processing..."):
            detector = SimpleDetector(max_items)
            items, logs = process_pdf(pdf_bytes, crop, detector)

            if items:
                df = pd.DataFrame(items)
                st.session_state.data = df
                st.session_state.reviewed = False
                st.success(f"✓ Extracted {len(items)} items")
            else:
                st.error("No items found")

# Material Review Section
if st.session_state.data is not None and not st.session_state.reviewed:
    df = st.session_state.data

    st.markdown("<div class='section-header'>🔍 Review Materials</div>", unsafe_allow_html=True)
    st.info("Check if materials were extracted correctly. Edit if 'Graphite bronze' merged with description.")

    # Show items with material issues
    edited_data = []

    for idx, row in df.iterrows():
        has_material = bool(row.get("Material"))
        raw_text = row.get("_raw", "")

        with st.container():
            cols = st.columns([0.5, 1, 2, 2, 0.5])

            with cols[0]:
                st.write(f"**#{int(row['Item No'])}**")
                if not has_material and raw_text:
                    st.markdown("<span class='highlight'>Check</span>", unsafe_allow_html=True)

            with cols[1]:
                st.write(row.get("Part No", "-"))

            with cols[2]:
                # Description editable
                current_desc = row.get("Description", "")
                new_desc = st.text_input("Desc", value=current_desc, key=f"desc_{idx}", label_visibility="collapsed")

            with cols[3]:
                # Material editable with suggestion
                current_mat = row.get("Material", "")

                # If no material but raw text exists, try to extract
                if not current_mat and raw_text:
                    _, suggested = extract_material(raw_text)
                    if suggested:
                        current_mat = suggested

                new_mat = st.text_input("Material", value=current_mat, key=f"mat_{idx}", label_visibility="collapsed")

                # Show raw text hint
                if raw_text and not row.get("Material"):
                    st.caption(f"Raw: {raw_text[:30]}...")

            with cols[4]:
                if st.button("✓", key=f"ok_{idx}"):
                    pass  # Just for visual feedback

            # Build edited row
            new_row = row.to_dict()
            new_row["Description"] = new_desc
            new_row["Material"] = new_mat
            edited_data.append(new_row)

    # Actions
    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("🤖 Auto-Fix All Materials", use_container_width=True):
            for i, row in enumerate(edited_data):
                if not row.get("Material") and row.get("_raw"):
                    _, mat = extract_material(row["_raw"])
                    if mat:
                        edited_data[i]["Material"] = mat
            st.session_state.data = pd.DataFrame(edited_data)
            st.rerun()

    with col2:
        if st.button("✅ Finalize BOQ", type="primary", use_container_width=True):
            final_df = pd.DataFrame(edited_data)
            # Remove internal columns
            final_df = final_df.drop(columns=[c for c in final_df.columns if c.startswith("_")], errors="ignore")
            st.session_state.data = final_df
            st.session_state.reviewed = True
            st.rerun()

    with col3:
        if st.button("⏭️ Skip Review", use_container_width=True):
            final_df = df.drop(columns=[c for c in df.columns if c.startswith("_")], errors="ignore")
            st.session_state.data = final_df
            st.session_state.reviewed = True
            st.rerun()

# Final Output
if st.session_state.reviewed and st.session_state.data is not None:
    df = st.session_state.data

    st.markdown("<div class='section-header'>📊 Final BOQ</div>", unsafe_allow_html=True)

    # Stats
    c1, c2, c3 = st.columns(3)
    c1.metric("Items", len(df))
    c2.metric("Pages", df["Page"].nunique() if "Page" in df.columns else 1)
    c3.metric("Materials", df["Material"].nunique() if "Material" in df.columns else 0)

    # Final edit
    final_edit = st.data_editor(df, num_rows="dynamic", use_container_width=True, hide_index=True)

    # Export
    st.markdown("<div class='section-header'>📦 Export</div>", unsafe_allow_html=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = (uploaded.name if uploaded else "BOQ").replace(".pdf", "")

    col1, col2 = st.columns(2)
    with col1:
        excel = create_excel(final_edit, yellow_header)
        st.download_button("📥 Excel", excel, f"{base}_{ts}.xlsx", 
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)
    with col2:
        csv = final_edit.to_csv(index=False).encode("utf-8")
        st.download_button("📄 CSV", csv, f"{base}_{ts}.csv", "text/csv", use_container_width=True)
