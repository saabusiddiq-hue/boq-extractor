"""
BOQ Extractor Pro v22.0 - CLEAN INFRASTRUCTURE
Layout: Preview(left) + Controls(top) + Table(bottom)
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import shutil
from datetime import datetime

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

from PIL import Image, ImageEnhance, ImageFilter

try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("❌ openpyxl not found")
    st.stop()

st.set_page_config(page_title="BOQ Extractor Pro", page_icon="📋", layout="wide")

# Custom CSS matching the sketch
st.markdown("""
<style>
    /* Main layout */
    .stApp { background-color: #0d1117; }

    /* Header bar */
    .top-bar {
        background: #161b22;
        border: 1px solid #30363d;
        border-radius: 8px;
        padding: 1rem;
        margin-bottom: 1rem;
        display: flex;
        align-items: center;
        gap: 1rem;
    }

    .upload-flow {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        background: #21262d;
        padding: 0.5rem 1rem;
        border-radius: 6px;
        border: 1px solid #30363d;
    }

    .arrow {
        color: #58a6ff;
        font-size: 1.2rem;
    }

    /* Preview box */
    .preview-box {
        background: #161b22;
        border: 2px solid #30363d;
        border-radius: 8px;
        padding: 1rem;
        height: 400px;
        display: flex;
        flex-direction: column;
    }

    .preview-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 0.5rem;
        padding-bottom: 0.5rem;
        border-bottom: 1px solid #30363d;
    }

    .icon-btn {
        background: none;
        border: none;
        color: #8b949e;
        cursor: pointer;
        padding: 0.25rem;
    }

    .icon-btn:hover {
        color: #58a6ff;
    }

    /* Sliders */
    .slider-container {
        background: #161b22;
        border: 1px solid #30363d;
        border-radius: 8px;
        padding: 1rem;
        margin-top: 1rem;
    }

    .slider-label {
        color: #8b949e;
        font-size: 0.8rem;
        margin-bottom: 0.25rem;
    }

    /* Analyze button */
    .analyze-btn {
        background: linear-gradient(90deg, #238636 0%, #2ea043 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        border-radius: 6px;
        font-weight: 600;
        cursor: pointer;
        width: 100%;
        margin: 1rem 0;
    }

    .analyze-btn:hover {
        opacity: 0.9;
    }

    /* Progress bar */
    .progress-container {
        background: #21262d;
        border: 1px solid #30363d;
        border-radius: 6px;
        height: 30px;
        margin: 1rem 0;
        overflow: hidden;
    }

    .progress-bar {
        height: 100%;
        background: linear-gradient(90deg, #1f6feb 0%, #58a6ff 100%);
        transition: width 0.3s ease;
        display: flex;
        align-items: center;
        justify-content: center;
        color: white;
        font-size: 0.8rem;
        font-weight: 600;
    }

    /* Table area */
    .table-container {
        background: #161b22;
        border: 1px solid #30363d;
        border-radius: 8px;
        padding: 1rem;
        margin-top: 1rem;
        min-height: 300px;
    }

    /* Material highlight */
    .material-tag {
        background: rgba(31, 111, 235, 0.2);
        color: #58a6ff;
        padding: 2px 8px;
        border-radius: 4px;
        font-size: 0.85rem;
        font-weight: 500;
    }

    /* Streamlit overrides */
    .stButton > button {
        width: 100%;
        border-radius: 6px;
        background: #238636;
        color: white;
        border: none;
        padding: 0.5rem;
    }

    .stButton > button:hover {
        background: #2ea043;
    }

    div[data-testid="stVerticalBlock"] > div[data-testid="stHorizontalBlock"] {
        gap: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Materials list
MATERIALS = [
    "Graphite bronze", "Graphite", "Bronze", "PTFE", "Lead",
    "SS304", "SS316", "SS316L", "Stainless Steel",
    "A36", "A105", "A193", "A194", "A240", "A516",
    "Carbon Steel", "Cast Iron", "CI"
]

def extract_material(text):
    """Extract material from description"""
    if not text:
        return text, ""

    # Check start
    for mat in sorted(MATERIALS, key=len, reverse=True):
        pattern = rf"^({re.escape(mat)})\s*[-,:/]?\s*(.*)"
        match = re.match(pattern, text, re.IGNORECASE)
        if match:
            return match.group(2).strip(), match.group(1)

    # Check end
    for mat in sorted(MATERIALS, key=len, reverse=True):
        pattern = rf"(.*?)\s*[-,:/]?\s*({re.escape(mat)})$"
        match = re.match(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip(), match.group(2)

    return text, ""

@st.cache_data
def render_preview(pdf_bytes, page_num, crop, zoom):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if page_num > len(doc):
            return None
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

        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        return pix.tobytes("png")
    except:
        return None

def extract_boq(pdf_bytes, crop, max_items, progress_bar):
    """Extract BOQ data"""
    items = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total_pages = len(pdf.pages)

        for page_num in range(1, total_pages + 1):
            progress = int((page_num / total_pages) * 100)
            progress_bar.progress(progress, text=f"Processing page {page_num}/{total_pages}")

            page = pdf.pages[page_num - 1]
            if crop:
                w, h = page.width, page.height
                page = page.crop((w*crop[0]/100, h*crop[1]/100, w*crop[2]/100, h*crop[3]/100))

            text = page.extract_text() or ""
            lines = text.split("\n")

            for line in lines:
                line = line.strip()
                if not line or not line[0].isdigit():
                    continue

                parts = line.split()
                if len(parts) < 3:
                    continue

                try:
                    item_no = int(parts[0])
                    if item_no > max_items:
                        continue

                    qty = 1
                    remaining = parts[1:]
                    if remaining and remaining[0].isdigit():
                        qty = int(remaining[0])
                        remaining = remaining[1:]

                    # Find part number
                    part_no = ""
                    desc_start = 0
                    for i, part in enumerate(remaining[:4]):
                        if re.match(r"^[A-Z0-9][-A-Z0-9x]+", part, re.IGNORECASE):
                            part_no = part
                            desc_start = i + 1
                            break

                    raw_desc = " ".join(remaining[desc_start:]) if desc_start < len(remaining) else ""
                    clean_desc, material = extract_material(raw_desc)

                    items.append({
                        "Item": item_no,
                        "Qty": qty,
                        "Part No": part_no,
                        "Description": clean_desc,
                        "Material": material,
                        "Page": page_num
                    })
                except:
                    continue

    return items

def create_excel(df, yellow_header=True):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ"

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_font = Font(bold=True, size=10)
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                   top=Side(style="thin"), bottom=Side(style="thin"))

    cols = ["Item", "Qty", "Part No", "Description", "Material", "Page"]
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

# Initialize session
if "sample_file" not in st.session_state:
    st.session_state.sample_file = None
if "target_file" not in st.session_state:
    st.session_state.target_file = None
if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = None

# ==================== UI LAYOUT ====================

# TOP BAR: Upload Flow
st.markdown("<div style='background:#161b22;border:1px solid #30363d;border-radius:8px;padding:1rem;margin-bottom:1rem;'>", unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns([2, 0.5, 2, 3])

with col1:
    st.markdown("📤 **Upload Sample**")
    sample = st.file_uploader("Sample (JPG/PNG/PDF)", type=["jpg", "jpeg", "png", "pdf"], label_visibility="collapsed")
    if sample:
        st.session_state.sample_file = sample.read()
        st.success("✓ Sample")

with col2:
    st.markdown("<div style='text-align:center;padding-top:2rem;'><span style='color:#58a6ff;font-size:1.5rem;'>→</span></div>", unsafe_allow_html=True)

with col3:
    st.markdown("📄 **Upload Target PDF**")
    target = st.file_uploader("Target PDF", type=["pdf"], label_visibility="collapsed")
    if target:
        st.session_state.target_file = target.read()
        st.success("✓ Target")

with col4:
    if st.session_state.sample_file and OCR_AVAILABLE:
        if st.button("🔍 Analyze Sample Image", use_container_width=True):
            with st.spinner("Analyzing..."):
                try:
                    image = Image.open(io.BytesIO(st.session_state.sample_file))
                    text = pytesseract.image_to_string(image)
                    st.session_state.sample_text = text
                    st.success("✅ Sample analyzed")
                except:
                    st.error("Analysis failed")

st.markdown("</div>", unsafe_allow_html=True)

# MAIN LAYOUT: Left (Preview) + Right (Controls & Table)
left_col, right_col = st.columns([1, 2.5])

# LEFT COLUMN: Preview & Settings
with left_col:
    st.markdown("<div class='preview-box'>", unsafe_allow_html=True)

    # Preview header with icons
    st.markdown("""
    <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:0.5rem;padding-bottom:0.5rem;border-bottom:1px solid #30363d;'>
        <span style='color:#8b949e;font-size:0.9rem;'>👁️ Preview</span>
        <div>
            <span style='color:#8b949e;cursor:pointer;margin-left:0.5rem;'>⚙️</span>
            <span style='color:#8b949e;cursor:pointer;margin-left:0.5rem;'>🏠</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Preview image
    if st.session_state.target_file:
        x = st.slider("X (Left)", 0, 100, 5, key="x_slider")
        y = st.slider("Y (Top)", 0, 100, 15, key="y_slider")
        z = st.slider("Z (Right)", 0, 100, 95, key="z_slider")
        w = st.slider("W (Bottom)", 0, 100, 60, key="w_slider")

        if z <= x: z = min(x + 10, 100)
        if w <= y: w = min(y + 10, 100)

        crop = (x, y, z, w)
        zoom = 2.0

        img = render_preview(st.session_state.target_file, 1, crop, zoom)
        if img:
            st.image(img, use_column_width=True)
    else:
        st.markdown("<div style='flex:1;display:flex;align-items:center;justify-content:center;color:#8b949e;'>No PDF loaded</div>", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

# RIGHT COLUMN: Analyze, Progress, Table
with right_col:
    # Analyze button
    if st.session_state.target_file:
        if st.button("🔍 ANALYZE SAMPLE IMAGE", use_container_width=True, type="primary"):
            progress_bar = st.progress(0, text="Starting...")

            with st.spinner("Extracting BOQ data..."):
                items = extract_boq(
                    st.session_state.target_file,
                    (st.session_state.x_slider, st.session_state.y_slider, 
                     st.session_state.z_slider, st.session_state.w_slider),
                    15,
                    progress_bar
                )

                if items:
                    st.session_state.extracted_data = pd.DataFrame(items)
                    progress_bar.empty()
                    st.success(f"✅ Extracted {len(items)} items")
                else:
                    progress_bar.empty()
                    st.error("No items found")

    # Progress bar (shown during extraction)
    # st.markdown("<div class='progress-container'><div class='progress-bar' style='width: 60%;'>60%</div></div>", unsafe_allow_html=True)

    # Table area
    st.markdown("<div class='table-container'>", unsafe_allow_html=True)

    if st.session_state.extracted_data is not None:
        df = st.session_state.extracted_data

        # Editable table
        edited_df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Item": st.column_config.NumberColumn("Item", width="small"),
                "Qty": st.column_config.NumberColumn("Qty", width="small"),
                "Part No": st.column_config.TextColumn("Part No", width="medium"),
                "Description": st.column_config.TextColumn("Description", width="large"),
                "Material": st.column_config.TextColumn("Material", width="medium"),
                "Page": st.column_config.NumberColumn("Page", width="small")
            }
        )

        st.session_state.extracted_data = edited_df

        # Export buttons
        col1, col2 = st.columns(2)
        with col1:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel = create_excel(edited_df, True)
            st.download_button(
                "📥 Download Excel",
                excel,
                f"BOQ_{ts}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col2:
            csv = edited_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📄 Download CSV",
                csv,
                f"BOQ_{ts}.csv",
                "text/csv",
                use_container_width=True
            )
    else:
        st.markdown("<div style='text-align:center;color:#8b949e;padding:3rem;'>Extracted data will appear here</div>", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)
