
import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import shutil
from datetime import datetime
from typing import List, Dict, Tuple

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
    st.error("pdfplumber not found")
    st.stop()

try:
    import fitz
except ImportError:
    st.error("PyMuPDF not found")
    st.stop()

from PIL import Image, ImageEnhance, ImageFilter

try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("openpyxl not found")
    st.stop()

st.set_page_config(page_title="BOQ Extractor Pro", page_icon="📋", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #0d1117; }
    .upload-bar { background: #161b22; border: 1px solid #30363d; border-radius: 8px; padding: 1rem; margin-bottom: 1rem; }
    .left-panel { background: #161b22; border: 2px solid #30363d; border-radius: 8px; padding: 1rem; }
    .right-panel { background: #161b22; border: 1px solid #30363d; border-radius: 8px; padding: 1rem; min-height: 600px; }
    .extract-btn { background: linear-gradient(90deg, #238636 0%, #2ea043 100%) !important; color: white !important; font-weight: 600 !important; padding: 1rem !important; font-size: 1.1rem !important; }
    .info-box { background: rgba(31, 111, 235, 0.1); border: 1px solid #1f6feb; border-radius: 6px; padding: 0.75rem; margin: 0.5rem 0; font-size: 0.85rem; }
    .fig-no-tag { background: rgba(35, 134, 54, 0.2); color: #3fb950; padding: 2px 6px; border-radius: 4px; font-weight: 500; }
</style>
""", unsafe_allow_html=True)

class BOQFormat:
    """Understand BOQ format from sample"""
    def __init__(self):
        self.columns = []
        self.fig_no_position = 2  # Default: position 2 (after Item, Qty)
        self.has_material_column = True
        self.fig_no_examples = []  # Store examples like PTFE, Graphite bronze

    def analyze(self, text: str):
        """Analyze sample text to understand format"""
        lines = text.strip().split("\n")

        # Find header
        for line in lines[:10]:
            line_upper = line.upper()
            if "ITEM" in line_upper:
                # Parse headers
                headers = [h.strip() for h in re.split(r"\s{2,}|\t", line) if h.strip()]
                self.columns = headers

                # Find FIG NO / PART NO position
                for i, h in enumerate(headers):
                    if any(x in h.upper() for x in ["FIG", "PART", "MARK"]):
                        self.fig_no_position = i
                        break

                # Check for MATERIAL column
                if any("MATERIAL" in h.upper() for h in headers):
                    self.has_material_column = True
                break

        # Collect FIG NO examples from data rows
        for line in lines[5:25]:
            if line.strip() and line[0].isdigit():
                parts = line.split()
                if len(parts) > self.fig_no_position:
                    # FIG NO can be 1 or 2 words
                    fig_candidate = parts[self.fig_no_position]

                    # Check if next word is also part of FIG NO (like "Graphite bronze")
                    if self.fig_no_position + 1 < len(parts):
                        next_word = parts[self.fig_no_position + 1]
                        if next_word.lower() in ["bronze", "plate", "steel", "pad"]:
                            fig_candidate += " " + next_word

                    if fig_candidate and fig_candidate not in self.fig_no_examples:
                        self.fig_no_examples.append(fig_candidate)

        return self

def extract_boq_v26(pdf_bytes: bytes, crop: Tuple, max_items: int, format_info: BOQFormat = None) -> List[Dict]:
    """Extract BOQ preserving FIG NO in description as intended"""
    items = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num in range(1, len(pdf.pages) + 1):
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

                    # Qty (position 1)
                    qty = 1
                    idx = 1
                    if idx < len(parts) and parts[idx].isdigit():
                        qty = int(parts[idx])
                        idx += 1

                    # FIG NO (can be 1 or 2 words)
                    fig_no = ""
                    if idx < len(parts):
                        fig_no = parts[idx]
                        idx += 1

                        # Check if next word extends FIG NO (Graphite bronze, SS Plate, etc.)
                        if idx < len(parts):
                            next_word = parts[idx]
                            # Common second words for FIG NO
                            if next_word.lower() in ["bronze", "plate", "steel", "pad", "sheet", "rod", "bush"]:
                                # Verify it's not actually the start of description
                                if len(parts) > idx + 1:  # There are more parts after
                                    fig_no += " " + next_word
                                    idx += 1

                    # Rest is Description + Material
                    remaining = parts[idx:] if idx < len(parts) else []

                    description = ""
                    material = ""

                    if remaining:
                        # Join all remaining as description initially
                        full_text = " ".join(remaining)

                        # Check if there's a material specification at the end
                        # Material patterns: A36, SS316, Per MSS-SP58, Gr. 8.8, etc.
                        material_patterns = [
                            r"(A36|A105|A193|A194|A240|A516)\b",
                            r"(SS316|SS316L|SS304|SS304L)\b",
                            r"(Per\s+MSS\-SP\d+|MSS\-SP\d+)",
                            r"(Gr\.\s*\d+\.?\d*)",
                            r"(CI\.\s*\d+|Cast\s+Iron)",
                            r"(Bronze|Graphite|PTFE|Carbon\s+Steel)"
                        ]

                        # Search from end for material
                        material_found = ""
                        for pattern in material_patterns:
                            matches = list(re.finditer(pattern, full_text, re.IGNORECASE))
                            if matches:
                                # Get the last match (closest to end)
                                last_match = matches[-1]
                                # Check if it's at the end or followed by minimal text
                                remaining_after = full_text[last_match.end():].strip()
                                if len(remaining_after) < 5:  # Material is at end
                                    material_found = last_match.group(1)
                                    # Description is everything before material
                                    description = full_text[:last_match.start()].strip(" -:/")
                                    break

                        if not material_found:
                            # No material found, everything is description
                            description = full_text

                    items.append({
                        "Item": item_no,
                        "Qty": qty,
                        "Fig No": fig_no,
                        "Description": description,
                        "Material": material,
                        "Page": page_num
                    })

                except Exception as e:
                    continue

    return items

@st.cache_data
def render_preview(pdf_bytes, page_num, crop, zoom):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if page_num > len(doc):
            return None
        page = doc[page_num - 1]

        if crop and len(crop) == 4:
            x1, y1, x2, y2 = crop
            rect = page.rect
            x1 = rect.width * x1 / 100
            y1 = rect.height * y1 / 100
            x2 = rect.width * x2 / 100
            y2 = rect.height * y2 / 100

            shape = page.new_shape()
            shape.draw_rect(fitz.Rect(x1, y1, x2, y2))
            shape.finish(color=(1, 0, 0), fill=(1, 0, 0), fill_opacity=0.1, width=2)
            shape.commit()
            page.insert_text(fitz.Point(x1 + 5, max(y1 - 5, 10)), "CROP", fontsize=10, color=(1, 0, 0))

        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        return pix.tobytes("png")
    except:
        return None

def create_excel(df, yellow_header=True):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ"

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_font = Font(bold=True, size=10)
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                   top=Side(style="thin"), bottom=Side(style="thin"))

    cols = ["Item", "Qty", "Fig No", "Description", "Material", "Page"]
    df = df[[c for c in cols if c in df.columns]]

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

# Session init
if "format" not in st.session_state:
    st.session_state.format = BOQFormat()
if "target_bytes" not in st.session_state:
    st.session_state.target_bytes = None
if "data" not in st.session_state:
    st.session_state.data = None

st.title("📋 BOQ Extractor Pro v26 - Preserve FIG NO in Description")

# TOP BAR
st.markdown("<div class='upload-bar'>", unsafe_allow_html=True)

col1, arr, col2, col3 = st.columns([2, 0.3, 2, 3])

with col1:
    st.write("📤 Upload Sample (JPG/PNG/PDF)")
    sample_file = st.file_uploader("Sample", type=["jpg", "jpeg", "png", "pdf"], label_visibility="collapsed")

with arr:
    st.write("&nbsp;")
    st.markdown("<div style='padding-top:2rem;text-align:center;color:#58a6ff;'>→</div>", unsafe_allow_html=True)

with col2:
    st.write("📄 Upload Target PDF *")
    target_file = st.file_uploader("Target", type=["pdf"], label_visibility="collapsed")
    if target_file:
        st.session_state.target_bytes = target_file.read()
        st.success("Target loaded")

with col3:
    if sample_file:
        st.write("&nbsp;")
        if st.button("🔍 Analyze Sample Format", use_container_width=True):
            with st.spinner("Learning format..."):
                try:
                    if sample_file.type == "application/pdf":
                        with pdfplumber.open(io.BytesIO(sample_file.read())) as pdf:
                            text = pdf.pages[0].extract_text() or ""
                    else:
                        img = Image.open(io.BytesIO(sample_file.read()))
                        enhancer = ImageEnhance.Contrast(img)
                        img = enhancer.enhance(2.0)
                        text = pytesseract.image_to_string(img, config="--psm 6")

                    st.session_state.format.analyze(text)
                    st.success(f"✅ Learned {len(st.session_state.format.fig_no_examples)} FIG NO patterns")
                except Exception as e:
                    st.error(f"Error: {e}")

st.markdown("</div>", unsafe_allow_html=True)

# Show learned format
if st.session_state.format.fig_no_examples:
    with st.expander("📊 Detected FIG NO Examples", expanded=False):
        st.write("Examples:", st.session_state.format.fig_no_examples[:10])
        st.info("ℹ️ FIG NO will be extracted separately, but DESCRIPTION keeps the full text including FIG NO prefix")

# MAIN LAYOUT
left_col, right_col = st.columns([1, 2.5])

# LEFT - Preview
with left_col:
    st.markdown("<div class='left-panel'>", unsafe_allow_html=True)
    st.write("👁️ Preview & Settings")

    x = st.slider("X (Left)", 0, 100, 5)
    y = st.slider("Y (Top)", 0, 100, 15)
    z = st.slider("Z (Right)", 0, 100, 95)
    w = st.slider("W (Bottom)", 0, 100, 60)

    if z <= x: z = min(x + 10, 100)
    if w <= y: w = min(y + 10, 100)

    crop = (x, y, z, w)

    if st.session_state.target_bytes:
        img = render_preview(st.session_state.target_bytes, 1, crop, 2.0)
        if img:
            st.image(img, use_column_width=True)
    else:
        st.info("Upload PDF")

    st.markdown("</div>", unsafe_allow_html=True)

# RIGHT - Table
with right_col:
    st.markdown("<div class='right-panel'>", unsafe_allow_html=True)

    if st.session_state.target_bytes:
        if st.button("🔍 EXTRACT BOQ DATA", key="extract_btn", use_container_width=True):
            progress = st.progress(0, text="Extracting...")

            items = extract_boq_v26(
                st.session_state.target_bytes,
                crop,
                15,
                st.session_state.format
            )

            progress.empty()

            if items:
                st.session_state.data = pd.DataFrame(items)
                st.success(f"✅ Extracted {len(items)} items!")
            else:
                st.error("No items found")

    # TABLE
    if st.session_state.data is not None and not st.session_state.data.empty:
        df = st.session_state.data

        # Info
        st.markdown("<div class='info-box'>FIG NO extracted separately | DESCRIPTION keeps full text | MATERIAL from end</div>", unsafe_allow_html=True)

        # Editable table
        edited_df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="boq_table",
            column_config={
                "Item": st.column_config.NumberColumn("Item", width="small"),
                "Qty": st.column_config.NumberColumn("Qty", width="small"),
                "Fig No": st.column_config.TextColumn("Fig No", width="medium"),
                "Description": st.column_config.TextColumn("Description", width="large"),
                "Material": st.column_config.TextColumn("Material", width="medium"),
                "Page": st.column_config.NumberColumn("Page", width="small")
            }
        )

        st.session_state.data = edited_df

        # Export
        st.divider()
        col_excel, col_csv = st.columns(2)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")

        with col_excel:
            excel = create_excel(edited_df, True)
            st.download_button("📥 Excel", excel, f"BOQ_{ts}.xlsx",
                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             use_container_width=True)

        with col_csv:
            csv = edited_df.to_csv(index=False).encode("utf-8")
            st.download_button("📄 CSV", csv, f"BOQ_{ts}.csv", "text/csv", use_container_width=True)
    else:
        st.info("BOQ data will appear here after extraction")

    st.markdown("</div>", unsafe_allow_html=True)
