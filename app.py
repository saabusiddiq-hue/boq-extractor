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
    .warning-box { background: rgba(248, 81, 73, 0.1); border: 1px solid #f85149; border-radius: 6px; padding: 0.75rem; margin: 0.5rem 0; }
    .success-box { background: rgba(35, 134, 54, 0.1); border: 1px solid #238636; border-radius: 6px; padding: 0.75rem; margin: 0.5rem 0; }
</style>
""", unsafe_allow_html=True)

class MaterialSeparator:
    """Separate Material from Description - Improved for Graphite Bronze"""

    def __init__(self):
        # Material patterns - ordered by specificity (most specific first)
        self.material_patterns = [
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

            # === GRAPHITE BRONZE - Made stronger and higher priority ===
            (r"Graphite\s+Bronze", "Graphite Bronze"),
            (r"Graphite\s+bronze", "Graphite bronze"),
            (r"Bronze\s+Graphite", "Bronze Graphite"),

            (r"Gr\.?\s*\d+\.?\d*", "Grade"),
            (r"CI\.?\s*\d+", "CI"),
            (r"Cast\s+Iron", "Cast Iron"),
            (r"Stainless\s+Steel", "Stainless Steel"),
            (r"Carbon\s+Steel", "Carbon Steel"),

            # Single words - AFTER combined Graphite Bronze
            (r"Graphite", "Graphite"),
            (r"Bronze", "Bronze"),
            (r"PTFE", "PTFE"),
        ]

    def separate(self, description: str) -> Tuple[str, str]:
        """
        Separate Material from Description
        Returns: (clean_description, material)
        """
        if not description:
            return "", ""

        # Search for material patterns from the END of the string
        for pattern, mat_type in self.material_patterns:
            matches = list(re.finditer(pattern, description, re.IGNORECASE))
            if matches:
                last_match = matches[-1]

                material = last_match.group(0)

                clean_desc = description[:last_match.start()].strip(" -:/\t")
                clean_desc = re.sub(r"\s+Per\s*$", "", clean_desc, flags=re.IGNORECASE)
                clean_desc = re.sub(r"\s*[-:]\s*$", "", clean_desc)   # Extra cleaning
                clean_desc = clean_desc.strip()

                return clean_desc, material

        # No material found
        return description.strip(), ""

class BOQExtractor:
    """Extract BOQ with proper column separation"""

    def __init__(self):
        self.material_sep = MaterialSeparator()

    def extract_line(self, line: str) -> Dict:
        """Extract columns from a BOQ line"""
        parts = line.split()
        if len(parts) < 3 or not parts[0].isdigit():
            return None

        try:
            item_no = int(parts[0])

            qty = 1
            idx = 1
            if idx < len(parts) and parts[idx].isdigit():
                qty = int(parts[idx])
                idx += 1

            fig_no = ""
            if idx < len(parts):
                fig_no = parts[idx]
                idx += 1
                if idx < len(parts):
                    next_word = parts[idx]
                    if next_word.lower() in ["bronze", "plate", "steel", "pad", "sheet"]:
                        fig_no += " " + next_word
                        idx += 1

            remaining = parts[idx:] if idx < len(parts) else []
            full_text = " ".join(remaining)

            clean_desc, material = self.material_sep.separate(full_text)

            return {
                "Item": item_no,
                "Qty": qty,
                "Fig No": fig_no,
                "Description": clean_desc,
                "Material": material
            }

        except Exception:
            return None

def extract_boq_v28(pdf_bytes: bytes, crop: Tuple, max_items: int) -> List[Dict]:
    items = []
    extractor = BOQExtractor()

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

                result = extractor.extract_line(line)
                if result and result["Item"] <= max_items:
                    result["Page"] = page_num
                    items.append(result)

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
if "target_bytes" not in st.session_state:
    st.session_state.target_bytes = None
if "data" not in st.session_state:
    st.session_state.data = None

st.title("📋 BOQ Extractor Pro v29 - Improved Graphite Bronze")

# TOP BAR
st.markdown("<div class='upload-bar'>", unsafe_allow_html=True)

col1, arr, col2 = st.columns([2, 0.3, 2])

with col1:
    st.write("📤 Upload Sample (Optional)")
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

st.markdown("</div>", unsafe_allow_html=True)

# MAIN LAYOUT
left_col, right_col = st.columns([1, 2.5])

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

with right_col:
    st.markdown("<div class='right-panel'>", unsafe_allow_html=True)

    if st.session_state.target_bytes:
        if st.button("🔍 EXTRACT BOQ DATA", key="extract_btn", use_container_width=True):
            progress = st.progress(0, text="Extracting...")

            items = extract_boq_v28(
                st.session_state.target_bytes,
                crop,
                15
            )

            progress.empty()

            if items:
                st.session_state.data = pd.DataFrame(items)
                with_material = sum(1 for item in items if item.get("Material"))
                st.success(f"✅ Extracted {len(items)} items! ({with_material} with Material)")
            else:
                st.error("No items found")

    if st.session_state.data is not None and not st.session_state.data.empty:
        df = st.session_state.data

        total = len(df)
        with_mat = (df["Material"] != "").sum()
        without_mat = total - with_mat

        cols_stats = st.columns(3)
        cols_stats[0].metric("Total", total)
        cols_stats[1].metric("With Material", with_mat)
        cols_stats[2].metric("No Material", without_mat)

        if with_mat > 0:
            st.markdown("<div class='success-box'>✅ Material successfully separated!</div>", unsafe_allow_html=True)

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
