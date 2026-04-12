
"""
BOQ Extractor Pro v12.0 - Interactive Preview Edition
Material-Anchored PDF to Excel/CSV Converter with Visual Anchor Points
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import base64
from datetime import datetime
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from PIL import Image, ImageDraw, ImageFont
import fitz  # PyMuPDF

# Page configuration
st.set_page_config(
    page_title="BOQ Extractor Pro",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for professional look
def load_css():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

        * {
            font-family: 'Inter', sans-serif;
        }

        .main-header {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            color: white;
            padding: 2rem;
            border-radius: 16px;
            text-align: center;
            margin-bottom: 2rem;
            box-shadow: 0 10px 40px rgba(30, 60, 114, 0.3);
        }

        .main-header h1 {
            font-size: 2.5rem;
            font-weight: 700;
            margin-bottom: 0.5rem;
        }

        .main-header p {
            font-size: 1.1rem;
            opacity: 0.9;
            margin-bottom: 1rem;
        }

        .badge-container {
            display: flex;
            gap: 10px;
            justify-content: center;
            flex-wrap: wrap;
        }

        .material-badge {
            background-color: rgba(255,255,255,0.2);
            color: white;
            padding: 6px 16px;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 500;
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255,255,255,0.3);
        }

        .stat-card {
            background: white;
            border-radius: 12px;
            padding: 1.5rem;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
            border-left: 4px solid #1e3c72;
        }

        .stat-value {
            font-size: 2rem;
            font-weight: 700;
            color: #1e3c72;
        }

        .stat-label {
            font-size: 0.9rem;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        .section-header {
            background: #f8f9fa;
            padding: 1rem 1.5rem;
            border-radius: 10px;
            margin: 2rem 0 1rem 0;
            border-left: 4px solid #1e3c72;
        }

        .section-header h3 {
            margin: 0;
            color: #1e3c72;
            font-weight: 600;
        }

        .upload-zone {
            border: 2px dashed #1e3c72;
            border-radius: 16px;
            padding: 3rem;
            text-align: center;
            background: #fafbfc;
            transition: all 0.3s ease;
        }

        .upload-zone:hover {
            background: #f0f4f8;
            border-color: #2a5298;
        }

        .preview-container {
            background: #1e1e1e;
            border-radius: 12px;
            padding: 1rem;
            position: relative;
        }

        .region-label {
            position: absolute;
            background: rgba(30, 60, 114, 0.9);
            color: white;
            padding: 4px 12px;
            border-radius: 4px;
            font-size: 0.8rem;
            font-weight: 600;
            z-index: 100;
        }

        .material-zone-indicator {
            display: inline-block;
            background: linear-gradient(135deg, #ff6b6b 0%, #ee5a6f 100%);
            color: white;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
            margin: 2px;
            box-shadow: 0 2px 8px rgba(238, 90, 111, 0.3);
        }

        .anchor-point {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            background: #e3f2fd;
            color: #1976d2;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
            margin: 3px;
            border: 2px solid #1976d2;
        }

        .anchor-point::before {
            content: "⚓";
        }

        .detection-result {
            background: #f3e5f5;
            border-left: 4px solid #9c27b0;
            padding: 1rem;
            border-radius: 8px;
            margin: 0.5rem 0;
        }

        .log-entry {
            font-family: 'Courier New', monospace;
            font-size: 0.85rem;
            padding: 0.5rem;
            border-bottom: 1px solid #eee;
        }

        .log-success { color: #28a745; }
        .log-error { color: #dc3545; }
        .log-info { color: #17a2b8; }

        .preview-box {
            background: #1e1e1e;
            color: #d4d4d4;
            padding: 1.5rem;
            border-radius: 12px;
            font-family: 'Courier New', monospace;
            font-size: 0.9rem;
            max-height: 400px;
            overflow-y: auto;
            white-space: pre-wrap;
        }

        .highlight-line {
            background: rgba(30, 60, 114, 0.3);
            padding: 2px 0;
            border-radius: 3px;
            border-left: 3px solid #ff6b6b;
        }

        .material-highlight {
            background: rgba(255, 107, 107, 0.3);
            color: #ff6b6b;
            padding: 2px 6px;
            border-radius: 4px;
            font-weight: bold;
        }

        .slider-container {
            background: white;
            padding: 1.5rem;
            border-radius: 12px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
            margin-bottom: 1rem;
        }

        .footer {
            text-align: center;
            padding: 2rem;
            color: #666;
            border-top: 1px solid #eee;
            margin-top: 3rem;
        }

        /* Custom scrollbar */
        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }

        ::-webkit-scrollbar-track {
            background: #f1f1f1;
            border-radius: 4px;
        }

        ::-webkit-scrollbar-thumb {
            background: #1e3c72;
            border-radius: 4px;
        }

        ::-webkit-scrollbar-thumb:hover {
            background: #2a5298;
        }

        .stSlider > div > div > div > div {
            background-color: #1e3c72 !important;
        }
    </style>
    """, unsafe_allow_html=True)

# =============================================================================
# MATERIAL-ANCHORED EXTRACTION ENGINE
# =============================================================================

class MaterialAnchoredExtractor:
    """
    Uses MATERIAL column as the anchor/pillar for parsing.
    """

    def __init__(self):
        # Material patterns with display names
        self.material_patterns = [
            (r'A240\s+SS316', 'A240 SS316'),
            (r'A36\b', 'A36'),
            (r'A105\b', 'A105'),
            (r'A193\s+GR\.B7', 'A193 GR.B7'),
            (r'A194\s+GR\.2H', 'A194 GR.2H'),
            (r'Per\s+MSS-SP[0-9]+', 'Per MSS-SP58'),
            (r'CARBON\s+STEEL', 'CARBON STEEL'),
            (r'A516\s+GR\.60', 'A516 GR.60'),
            (r'A240\b', 'A240'),
            (r'A193\b', 'A193'),
            (r'A194\b', 'A194'),
            (r'A516\b', 'A516'),
        ]

        self.part_patterns = [
            r'^[A-Z][0-9]+-[A-Z][0-9]+$',
            r'^V[0-9]+-[0-9]+-[A-Z]+[0-9]*$',
            r'^M[0-9]+x[0-9]+$',
            r'^PC[0-9]+-[0-9\-]+$',
            r'^[0-9]+x[0-9]+x[0-9]+$',
            r'^Graphite\s+bronze$',
            r'^[A-Z][0-9]+$',
        ]

    def extract_material_from_end(self, text):
        """Extract material from the END of a line"""
        text = text.strip()

        for pattern, display in self.material_patterns:
            matches = list(re.finditer(pattern, text, re.IGNORECASE))
            if matches:
                last_match = matches[-1]
                material = last_match.group(0)
                before_material = text[:last_match.start()].strip()
                return material, before_material, last_match.start(), last_match.end()

        return "", text, -1, -1

    def find_all_materials_in_text(self, text):
        """Find all material occurrences in text with positions"""
        materials_found = []

        for pattern, display in self.material_patterns:
            for match in re.finditer(pattern, text, re.IGNORECASE):
                materials_found.append({
                    'material': match.group(0),
                    'start': match.start(),
                    'end': match.end(),
                    'line': text[:match.start()].count('\n') + 1
                })

        return sorted(materials_found, key=lambda x: x['start'])

    def is_part_number(self, text):
        text = text.strip()
        return any(re.match(p, text, re.IGNORECASE) for p in self.part_patterns)

    def is_item_number(self, text):
        try:
            num = int(text.strip())
            return 1 <= num <= 100
        except:
            return False

    def is_quantity(self, text):
        try:
            num = int(text.strip())
            return 1 <= num <= 1000
        except:
            return False

    def parse_line_material_anchor(self, line):
        """Parse a BOQ line using Material as the rightmost anchor"""
        line = line.strip()
        if not line:
            return None

        skip_words = ['Item', 'No.', 'Fig.', 'Description', 'Material', "Req'd", 
                      'LOCATION', 'PLAN', 'TITLE', 'SUPPORT ASSEMBLY', 'DRAWING',
                      'SpringDetails', 'Serial No', 'NOTE:', 'Pipe NB']

        first_word = line.split()[0] if line.split() else ""
        if any(sw.startswith(first_word) for sw in skip_words) or len(line) < 20:
            return None

        parts = line.split()
        if len(parts) < 5:
            return None

        try:
            if not (self.is_item_number(parts[0]) and self.is_quantity(parts[1])):
                return None

            item_no = int(parts[0])
            qty = int(parts[1])

            material = ""
            material_end_idx = len(parts)

            for check_len in [3, 2, 1]:
                if len(parts) >= 2 + check_len:
                    potential_material = ' '.join(parts[-check_len:])
                    material_found, _, _, _ = self.extract_material_from_end(potential_material)
                    if material_found:
                        material = material_found
                        material_end_idx = len(parts) - check_len
                        break

            middle_parts = parts[2:material_end_idx]

            if not middle_parts:
                return None

            part_no = middle_parts[0]

            if not self.is_part_number(part_no):
                if len(middle_parts) > 1 and self.is_part_number(middle_parts[1]):
                    part_no = middle_parts[1]
                    description = ' '.join(middle_parts[2:])
                else:
                    return None
            else:
                description = ' '.join(middle_parts[1:])

            return {
                'Item No': item_no,
                'Quantity': qty,
                'Part No': part_no,
                'Description': description,
                'Material': material
            }

        except Exception as e:
            return None

    def extract_from_region(self, text, drawing_no="", mark_no=""):
        """Extract all items from text region"""
        items = []
        if not text:
            return items

        lines = text.split('\n')

        for line in lines:
            parsed = self.parse_line_material_anchor(line)
            if parsed:
                parsed['Drawing No'] = drawing_no
                parsed['Mark No'] = mark_no
                items.append(parsed)

        return items

# =============================================================================
# INTERACTIVE PREVIEW GENERATOR
# =============================================================================

def generate_annotated_preview(pdf_file, page_num, x_range, y_range, show_material_zones=True):
    """
    Generate an annotated preview showing:
    - Extraction region (blue box)
    - Material anchor zones (red highlights)
    - Detected material positions
    """
    try:
        pdf_file.seek(0)
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")

        if page_num > len(doc):
            doc.close()
            return None, "Page number exceeds document length"

        page = doc[page_num - 1]
        rect = page.rect
        width, height = rect.width, rect.height

        # Calculate extraction region based on sliders
        left = width * (x_range[0] / 100)
        right = width * (x_range[1] / 100)
        top = height * (y_range[0] / 100)
        bottom = height * (y_range[1] / 100)

        # Draw extraction region
        extract_rect = fitz.Rect(left, top, right, bottom)
        page.draw_rect(extract_rect, color=(0, 0.5, 1), width=3)  # Blue

        # Add region label
        label_point = fitz.Point(left + 10, top - 10)
        page.insert_text(label_point, "EXTRACTION REGION", 
                        fontsize=12, color=(0, 0.5, 1), fontname="helv")

        # Extract text to find material positions
        if show_material_zones:
            text = page.get_text()
            extractor = MaterialAnchoredExtractor()
            materials = extractor.find_all_materials_in_text(text)

            # Get text blocks with positions
            blocks = page.get_text("blocks")

            for block in blocks:
                x0, y0, x1, y1, text_content, block_no, block_type = block

                # Check if this block contains materials
                for mat_info in materials:
                    if mat_info['material'] in text_content:
                        # Draw material zone indicator
                        mat_rect = fitz.Rect(x0 - 5, y0 - 5, x1 + 5, y1 + 5)
                        page.draw_rect(mat_rect, color=(1, 0.2, 0.2), width=2)  # Red

                        # Add material label
                        label_y = max(y0 - 20, 20)
                        label_point = fitz.Point(x0, label_y)
                        page.insert_text(label_point, f"⚓ {mat_info['material']}", 
                                        fontsize=10, color=(1, 0.2, 0.2), fontname="helv")

        # Render to image
        mat = fitz.Matrix(1.2, 1.2)
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")

        doc.close()
        return img_data, None

    except Exception as e:
        return None, str(e)

def extract_text_with_preview(pdf_file, page_num, x_range, y_range):
    """Extract text from specific region and show material detection"""
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            page = pdf.pages[page_num - 1]
            width, height = page.width, page.height

            left = width * (x_range[0] / 100)
            right = width * (x_range[1] / 100)
            top = height * (y_range[0] / 100)
            bottom = height * (y_range[1] / 100)

            region = (left, top, right, bottom)
            cropped = page.crop(region)
            text = cropped.extract_text()

            # Analyze materials in extracted text
            extractor = MaterialAnchoredExtractor()
            materials = extractor.find_all_materials_in_text(text)

            return text, materials

    except Exception as e:
        return "", []

# =============================================================================
# MAIN EXTRACTION FUNCTION
# =============================================================================

def extract_with_spatial_material(pdf_file, start_page, end_page, x_range, y_range, progress_bar=None):
    """Extract with configurable spatial regions"""
    all_items = []
    logs = []

    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            total_pages = len(pdf.pages)
            actual_end = min(end_page, total_pages) if end_page != 999 else total_pages

            logs.append(f"📁 Processing {actual_end - start_page + 1} pages (Total: {total_pages})")

            extractor = MaterialAnchoredExtractor()

            for page_num in range(start_page - 1, actual_end):
                page = pdf.pages[page_num]

                if progress_bar:
                    progress = (page_num - start_page + 2) / (actual_end - start_page + 1)
                    progress_bar.progress(min(progress, 1.0), 
                                        f"Processing page {page_num + 1}/{actual_end}...")

                width, height = page.width, page.height

                # Use configurable regions
                left = width * (x_range[0] / 100)
                right = width * (x_range[1] / 100)
                top = height * (y_range[0] / 100)
                bottom = height * (y_range[1] / 100)

                table_region = (left, top, right, bottom)

                # Title block (fixed at bottom right)
                title_region = (
                    width * 0.60,
                    height * 0.82,
                    width * 0.98,
                    height * 0.98
                )

                # Extract
                table_crop = page.crop(table_region)
                table_text = table_crop.extract_text()

                title_crop = page.crop(title_region)
                title_text = title_crop.extract_text()

                # Get metadata
                drawing_no = ""
                mark_no = ""

                if title_text:
                    dwg_match = re.search(r'Q24250\s+SUPP\s+05-[0-9]+', title_text)
                    if dwg_match:
                        drawing_no = dwg_match.group(0)

                    mark_match = re.search(r'POS-[0-9]{4}-[0-9]{3}-[A-Z0-9]+-[0-9A-Z-]+', title_text)
                    if mark_match:
                        mark_no = mark_match.group(0)

                # Parse items
                items = extractor.extract_from_region(table_text, drawing_no, mark_no)

                # Count materials found
                materials_in_page = extractor.find_all_materials_in_text(table_text)

                status = "✅" if items else "⚠️"
                logs.append(f"{status} Page {page_num+1}: {len(items)} items, {len(materials_in_page)} materials | {drawing_no}")

                all_items.extend(items)

    except Exception as e:
        logs.append(f"❌ Error: {str(e)}")

    return all_items, logs

# =============================================================================
# EXCEL EXPORT
# =============================================================================

def make_excel(df, yellow_header=True):
    out = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ"

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    cols = ['Drawing No', 'Mark No', 'Item No', 'Quantity', 'Part No', 'Description', 'Material']
    available = [c for c in cols if c in df.columns]
    df_out = df[available]

    for col_num, header in enumerate(df_out.columns, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        if yellow_header:
            cell.fill = yellow_fill
        cell.font = Font(bold=True, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for row_num, row in enumerate(df_out.values, 2):
        for col_num, val in enumerate(row, 1):
            cell = ws.cell(row=row_num, column=col_num, value=val)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 28
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 22
    ws.column_dimensions['F'].width = 40
    ws.column_dimensions['G'].width = 18

    ws.freeze_panes = 'A2'

    wb.save(out)
    out.seek(0)
    return out

# =============================================================================
# MAIN APPLICATION
# =============================================================================

def main():
    load_css()

    # Header
    st.markdown("""
    <div class="main-header">
        <h1>📋 BOQ Extractor Pro v12.0</h1>
        <p>Interactive Preview with Material Anchor Recognition</p>
        <div class="badge-container">
            <span class="material-badge">⚓ Material Anchors</span>
            <span class="material-badge">🎯 Visual Zones</span>
            <span class="material-badge">📊 Live Preview</span>
            <span class="material-badge">🔧 Adjustable Regions</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Session state
    if 'data' not in st.session_state:
        st.session_state.data = pd.DataFrame()
    if 'logs' not in st.session_state:
        st.session_state.logs = []
    if 'preview_text' not in st.session_state:
        st.session_state.preview_text = ""
    if 'pdf_file' not in st.session_state:
        st.session_state.pdf_file = None
    if 'detected_materials' not in st.session_state:
        st.session_state.detected_materials = []

    # Sidebar - Settings
    with st.sidebar:
        st.markdown("### ⚙️ Page Settings")
        page_start = st.number_input("Start Page:", 1, 1000, 1)
        page_end = st.number_input("End Page:", 1, 1000, 999, help="999 = all pages")

        st.markdown("---")

        st.markdown("### 📐 Region Settings (Preview Page)")
        preview_page = st.number_input("Preview Page:", 1, 1000, 1)

        st.markdown("<div class='slider-container'>", unsafe_allow_html=True)
        x_range = st.slider("X Range (% width):", 0, 100, (50, 98), help="Left to Right position")
        y_range = st.slider("Y Range (% height):", 0, 100, (8, 45), help="Top to Bottom position")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("---")

        yellow_header = st.checkbox("Yellow Excel Header", True)
        show_material_zones = st.checkbox("Show Material Zones in Preview", True)

        st.markdown("---")

        st.markdown("### 🎯 Material Anchors")
        st.markdown("<small>Auto-detected materials:</small>", unsafe_allow_html=True)
        materials = ["A240 SS316", "A36", "A105", "A193 GR.B7", "A194 GR.2H", 
                    "Per MSS-SP58", "CARBON STEEL", "A516 GR.60"]
        for mat in materials:
            st.markdown(f"<span class='anchor-point'>{mat}</span>", unsafe_allow_html=True)

        st.markdown("---")

        if st.button("🗑️ Clear All Data", type="secondary"):
            st.session_state.data = pd.DataFrame()
            st.session_state.logs = []
            st.session_state.preview_text = ""
            st.session_state.detected_materials = []
            st.rerun()

    # Main content
    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown("<div class='section-header'><h3>📤 Upload PDF</h3></div>", 
                   unsafe_allow_html=True)

        uploaded = st.file_uploader(
            "Drop your BOQ PDF here or click to browse",
            type=['pdf'],
            help="Upload PDF with BOQ tables. The app will detect material anchor points automatically."
        )

        if uploaded:
            st.session_state.pdf_file = uploaded
            st.success(f"✅ Loaded: {uploaded.name} ({uploaded.size:,} bytes)")

    with col2:
        st.markdown("<div class='section-header'><h3>🚀 Actions</h3></div>", 
                   unsafe_allow_html=True)

        if uploaded:
            col_prev, col_ext = st.columns(2)

            with col_prev:
                preview_btn = st.button("👁️ Preview", type="secondary", use_container_width=True)

            with col_ext:
                extract_btn = st.button("🔍 Extract BOQ", type="primary", use_container_width=True)

    # Interactive Preview Section
    if uploaded and preview_btn:
        st.markdown("<div class='section-header'><h3>🎯 Interactive Preview with Anchor Points</h3></div>", 
                   unsafe_allow_html=True)

        with st.spinner("Generating annotated preview..."):
            img_data, error = generate_annotated_preview(
                uploaded, preview_page, x_range, y_range, show_material_zones
            )

            if error:
                st.error(f"Error: {error}")
            elif img_data:
                col_img, col_info = st.columns([3, 2])

                with col_img:
                    st.image(img_data, caption=f"Page {preview_page} - Blue=Extraction Region, Red=Material Anchors", 
                            use_column_width=True)

                with col_info:
                    st.markdown("### 📊 Region Configuration")
                    st.info(f"""
                    **Extraction Region:**
                    - X: {x_range[0]}% - {x_range[1]}% of width
                    - Y: {y_range[0]}% - {y_range[1]}% of height

                    **Legend:**
                    - 🔵 Blue box = Text extraction area
                    - 🔴 Red boxes = Detected material anchors
                    - ⚓ Labels = Material codes found
                    """)

                    # Extract and show text
                    text, materials = extract_text_with_preview(uploaded, preview_page, x_range, y_range)
                    st.session_state.preview_text = text
                    st.session_state.detected_materials = materials

                    if materials:
                        st.markdown("### ⚓ Materials Detected in Region")
                        for mat in materials[:10]:
                            st.markdown(f"<span class='material-zone-indicator'>{mat['material']}</span>", 
                                       unsafe_allow_html=True)
                    else:
                        st.warning("No material anchors found in this region")

        # Show extracted text
        if st.session_state.preview_text:
            st.markdown("### 📝 Extracted Text Preview")

            # Highlight materials in text
            highlighted_text = st.session_state.preview_text
            for mat in st.session_state.detected_materials:
                pattern = re.escape(mat['material'])
                highlighted_text = re.sub(
                    f"({pattern})",
                    r"<span class='material-highlight'>\1</span>",
                    highlighted_text,
                    flags=re.IGNORECASE
                )

            st.markdown(f"<div class='preview-box'>{highlighted_text}</div>", unsafe_allow_html=True)

            # Show parsing analysis
            st.markdown("### 🔍 Parsing Analysis")
            extractor = MaterialAnchoredExtractor()
            lines = st.session_state.preview_text.split('\n')

            parsed_items = []
            for i, line in enumerate(lines):
                parsed = extractor.parse_line_material_anchor(line)
                if parsed:
                    parsed_items.append((i+1, parsed))

            if parsed_items:
                st.success(f"✅ Found {len(parsed_items)} valid BOQ items")
                for line_num, item in parsed_items[:5]:
                    with st.container():
                        st.markdown(f"""
                        <div class='detection-result'>
                            <strong>Line {line_num}:</strong> Item {item['Item No']} | 
                            Qty: {item['Quantity']} | Part: {item['Part No']} | 
                            Material: <span class='material-zone-indicator'>{item['Material']}</span>
                        </div>
                        """, unsafe_allow_html=True)
            else:
                st.error("❌ No valid BOQ items found. Adjust region sliders to capture the table.")

    # Extraction
    if uploaded and extract_btn:
        progress_bar = st.progress(0, "Starting extraction...")

        with st.spinner("Processing all pages with material anchors..."):
            items, logs = extract_with_spatial_material(
                uploaded, page_start, page_end, x_range, y_range, progress_bar
            )

            st.session_state.logs = logs

            if items:
                df = pd.DataFrame(items)
                cols = ['Drawing No', 'Mark No', 'Item No', 'Quantity', 'Part No', 'Description', 'Material']
                available = [c for c in cols if c in df.columns]
                st.session_state.data = df[available]

                progress_bar.empty()
                st.success(f"✅ Extracted {len(df)} items from {df['Drawing No'].nunique()} drawings!")
            else:
                progress_bar.empty()
                st.error("❌ No items found. Use Preview mode to adjust extraction region.")

    # Results section
    if not st.session_state.data.empty:
        df = st.session_state.data

        st.markdown("<div class='section-header'><h3>📊 Extraction Results</h3></div>", 
                   unsafe_allow_html=True)

        # Stats
        c1, c2, c3, c4 = st.columns(4)

        with c1:
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-value">{len(df)}</div>
                <div class="stat-label">Total Items</div>
            </div>
            """, unsafe_allow_html=True)

        with c2:
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-value">{df['Drawing No'].nunique()}</div>
                <div class="stat-label">Drawings</div>
            </div>
            """, unsafe_allow_html=True)

        with c3:
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-value">{df['Mark No'].nunique()}</div>
                <div class="stat-label">Mark Numbers</div>
            </div>
            """, unsafe_allow_html=True)

        with c4:
            total_qty = int(df['Quantity'].sum()) if 'Quantity' in df.columns else 0
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-value">{total_qty:,}</div>
                <div class="stat-label">Total Quantity</div>
            </div>
            """, unsafe_allow_html=True)

        # Material distribution
        if 'Material' in df.columns:
            st.markdown("### 🧱 Material Distribution (Anchor-Based)")
            mat_counts = df['Material'].value_counts()

            col_chart, col_table = st.columns([2, 1])
            with col_chart:
                st.bar_chart(mat_counts)
            with col_table:
                st.dataframe(mat_counts.reset_index().rename(
                    columns={'index': 'Material', 'Material': 'Count'}
                ), use_container_width=True, hide_index=True)

        # Data editor
        st.markdown("### ✏️ Review & Edit Data")
        edited = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Drawing No": st.column_config.TextColumn("Drawing No", width="medium"),
                "Mark No": st.column_config.TextColumn("Mark No", width="medium"),
                "Item No": st.column_config.NumberColumn("Item No", width="small"),
                "Quantity": st.column_config.NumberColumn("Qty", width="small"),
                "Part No": st.column_config.TextColumn("Part No", width="medium"),
                "Description": st.column_config.TextColumn("Description", width="large"),
                "Material": st.column_config.TextColumn("Material", width="medium"),
            }
        )

        # Export
        st.markdown("<div class='section-header'><h3>📦 Export</h3></div>", 
                   unsafe_allow_html=True)

        col_excel, col_csv = st.columns(2)

        with col_excel:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel = make_excel(edited, yellow_header)
            st.download_button(
                label="📥 Download Excel (.xlsx)",
                data=excel,
                file_name=f"BOQ_Extract_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with col_csv:
            csv = edited.to_csv(index=False)
            st.download_button(
                label="📄 Download CSV (.csv)",
                data=csv,
                file_name=f"BOQ_Extract_{ts}.csv",
                mime="text/csv",
                use_container_width=True
            )

    # Logs
    if st.session_state.logs:
        with st.expander("🔍 Processing Logs", expanded=False):
            for log in st.session_state.logs:
                if log.startswith("✅"):
                    st.markdown(f"<span class='log-success'>{log}</span>", unsafe_allow_html=True)
                elif log.startswith("❌"):
                    st.markdown(f"<span class='log-error'>{log}</span>", unsafe_allow_html=True)
                elif log.startswith("⚠️"):
                    st.markdown(f"<span style='color: #ffc107;'>{log}</span>", unsafe_allow_html=True)
                else:
                    st.markdown(f"<span class='log-info'>{log}</span>", unsafe_allow_html=True)

    # Instructions
    with st.expander("📖 How to Use - Interactive Preview Guide", expanded=False):
        st.markdown("""
        ### Step-by-Step Guide with Visual Preview

        #### 1. Upload PDF
        Drop your BOQ PDF file in the upload area.

        #### 2. Adjust Region (IMPORTANT)
        Use the **sliders in the sidebar** to position the extraction region:
        - **X Range**: Controls left-right position (default: 50%-98% for right-side tables)
        - **Y Range**: Controls top-bottom position (default: 8%-45% for upper tables)

        #### 3. Preview with Anchor Points
        Click **"👁️ Preview"** to see:
        - 🔵 **Blue box**: Your extraction region
        - 🔴 **Red boxes**: Auto-detected material anchor points (A240 SS316, A36, etc.)
        - ⚓ **Labels**: Material codes the system found

        #### 4. Verify Material Detection
        Check that red boxes appear around material codes. If not:
        - Adjust X/Y sliders to include the material column
        - Ensure material codes are visible (A240 SS316, Per MSS-SP58, etc.)

        #### 5. Extract
        Click **"🔍 Extract BOQ"** to process all pages.

        ### Understanding Material Anchors

        The system uses **material codes** as anchor points to parse the table:
        ```
        Item  Qty  Part No        Description              Material
          1    1   3x240x240      SS Plate(Mirror Finish)  ⚓ A240 SS316
          2    1   V2-26-BM1      Variable Effort Support  ⚓ Per MSS-SP58
        ```

        The ⚓ anchor (Material column) tells the system where each row ends.

        ### Troubleshooting

        | Problem | Solution |
        |---------|----------|
        | No materials detected | Adjust region to include rightmost column |
        | Wrong item count | Check that material codes are clearly visible |
        | Missing rows | Ensure all rows have material codes |
        """)

    # Footer
    st.markdown("""
    <div class="footer">
        <p>📋 BOQ Extractor Pro v12.0 | Interactive Preview with Material Anchor Recognition</p>
        <p style="font-size: 0.8rem; color: #999;">
            Visual anchor points | Adjustable extraction regions | Material-zone detection
        </p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
