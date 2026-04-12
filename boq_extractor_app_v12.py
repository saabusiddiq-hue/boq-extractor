"""
BOQ Extractor Pro v15.0 - Auto Table Detection Edition
Smart Content-Based Table Extraction with Visual Grid Overlay
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import fitz
from PIL import Image, ImageDraw

# Page configuration - DARK THEME
st.set_page_config(
    page_title="BOQ Extractor Pro - Auto Detect",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# DARK THEME CUSTOM CSS
def load_css():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        
        .stApp {
            background-color: #0f1117 !important;
        }
        
        [data-testid="stSidebar"] {
            background-color: #161b22 !important;
            border-right: 1px solid #30363d;
        }
        
        h1, h2, h3, h4, h5, h6 {
            color: #e6edf3 !important;
            font-family: 'Inter', sans-serif;
            font-weight: 600;
        }
        
        .main-header {
            background: linear-gradient(135deg, #1f6feb 0%, #388bfd 100%);
            color: white;
            padding: 2rem;
            border-radius: 16px;
            text-align: center;
            margin-bottom: 2rem;
            box-shadow: 0 10px 40px rgba(31, 111, 235, 0.3);
            border: 1px solid #30363d;
        }
        
        .main-header h1 {
            font-size: 2.5rem;
            font-weight: 700;
            margin-bottom: 0.5rem;
            color: white !important;
        }
        
        .badge-container {
            display: flex;
            gap: 10px;
            justify-content: center;
            flex-wrap: wrap;
        }
        
        .feature-badge {
            background-color: rgba(255,255,255,0.15);
            color: white;
            padding: 6px 16px;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 500;
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255,255,255,0.2);
        }
        
        .section-header {
            background: linear-gradient(90deg, #21262d 0%, #161b22 100%);
            padding: 1rem 1.5rem;
            border-radius: 10px;
            margin: 2rem 0 1rem 0;
            border-left: 4px solid #1f6feb;
            border: 1px solid #30363d;
        }
        
        .section-header h3 {
            margin: 0;
            color: #58a6ff !important;
            font-weight: 600;
        }
        
        .stat-card {
            background: #161b22;
            border-radius: 12px;
            padding: 1.5rem;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            border-left: 4px solid #1f6feb;
            border: 1px solid #30363d;
        }
        
        .stat-value {
            font-size: 2rem;
            font-weight: 700;
            color: #58a6ff;
        }
        
        .stat-label {
            font-size: 0.9rem;
            color: #8b949e;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .extracted-text {
            background: #0d1117;
            color: #e6edf3;
            padding: 1.5rem;
            border-radius: 12px;
            font-family: 'Courier New', monospace;
            font-size: 0.85rem;
            max-height: 400px;
            overflow-y: auto;
            white-space: pre-wrap;
            border: 1px solid #30363d;
        }
        
        .slider-label {
            font-weight: 600;
            color: #e6edf3;
            margin-bottom: 0.5rem;
            font-size: 0.9rem;
        }
        
        .crop-info {
            background: #21262d;
            border-radius: 8px;
            padding: 1rem;
            border: 1px solid #30363d;
            margin-top: 1rem;
        }
        
        .crop-info h4 {
            color: #58a6ff !important;
            margin-bottom: 0.5rem;
        }
        
        .crop-dimension {
            display: flex;
            justify-content: space-between;
            padding: 0.3rem 0;
            border-bottom: 1px solid #30363d;
            color: #c9d1d9;
        }
        
        .stButton>button {
            background: linear-gradient(135deg, #238636 0%, #2ea043 100%) !important;
            color: white !important;
            border: none !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
        }
        
        .stButton>button:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 4px 12px rgba(35, 134, 54, 0.4) !important;
        }
        
        .stButton>button[kind="secondary"] {
            background: #21262d !important;
            border: 1px solid #30363d !important;
            color: #c9d1d9 !important;
        }
        
        .stSlider>div>div>div {
            background-color: #1f6feb !important;
        }
        
        .stFileUploader>div>div {
            background-color: #161b22 !important;
            border: 2px dashed #30363d !important;
            border-radius: 10px !important;
            color: #8b949e !important;
        }
        
        .stSuccess {
            background-color: rgba(35, 134, 54, 0.1) !important;
            border: 1px solid #238636 !important;
            color: #3fb950 !important;
        }
        
        .stWarning {
            background-color: rgba(210, 153, 34, 0.1) !important;
            border: 1px solid #d29922 !important;
            color: #e3b341 !important;
        }
        
        .stInfo {
            background-color: rgba(56, 139, 253, 0.1) !important;
            border: 1px solid #388bfd !important;
            color: #58a6ff !important;
        }
        
        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }
        
        ::-webkit-scrollbar-track {
            background: #161b22;
            border-radius: 4px;
        }
        
        ::-webkit-scrollbar-thumb {
            background: #1f6feb;
            border-radius: 4px;
        }
        
        .footer {
            text-align: center;
            padding: 2rem;
            color: #8b949e;
            border-top: 1px solid #30363d;
            margin-top: 3rem;
        }
        
        p, span, label {
            color: #c9d1d9 !important;
        }
        
        .detected-column {
            background: rgba(31, 111, 235, 0.2);
            border-left: 3px solid #1f6feb;
            padding: 0.5rem;
            margin: 0.25rem 0;
            border-radius: 4px;
            color: #e6edf3;
        }
    </style>
    """, unsafe_allow_html=True)

# =============================================================================
# SMART TABLE DETECTOR - Content Based
# =============================================================================

class SmartTableDetector:
    """
    Detects table structure based on content patterns rather than grid lines.
    Optimized for BOQ tables with columns: Item, No. Req'd, Fig. No., Description, Material
    """
    
    def __init__(self):
        self.column_patterns = {
            'item_no': [r'^\d+$', r'^\d+\.?$'],  # Just numbers: 1, 2, 3, 4...
            'req_qty': [r'^\d+$', r'^\d+\s*(?:EA|PC|SET|NOS)?$'],  # Quantities: 1, 2, 4
            'fig_no': [r'^[A-Z]\d+[-_][A-Z]?\d+', r'^F\d+[-_]M\d+', r'^V\d+[-_]\d+[-_][A-Z]+\d*', r'^PC\d+[-_]\d+'],  # F125-M36, V2-21-TS2
            'description': [r'(?:Beam|Support|Nut|Rod|Clamp|Pipe|Thread|Weldless|Inverted|Variable|Full|All)\s*(?:Welding|Effort|Eye|Attachment)?'],  # Text descriptions
            'material': [r'A36\b', r'A105\b', r'A193\s*GR?[.]?B7', r'A194\s*GR?[.]?2H', r'Per\s*MSS[-]?SP\d+', r'SS316', r'SS304', r'A516']  # Material codes
        }
        
        self.max_items = 15  # Only extract items 1-15
        
    def analyze_text_structure(self, text_lines):
        """
        Analyze text to detect column boundaries based on word positions.
        Returns detected columns with their x-position ranges.
        """
        if not text_lines:
            return []
        
        # Collect all words with their positions
        word_positions = []
        for line_idx, line in enumerate(text_lines):
            words = line.strip().split()
            # Estimate word positions (simplified - assumes monospace/fixed width)
            char_pos = 0
            for word in words:
                word_positions.append({
                    'word': word,
                    'line': line_idx,
                    'start_pos': char_pos,
                    'end_pos': char_pos + len(word)
                })
                char_pos += len(word) + 1  # +1 for space
        
        # Detect column boundaries by finding gaps in word distribution
        # This is a simplified version - in production, use pdfplumber's char/word positions
        return word_positions
    
    def detect_columns_from_pdf(self, page):
        """
        Use pdfplumber to detect column positions from PDF character data.
        """
        chars = page.chars
        if not chars:
            return None
        
        # Group characters by their x0 position (left edge)
        # Round to nearest 10 pixels to group into columns
        x_positions = {}
        for char in chars:
            x_rounded = round(char['x0'] / 10) * 10
            if x_rounded not in x_positions:
                x_positions[x_rounded] = []
            x_positions[x_rounded].append(char)
        
        # Find clusters of x-positions that form columns
        sorted_x = sorted(x_positions.keys())
        columns = []
        current_col = [sorted_x[0]] if sorted_x else []
        
        for i in range(1, len(sorted_x)):
            if sorted_x[i] - sorted_x[i-1] < 50:  # Within 50 pixels = same column
                current_col.append(sorted_x[i])
            else:
                # Save current column
                avg_x = sum(current_col) / len(current_col)
                columns.append({
                    'x': avg_x,
                    'width': 80,  # Estimated width
                    'chars': sum(len(x_positions[x]) for x in current_col)
                })
                current_col = [sorted_x[i]]
        
        if current_col:
            avg_x = sum(current_col) / len(current_col)
            columns.append({
                'x': avg_x,
                'width': 80,
                'chars': sum(len(x_positions[x]) for x in current_col)
            })
        
        # Filter to keep only significant columns (with enough characters)
        significant_cols = [c for c in columns if c['chars'] > 20]
        
        return significant_cols[:5]  # Return top 5 columns
    
    def parse_table_content(self, text, column_boundaries=None):
        """
        Parse table content into structured data.
        Handles the specific format: Item | Qty | Fig No | Description | Material
        """
        lines = text.strip().split('\n')
        items = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Try to parse based on expected patterns
            parsed = self._parse_line_smart(line)
            
            if parsed and parsed.get('Item No'):
                item_no = parsed['Item No']
                # Only include items 1-15
                if 1 <= item_no <= self.max_items:
                    items.append(parsed)
                elif item_no > self.max_items:
                    # Stop after item 15 (Spring details usually start after)
                    break
        
        return items
    
    def _parse_line_smart(self, line):
        """
        Smart line parsing that detects columns by content type.
        """
        # Skip header lines
        skip_keywords = ['ITEM', 'NO.', "REQ'D", 'FIG', 'DESCRIPTION', 'MATERIAL', 
                        'NO', 'REQUIRED', 'FIGURE', 'PART', 'DRAWING']
        
        first_word = line.split()[0].upper() if line.split() else ""
        if any(kw in first_word for kw in skip_keywords) and len(line.split()) <= 3:
            return None
        
        parts = line.split()
        if len(parts) < 3:
            return None
        
        result = {
            'Item No': None,
            'Quantity': None,
            'Fig No': '',
            'Description': '',
            'Material': ''
        }
        
        # Pattern 1: Item No is first number (1-15)
        if parts[0].isdigit():
            item_no = int(parts[0])
            if 1 <= item_no <= 15:
                result['Item No'] = item_no
                remaining = parts[1:]
                
                # Next should be quantity (usually small number 1-10)
                if remaining and remaining[0].isdigit() and 1 <= int(remaining[0]) <= 10:
                    result['Quantity'] = int(remaining[0])
                    remaining = remaining[1:]
                
                # Look for Fig No pattern (F125-M36, V2-21-TS2, etc.)
                fig_idx = -1
                for i, part in enumerate(remaining[:3]):
                    if self._is_fig_no(part):
                        result['Fig No'] = part
                        fig_idx = i
                        break
                
                # Everything between Fig No and Material is Description
                if fig_idx >= 0:
                    desc_parts = []
                    mat_parts = []
                    
                    # Check remaining parts for material patterns
                    for i, part in enumerate(remaining[fig_idx+1:], start=fig_idx+1):
                        if self._is_material(part) or (i < len(remaining)-1 and self._is_material(part + ' ' + remaining[i+1])):
                            mat_parts.append(part)
                        elif mat_parts:
                            mat_parts.append(part)
                        else:
                            desc_parts.append(part)
                    
                    result['Description'] = ' '.join(desc_parts)
                    result['Material'] = ' '.join(mat_parts) if mat_parts else self._extract_material_from_end(' '.join(remaining[fig_idx+1:]))
                else:
                    # No fig no found, treat as description
                    result['Description'] = ' '.join(remaining)
        
        return result if result['Item No'] else None
    
    def _is_fig_no(self, text):
        """Check if text matches Fig No pattern"""
        patterns = [
            r'^[A-Z]\d+[-_][A-Z]?\d+',
            r'^F\d+[-_]M\d+',
            r'^V\d+[-_]\d+[-_][A-Z]+\d*',
            r'^PC\d+[-_]\d+',
            r'^F\d+[-_]TS\d+',
            r'^F\d+[-_]M\d+\s*L?=\d*'
        ]
        return any(re.match(p, text, re.IGNORECASE) for p in patterns)
    
    def _is_material(self, text):
        """Check if text is a material code"""
        patterns = [r'A36\b', r'A105\b', r'A193', r'A194', r'MSS[-]?SP', r'SS316', r'SS304', r'A516', r'A240']
        return any(re.search(p, text, re.IGNORECASE) for p in patterns)
    
    def _extract_material_from_end(self, text):
        """Extract material from end of description"""
        # Common material patterns at end of line
        mat_match = re.search(r'(A36|A105|A193\s*GR?[.]?B7|A194\s*GR?[.]?2H|Per\s*MSS[-]?SP\d+|SS316|SS304|A516)(.*?)$', text, re.IGNORECASE)
        if mat_match:
            return mat_match.group(0).strip()
        return ''

# =============================================================================
# VISUAL GRID OVERLAY
# =============================================================================

def add_grid_overlay(img_data, column_boundaries, row_positions=None):
    """
    Add visual grid lines to the preview image.
    """
    try:
        # Convert bytes to PIL Image
        img = Image.open(io.BytesIO(img_data))
        draw = ImageDraw.Draw(img)
        
        width, height = img.size
        
        # Draw vertical column lines
        for col in column_boundaries:
            x = int(col['x'])
            if 0 <= x < width:
                # Draw vertical line
                draw.line([(x, 0), (x, height)], fill=(31, 111, 235), width=2)  # Blue lines
                # Draw column label at top
                draw.rectangle([x-2, 0, x+2, 20], fill=(31, 111, 235))
        
        # Draw horizontal row lines if provided
        if row_positions:
            for y in row_positions:
                if 0 <= y < height:
                    draw.line([(0, y), (width, y)], fill=(248, 81, 73), width=1)  # Red lines
        
        # Convert back to bytes
        output = io.BytesIO()
        img.save(output, format='PNG')
        return output.getvalue()
    except Exception as e:
        return img_data

# =============================================================================
# PDF PROCESSING WITH AUTO DETECTION
# =============================================================================

def render_pdf_with_grid(pdf_file, page_num, detector, zoom_level=2.0):
    """
    Render PDF with automatic column detection and grid overlay.
    """
    try:
        pdf_file.seek(0)
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        
        if page_num > len(doc):
            doc.close()
            return None, None, "Invalid page number"
        
        page = doc[page_num - 1]
        
        # Detect columns using pdfplumber
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            plum_page = pdf.pages[page_num - 1]
            columns = detector.detect_columns_from_pdf(plum_page)
        
        # Render base image
        mat = fitz.Matrix(zoom_level, zoom_level)
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")
        
        # Add grid overlay if columns detected
        if columns:
            img_data = add_grid_overlay(img_data, columns)
        
        doc.close()
        return img_data, columns, None
        
    except Exception as e:
        return None, None, str(e)

def extract_with_smart_detection(pdf_file, page_num, crop_box, detector):
    """
    Extract text using smart table detection.
    """
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            page = pdf.pages[page_num - 1]
            width, height = page.width, page.height
            
            # Apply crop if specified
            if crop_box:
                x1_pct, y1_pct, x2_pct, y2_pct = crop_box
                x1 = width * (x1_pct / 100)
                y1 = height * (y1_pct / 100)
                x2 = width * (x2_pct / 100)
                y2 = height * (y2_pct / 100)
                page = page.crop((x1, y1, x2, y2))
            
            text = page.extract_text()
            lines = text.split('\n') if text else []
            
            # Parse with smart detector
            items = detector.parse_table_content(text)
            
            return items, text, None
            
    except Exception as e:
        return [], "", str(e)

# =============================================================================
# EXCEL EXPORT
# =============================================================================

def create_excel_download(df, yellow_header=True):
    """Create styled Excel file"""
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ Data"
    
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_font = Font(bold=True, size=11, color="000000")
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # Reorder columns
    preferred_order = ['Item No', 'Quantity', 'Fig No', 'Description', 'Material', 'Page']
    available_cols = [c for c in preferred_order if c in df.columns]
    other_cols = [c for c in df.columns if c not in preferred_order]
    final_cols = available_cols + other_cols
    
    df_export = df[final_cols]
    
    # Write headers
    for col_num, header in enumerate(df_export.columns, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        if yellow_header:
            cell.fill = yellow_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    
    # Write data
    for row_num, row_data in enumerate(df_export.values, 2):
        for col_num, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    
    # Auto-adjust column widths
    for idx, col in enumerate(df_export.columns, 1):
        max_length = len(str(col))
        
        if idx <= 26:
            col_letter = chr(64 + idx)
        else:
            col_letter = chr(64 + (idx-1)//26) + chr(65 + (idx-1)%26)
        
        for row_idx in range(1, len(df_export) + 2):
            cell_value = ws.cell(row=row_idx, column=idx).value
            try:
                cell_length = len(str(cell_value)) if cell_value is not None else 0
                if cell_length > max_length:
                    max_length = cell_length
            except:
                pass
        
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[col_letter].width = adjusted_width
    
    ws.freeze_panes = 'A2'
    wb.save(output)
    output.seek(0)
    return output

# =============================================================================
# MAIN APPLICATION
# =============================================================================

def main():
    load_css()
    
    st.markdown("""
    <div class="main-header">
        <h1>📋 BOQ Extractor Pro v15.0</h1>
        <p>Auto Table Detection - Smart Content Recognition</p>
        <div class="badge-container">
            <span class="feature-badge">🤖 Auto Detect</span>
            <span class="feature-badge">📐 Visual Grid</span>
            <span class="feature-badge">🎯 Items 1-15</span>
            <span class="feature-badge">📊 Smart Parse</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Initialize session state
    if 'extracted_data' not in st.session_state:
        st.session_state.extracted_data = pd.DataFrame()
    if 'detector' not in st.session_state:
        st.session_state.detector = SmartTableDetector()
    if 'pdf_file' not in st.session_state:
        st.session_state.pdf_file = None
    if 'detected_columns' not in st.session_state:
        st.session_state.detected_columns = None
    
    # Sidebar
    with st.sidebar:
        st.markdown("### 🎯 Detection Settings")
        
        max_items = st.number_input("Max Items to Extract", 1, 50, 15, 
                                  help="Stop extraction after this item number (excludes Spring details)")
        st.session_state.detector.max_items = max_items
        
        st.markdown("---")
        
        st.markdown("### 📐 Region (Optional)")
        st.markdown("<div class='slider-label'>Left (X1) %</div>", unsafe_allow_html=True)
        x1 = st.slider("X1", 0, 100, 5, key="x1", label_visibility="collapsed")
        
        st.markdown("<div class='slider-label'>Top (Y1) %</div>", unsafe_allow_html=True)
        y1 = st.slider("Y1", 0, 100, 15, key="y1", label_visibility="collapsed")
        
        st.markdown("<div class='slider-label'>Right (X2) %</div>", unsafe_allow_html=True)
        x2 = st.slider("X2", 0, 100, 95, key="x2", label_visibility="collapsed")
        
        st.markdown("<div class='slider-label'>Bottom (Y2) %</div>", unsafe_allow_html=True)
        y2 = st.slider("Y2", 0, 100, 60, key="y2", label_visibility="collapsed")
        
        crop_box = (x1, y1, x2, y2) if x2 > x1 and y2 > y1 else None
        
        st.markdown("---")
        
        st.markdown("### 🔍 Preview Settings")
        zoom_level = st.select_slider("Zoom", options=[1.0, 1.5, 2.0, 2.5, 3.0], value=2.0, format_func=lambda x: f"{x}x")
        
        st.markdown("---")
        
        yellow_header = st.checkbox("Yellow Excel Header", value=True)
        
        if st.button("🗑️ Clear Data", type="secondary"):
            st.session_state.extracted_data = pd.DataFrame()
            st.session_state.pdf_file = None
            st.session_state.detected_columns = None
            st.rerun()
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("<div class='section-header'><h3>📤 Upload PDF</h3></div>", unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader("Upload BOQ PDF", type=['pdf'])
        
        if uploaded_file:
            st.success(f"✅ Loaded: {uploaded_file.name}")
            st.session_state.pdf_file = uploaded_file
    
    with col2:
        st.markdown("<div class='section-header'><h3>🚀 Actions</h3></div>", unsafe_allow_html=True)
        
        if uploaded_file:
            detect_btn = st.button("🔍 Auto Detect & Extract", use_container_width=True, type="primary")
            preview_btn = st.button("👁️ Preview with Grid", use_container_width=True, type="secondary")
    
    # Preview with Grid
    if uploaded_file and preview_btn:
        st.markdown("<div class='section-header'><h3>🎯 Visual Grid Preview</h3></div>", unsafe_allow_html=True)
        
        with st.spinner("Detecting table structure..."):
            img_data, columns, error = render_pdf_with_grid(
                uploaded_file, 1, st.session_state.detector, zoom_level
            )
            
            if error:
                st.error(f"Error: {error}")
            else:
                st.session_state.detected_columns = columns
                
                # Show image with grid
                st.image(img_data, caption="Detected columns shown with blue vertical lines", use_column_width=True)
                
                if columns:
                    st.success(f"✅ Detected {len(columns)} columns automatically")
                    
                    # Show column positions
                    col_data = []
                    for i, col in enumerate(columns, 1):
                        col_data.append({
                            "Column": f"Col {i}",
                            "Position (px)": f"{col['x']:.0f}",
                            "Char Count": col['chars']
                        })
                    
                    st.markdown("### 📊 Detected Column Positions")
                    st.dataframe(pd.DataFrame(col_data), hide_index=True, use_container_width=True)
                else:
                    st.warning("⚠️ No columns detected. Try adjusting crop region.")
    
    # Auto Detect & Extract
    if uploaded_file and detect_btn:
        st.markdown("<div class='section-header'><h3>⚙️ Smart Extraction</h3></div>", unsafe_allow_html=True)
        
        with st.spinner("Analyzing table content..."):
            items, raw_text, error = extract_with_smart_detection(
                uploaded_file, 1, crop_box, st.session_state.detector
            )
            
            if error:
                st.error(f"Extraction error: {error}")
            elif items:
                df = pd.DataFrame(items)
                st.session_state.extracted_data = df
                
                st.success(f"✅ Extracted {len(df)} items (Items 1-{max_items})")
                
                # Show raw text preview
                with st.expander("📝 View Raw Extracted Text"):
                    st.markdown(f"<div class='extracted-text'>{raw_text[:2000]}</div>", unsafe_allow_html=True)
            else:
                st.warning("⚠️ No items found. Check if table is in selected region.")
    
    # Results
    if not st.session_state.extracted_data.empty:
        df = st.session_state.extracted_data
        
        st.markdown("<div class='section-header'><h3>📊 Extraction Results</h3></div>", unsafe_allow_html=True)
        
        # Stats
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-value">{len(df)}</div>
                <div class="stat-label">Items Extracted</div>
            </div>
            """, unsafe_allow_html=True)
        
        with c2:
            total_qty = int(df['Quantity'].sum()) if 'Quantity' in df.columns else 0
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-value">{total_qty}</div>
                <div class="stat-label">Total Quantity</div>
            </div>
            """, unsafe_allow_html=True)
        
        with c3:
            materials = df['Material'].nunique() if 'Material' in df.columns else 0
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-value">{materials}</div>
                <div class="stat-label">Materials</div>
            </div>
            """, unsafe_allow_html=True)
        
        # Data editor
        st.markdown("### ✏️ Review Data")
        
        edited_df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Item No": st.column_config.NumberColumn("Item", width="small"),
                "Quantity": st.column_config.NumberColumn("Qty", width="small"),
                "Fig No": st.column_config.TextColumn("Fig No", width="medium"),
                "Description": st.column_config.TextColumn("Description", width="large"),
                "Material": st.column_config.TextColumn("Material", width="medium"),
            }
        )
        
        # Export
        st.markdown("<div class='section-header'><h3>📦 Export</h3></div>", unsafe_allow_html=True)
        
        col_excel, col_csv = st.columns(2)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        with col_excel:
            try:
                excel_data = create_excel_download(edited_df, yellow_header)
                st.download_button(
                    label="📥 Download Excel",
                    data=excel_data,
                    file_name=f"BOQ_Auto_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Export error: {e}")
        
        with col_csv:
            csv_data = edited_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📄 Download CSV",
                data=csv_data,
                file_name=f"BOQ_Auto_{timestamp}.csv",
                mime="text/csv",
                use_container_width=True
            )
    
    # Instructions
    with st.expander("📖 How Auto Detection Works"):
        st.markdown("""
        ### 🤖 Smart Table Detection
        
        **1. Content-Based Detection**
        - Detects columns by analyzing text positions
        - No need for visible grid lines
        - Identifies: Item No | Qty | Fig No | Description | Material
        
        **2. Visual Grid Overlay**
        - Blue lines show detected column boundaries
        - Helps verify correct detection
        - Adjust crop region if needed
        
        **3. Smart Parsing**
        - Recognizes Fig No patterns: F125-M36, V2-21-TS2, PC3-400
        - Extracts materials: A36, A105, A193 GR.B7, Per MSS-SP58
        - Stops at item 15 (excludes Spring details section)
        
        **4. Manual Adjustments**
        - Use crop sliders to focus on table area
        - Change "Max Items" if needed
        - Edit data before export
        """)
    
    st.markdown("""
    <div class="footer">
        <p>📋 BOQ Extractor Pro v15.0 | Auto Detection Edition</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
