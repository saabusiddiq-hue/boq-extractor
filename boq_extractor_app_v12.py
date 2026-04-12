"""
BOQ Extractor Pro v13.1 - Region-Based Crop Extraction (Fixed)
PDF to Excel/CSV Converter with Interactive Area Selection & Adjustable Preview
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
    page_title="BOQ Extractor Pro - Crop Region",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for professional UI
def load_css():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

        * {
            font-family: 'Inter', sans-serif;
        }

        .main-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 2rem;
            border-radius: 16px;
            text-align: center;
            margin-bottom: 2rem;
            box-shadow: 0 10px 40px rgba(102, 126, 234, 0.3);
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

        .feature-badge {
            background-color: rgba(255,255,255,0.2);
            color: white;
            padding: 6px 16px;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 500;
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255,255,255,0.3);
        }

        .crop-controls {
            background: #f8f9fa;
            border-radius: 12px;
            padding: 1.5rem;
            border: 2px solid #e9ecef;
            margin-bottom: 1rem;
        }

        .preview-box {
            border: 3px solid #667eea;
            border-radius: 12px;
            overflow: hidden;
            position: relative;
            background: #1e1e1e;
        }

        .region-overlay {
            position: absolute;
            border: 3px dashed #ff6b6b;
            background: rgba(255, 107, 107, 0.1);
            pointer-events: none;
        }

        .stat-card {
            background: white;
            border-radius: 12px;
            padding: 1.5rem;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
            border-left: 4px solid #667eea;
        }

        .stat-value {
            font-size: 2rem;
            font-weight: 700;
            color: #667eea;
        }

        .stat-label {
            font-size: 0.9rem;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        .section-header {
            background: linear-gradient(90deg, #f8f9fa 0%, #e9ecef 100%);
            padding: 1rem 1.5rem;
            border-radius: 10px;
            margin: 2rem 0 1rem 0;
            border-left: 4px solid #667eea;
        }

        .section-header h3 {
            margin: 0;
            color: #667eea;
            font-weight: 600;
        }

        .extracted-text {
            background: #1e1e1e;
            color: #d4d4d4;
            padding: 1.5rem;
            border-radius: 12px;
            font-family: 'Courier New', monospace;
            font-size: 0.9rem;
            max-height: 400px;
            overflow-y: auto;
            white-space: pre-wrap;
            border: 1px solid #333;
        }

        .data-grid {
            background: white;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        }

        .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            padding: 0.75rem 1.5rem;
            border-radius: 8px;
            font-weight: 600;
            cursor: pointer;
            transition: transform 0.2s;
        }

        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
        }

        .success-badge {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background: #d4edda;
            color: #155724;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
        }

        .warning-badge {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background: #fff3cd;
            color: #856404;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
        }

        .slider-label {
            font-weight: 600;
            color: #495057;
            margin-bottom: 0.5rem;
            font-size: 0.9rem;
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
            background: #667eea;
            border-radius: 4px;
        }

        .footer {
            text-align: center;
            padding: 2rem;
            color: #666;
            border-top: 1px solid #eee;
            margin-top: 3rem;
        }
        
        .crop-preview-large {
            max-width: 100%;
            border: 2px solid #667eea;
            border-radius: 8px;
            margin-top: 10px;
        }
    </style>
    """, unsafe_allow_html=True)

# =============================================================================
# SMART BOQ PARSER
# =============================================================================

class BOQParser:
    """
    Intelligent BOQ parser that extracts data from text without strict material anchoring.
    Uses pattern recognition and positional analysis.
    """

    def __init__(self):
        # Common material patterns for recognition (not anchoring)
        self.material_patterns = [
            r'A240\s+SS316', r'A240\s+SS304', r'A36\b', r'A105\b',
            r'A193\s+GR\.B7', r'A194\s+GR\.2H', r'A193\s+B7', r'A194\s+2H',
            r'Per\s+MSS-SP[0-9]+', r'CARBON\s+STEEL', r'SS316', r'SS304',
            r'A516\s+GR\.60', r'A516\s+GR\.70', r'A240\b', r'A193\b',
            r'A194\b', r'A516\b', r'Graphite\s+bronze', r'BRONZE', r'SS\s+316'
        ]
        
        # Part number patterns
        self.part_patterns = [
            r'^[A-Z][0-9]+-[A-Z][0-9]+$',           # M1-V1, P1-S1
            r'^V[0-9]+-[0-9]+-[A-Z]+[0-9]*$',       # V2-26-BM1
            r'^M[0-9]+x[0-9]+$',                     # M10x20
            r'^PC[0-9]+-[0-9\-]+$',                  # PC1-2, PC1-2-3
            r'^[0-9]+x[0-9]+x[0-9]+$',               # 100x50x10
            r'^[A-Z][0-9]+$',                        # P1, V1
            r'^[0-9]+-[A-Z][0-9]+$',                 # 1-M1
        ]

    def is_item_number(self, text):
        """Check if text is likely an item number (1-999)"""
        try:
            num = int(text.strip())
            return 1 <= num <= 999
        except:
            return False

    def is_quantity(self, text):
        """Check if text is likely a quantity (1-9999)"""
        try:
            num = int(text.strip())
            return 1 <= num <= 9999
        except:
            return False

    def is_part_number(self, text):
        """Check if text matches part number patterns"""
        text = text.strip()
        if not text:
            return False
        return any(re.match(pattern, text, re.IGNORECASE) for pattern in self.part_patterns)

    def extract_material(self, text):
        """Extract material from text if present"""
        text = text.strip()
        for pattern in self.material_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(0), text[:match.start()].strip(), text[match.end():].strip()
        return "", text, ""

    def parse_line(self, line):
        """
        Parse a single BOQ line using flexible pattern matching.
        Expected formats:
        - "1 2 M1-V1 Description text A240 SS316"
        - "1 2 V2-26-BM1 Variable Spring Hanger Per MSS-SP58"
        - "Item Qty PartNo Description Material"
        """
        line = line.strip()
        if not line or len(line) < 10:
            return None

        # Skip header lines
        skip_keywords = ['ITEM', 'NO.', 'QTY', 'QUANTITY', 'PART', 'DESCRIPTION', 
                        'MATERIAL', 'DRAWING', 'TITLE', 'NOTE:', 'DATE:']
        first_word = line.split()[0].upper() if line.split() else ""
        if any(kw in first_word for kw in skip_keywords):
            return None

        parts = line.split()
        if len(parts) < 4:
            return None

        # Try to identify columns
        item_no = None
        quantity = None
        part_no = None
        description = ""
        material = ""

        # Strategy 1: First number is Item No, second is Quantity
        if self.is_item_number(parts[0]) and self.is_quantity(parts[1]):
            item_no = int(parts[0])
            quantity = int(parts[1])
            
            # Look for part number in remaining parts
            remaining = parts[2:]
            
            # Find part number (usually first or second token after qty)
            part_idx = -1
            for i, part in enumerate(remaining[:3]):  # Check first 3 tokens
                if self.is_part_number(part):
                    part_no = part
                    part_idx = i
                    break
            
            if part_no is None:
                # Try to find anything that looks like a part code
                for i, part in enumerate(remaining[:3]):
                    if len(part) >= 3 and any(c.isdigit() for c in part) and any(c.isalpha() for c in part):
                        part_no = part
                        part_idx = i
                        break
            
            if part_no:
                # Everything between part_no and material is description
                remaining_str = ' '.join(remaining[part_idx+1:])
                
                # Try to extract material from end
                mat, desc_before, desc_after = self.extract_material(remaining_str)
                if mat:
                    material = mat
                    description = desc_before
                else:
                    description = remaining_str
            else:
                # No part number found, treat rest as description
                description = ' '.join(remaining)
                
        else:
            return None

        if item_no is None or quantity is None:
            return None

        return {
            'Item No': item_no,
            'Quantity': quantity,
            'Part No': part_no or "",
            'Description': description.strip(),
            'Material': material
        }

    def parse_text_block(self, text):
        """Parse multiple lines of text"""
        items = []
        lines = text.split('\n')
        
        for line in lines:
            parsed = self.parse_line(line)
            if parsed:
                items.append(parsed)
        
        return items

# =============================================================================
# PDF PROCESSING FUNCTIONS
# =============================================================================

def render_pdf_page(pdf_file, page_num, crop_box=None, zoom_level=2.0):
    """
    Render PDF page with optional crop box overlay.
    crop_box: (x1, y1, x2, y2) in percentages
    zoom_level: Multiplier for resolution (1.0 = default, 2.0 = 2x, etc.)
    """
    try:
        pdf_file.seek(0)
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        
        if page_num > len(doc):
            doc.close()
            return None, "Invalid page number"
        
        page = doc[page_num - 1]
        rect = page.rect
        width, height = rect.width, rect.height
        
        # Draw crop box if provided
        if crop_box:
            x1_pct, y1_pct, x2_pct, y2_pct = crop_box
            x1 = width * (x1_pct / 100)
            y1 = height * (y1_pct / 100)
            x2 = width * (x2_pct / 100)
            y2 = height * (y2_pct / 100)
            
            # Draw crop rectangle - FIXED: removed opacity parameter
            crop_rect = fitz.Rect(x1, y1, x2, y2)
            # Draw border only (no fill to avoid opacity issue)
            page.draw_rect(crop_rect, color=(1, 0.2, 0.2), width=3)  # Red border
            
            # Add semi-transparent fill using a separate shape approach
            # Create a shape for the semi-transparent fill
            shape = page.new_shape()
            shape.draw_rect(crop_rect)
            # Light red fill with transparency (color values 0-1, alpha 0.1)
            shape.finish(color=(1, 0.2, 0.2), fill=(1, 0.8, 0.8), fill_opacity=0.3)
            shape.commit()
            
            # Add label
            label_point = fitz.Point(x1 + 5, y1 - 5)
            page.insert_text(label_point, "EXTRACTION AREA", 
                           fontsize=14, color=(1, 0.2, 0.2), fontname="helv")
        
        # Render with configurable zoom
        mat = fitz.Matrix(zoom_level, zoom_level)
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")
        
        doc.close()
        return img_data, None
        
    except Exception as e:
        return None, str(e)

def extract_text_from_region(pdf_file, page_num, crop_box):
    """Extract text from specific region of PDF"""
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            page = pdf.pages[page_num - 1]
            width, height = page.width, page.height
            
            x1_pct, y1_pct, x2_pct, y2_pct = crop_box
            x1 = width * (x1_pct / 100)
            y1 = height * (y1_pct / 100)
            x2 = width * (x2_pct / 100)
            y2 = height * (y2_pct / 100)
            
            # Crop and extract
            cropped = page.crop((x1, y1, x2, y2))
            text = cropped.extract_text()
            
            return text
            
    except Exception as e:
        return f"Error extracting text: {str(e)}"

def process_pdf_batch(pdf_file, start_page, end_page, crop_box, progress_callback=None):
    """Process multiple pages with the same crop region"""
    all_items = []
    logs = []
    
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            total_pages = len(pdf.pages)
            actual_end = min(end_page, total_pages) if end_page != 999 else total_pages
            
            logs.append(f"📁 Processing pages {start_page} to {actual_end} (Total: {total_pages})")
            
            parser = BOQParser()
            
            for page_num in range(start_page - 1, actual_end):
                if progress_callback:
                    progress = (page_num - start_page + 2) / (actual_end - start_page + 1)
                    progress_callback(min(progress, 1.0), f"Processing page {page_num + 1}...")
                
                page = pdf.pages[page_num]
                width, height = page.width, page.height
                
                # Calculate crop coordinates
                x1_pct, y1_pct, x2_pct, y2_pct = crop_box
                x1 = width * (x1_pct / 100)
                y1 = height * (y1_pct / 100)
                x2 = width * (x2_pct / 100)
                y2 = height * (y2_pct / 100)
                
                # Extract from region
                cropped = page.crop((x1, y1, x2, y2))
                text = cropped.extract_text()
                
                # Parse items
                items = parser.parse_text_block(text)
                
                # Add metadata if available (look for drawing info in bottom or top)
                # This is optional enhancement
                full_text = page.extract_text()
                drawing_no = ""
                
                # Try to find drawing number patterns
                dwg_patterns = [
                    r'Q\d+\s+SUPP\s+\d+-\d+',
                    r'DWG[.\s]*[A-Z0-9\-]+',
                    r'DRAWING[.\s]*NO[.\s]*[A-Z0-9\-]+'
                ]
                for pattern in dwg_patterns:
                    match = re.search(pattern, full_text, re.IGNORECASE)
                    if match:
                        drawing_no = match.group(0)
                        break
                
                for item in items:
                    item['Drawing No'] = drawing_no
                    item['Page'] = page_num + 1
                
                status = "✅" if items else "⚠️"
                logs.append(f"{status} Page {page_num+1}: {len(items)} items extracted")
                all_items.extend(items)
                
    except Exception as e:
        logs.append(f"❌ Error: {str(e)}")
    
    return all_items, logs

# =============================================================================
# EXCEL EXPORT
# =============================================================================

def create_excel_download(df, yellow_header=True):
    """Create styled Excel file"""
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ Data"
    
    # Styles
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_font = Font(bold=True, size=11, color="000000")
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # Reorder columns
    preferred_order = ['Drawing No', 'Page', 'Item No', 'Quantity', 'Part No', 'Description', 'Material']
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
        for cell in ws.column[idx]:
            try:
                cell_length = len(str(cell.value))
                if cell_length > max_length:
                    max_length = cell_length
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[chr(64 + idx)].width = adjusted_width
    
    # Freeze header
    ws.freeze_panes = 'A2'
    
    wb.save(output)
    output.seek(0)
    return output

# =============================================================================
# MAIN APPLICATION
# =============================================================================

def main():
    load_css()
    
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>📋 BOQ Extractor Pro v13.1</h1>
        <p>Crop Region-based PDF to Excel/CSV Converter</p>
        <div class="badge-container">
            <span class="feature-badge">🎯 Crop Region Selection</span>
            <span class="feature-badge">🔍 Large Preview</span>
            <span class="feature-badge">📊 Excel/CSV Export</span>
            <span class="feature-badge">👁️ Live Preview</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Initialize session state
    if 'extracted_data' not in st.session_state:
        st.session_state.extracted_data = pd.DataFrame()
    if 'extraction_logs' not in st.session_state:
        st.session_state.extraction_logs = []
    if 'crop_coords' not in st.session_state:
        st.session_state.crop_coords = (50, 10, 95, 50)  # Default: right side, upper portion
    if 'preview_text' not in st.session_state:
        st.session_state.preview_text = ""
    if 'zoom_level' not in st.session_state:
        st.session_state.zoom_level = 2.0
    
    # Sidebar controls
    with st.sidebar:
        st.markdown("### 📐 Region Configuration")
        
        st.markdown("<div class='slider-label'>Left (X1) %</div>", unsafe_allow_html=True)
        x1 = st.slider("X1", 0, 100, st.session_state.crop_coords[0], key="x1", label_visibility="collapsed")
        
        st.markdown("<div class='slider-label'>Top (Y1) %</div>", unsafe_allow_html=True)
        y1 = st.slider("Y1", 0, 100, st.session_state.crop_coords[1], key="y1", label_visibility="collapsed")
        
        st.markdown("<div class='slider-label'>Right (X2) %</div>", unsafe_allow_html=True)
        x2 = st.slider("X2", 0, 100, st.session_state.crop_coords[2], key="x2", label_visibility="collapsed")
        
        st.markdown("<div class='slider-label'>Bottom (Y2) %</div>", unsafe_allow_html=True)
        y2 = st.slider("Y2", 0, 100, st.session_state.crop_coords[3], key="y2", label_visibility="collapsed")
        
        # Ensure valid coordinates
        if x2 <= x1:
            x2 = x1 + 5
        if y2 <= y1:
            y2 = y1 + 5
            
        current_crop = (x1, y1, x2, y2)
        st.session_state.crop_coords = current_crop
        
        st.markdown("---")
        
        st.markdown("### 🔍 Preview Settings")
        zoom_level = st.select_slider(
            "Preview Zoom Level",
            options=[1.0, 1.5, 2.0, 2.5, 3.0],
            value=st.session_state.zoom_level,
            format_func=lambda x: f"{x}x"
        )
        st.session_state.zoom_level = zoom_level
        
        st.markdown("---")
        
        st.markdown("### 📄 Page Settings")
        preview_page = st.number_input("Preview Page", 1, 1000, 1)
        start_page = st.number_input("Start Page", 1, 1000, 1)
        end_page = st.number_input("End Page (999=all)", 1, 1000, 999)
        
        st.markdown("---")
        
        st.markdown("### 🎨 Export Options")
        yellow_header = st.checkbox("Yellow Excel Header", value=True)
        
        st.markdown("---")
        
        if st.button("🗑️ Clear All Data", type="secondary"):
            st.session_state.extracted_data = pd.DataFrame()
            st.session_state.extraction_logs = []
            st.session_state.preview_text = ""
            st.rerun()
    
    # Main content
    col1, col2 = st.columns([3, 2])
    
    with col1:
        st.markdown("<div class='section-header'><h3>📤 Upload PDF</h3></div>", unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader(
            "Choose a PDF file containing BOQ tables",
            type=['pdf'],
            help="Upload a PDF with Bill of Quantities. Use the crop tool to select the table area."
        )
        
        if uploaded_file:
            st.success(f"✅ Loaded: {uploaded_file.name} ({uploaded_file.size:,} bytes)")
    
    with col2:
        st.markdown("<div class='section-header'><h3>🚀 Actions</h3></div>", unsafe_allow_html=True)
        
        if uploaded_file:
            col_a, col_b = st.columns(2)
            with col_a:
                preview_btn = st.button("👁️ Preview Region", use_container_width=True, type="secondary")
            with col_b:
                extract_btn = st.button("🔍 Extract BOQ", use_container_width=True, type="primary")
    
    # Preview Section
    if uploaded_file and preview_btn:
        st.markdown("<div class='section-header'><h3>🎯 Region Preview</h3></div>", unsafe_allow_html=True)
        
        with st.spinner("Rendering preview..."):
            img_data, error = render_pdf_page(uploaded_file, preview_page, current_crop, zoom_level)
            
            if error:
                st.error(f"Error: {error}")
            else:
                # Use full width for larger preview
                st.markdown("### 📄 Page Preview with Crop Area")
                st.image(img_data, caption=f"Page {preview_page} - Red box shows extraction area (Zoom: {zoom_level}x)", use_column_width=True)
                
                # Two columns for details and text
                col_details, col_text = st.columns([1, 1])
                
                with col_details:
                    st.markdown("### 📊 Crop Coordinates")
                    st.info(f"""
                    **Current Selection:**
                    - Left (X1): {x1}%
                    - Top (Y1): {y1}%
                    - Right (X2): {x2}%
                    - Bottom (Y2): {y2}%
                    
                    **Preview Zoom:** {zoom_level}x
                    
                    **Tips:**
                    - Adjust sliders to cover the BOQ table
                    - Use higher zoom for precise selection
                    - Red box shows what will be extracted
                    """)
                
                with col_text:
                    # Extract sample text
                    sample_text = extract_text_from_region(uploaded_file, preview_page, current_crop)
                    st.session_state.preview_text = sample_text
                    
                    if sample_text:
                        st.markdown("### 📝 Sample Extracted Text")
                        st.markdown(f"<div class='extracted-text'>{sample_text[:1500]}{'...' if len(sample_text) > 1500 else ''}</div>", unsafe_allow_html=True)
                        
                        # Parse sample
                        parser = BOQParser()
                        sample_items = parser.parse_text_block(sample_text)
                        
                        if sample_items:
                            st.success(f"✅ Found {len(sample_items)} items in preview")
                        else:
                            st.warning("⚠️ No BOQ items detected. Adjust region or check format.")
                    else:
                        st.warning("No text found in selected region")
    
    # Extraction Section
    if uploaded_file and extract_btn:
        st.markdown("<div class='section-header'><h3>⚙️ Processing</h3></div>", unsafe_allow_html=True)
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        def update_progress(progress, message):
            progress_bar.progress(progress, message)
        
        with st.spinner("Extracting BOQ data..."):
            items, logs = process_pdf_batch(uploaded_file, start_page, end_page, current_crop, update_progress)
            
            st.session_state.extraction_logs = logs
            
            if items:
                df = pd.DataFrame(items)
                st.session_state.extracted_data = df
                progress_bar.empty()
                status_text.success(f"✅ Successfully extracted {len(df)} items!")
            else:
                progress_bar.empty()
                status_text.error("❌ No items extracted. Check region settings.")
    
    # Results Section
    if not st.session_state.extracted_data.empty:
        df = st.session_state.extracted_data
        
        st.markdown("<div class='section-header'><h3>📊 Extraction Results</h3></div>", unsafe_allow_html=True)
        
        # Statistics
        c1, c2, c3, c4 = st.columns(4)
        
        with c1:
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-value">{len(df)}</div>
                <div class="stat-label">Total Items</div>
            </div>
            """, unsafe_allow_html=True)
        
        with c2:
            unique_drawings = df['Drawing No'].nunique() if 'Drawing No' in df.columns else 0
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-value">{unique_drawings}</div>
                <div class="stat-label">Drawings</div>
            </div>
            """, unsafe_allow_html=True)
        
        with c3:
            if 'Material' in df.columns:
                unique_materials = df['Material'].nunique()
            else:
                unique_materials = 0
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-value">{unique_materials}</div>
                <div class="stat-label">Materials</div>
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
        
        # Data Editor
        st.markdown("### ✏️ Review and Edit Data")
        
        edited_df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Drawing No": st.column_config.TextColumn("Drawing No", width="medium"),
                "Page": st.column_config.NumberColumn("Page", width="small"),
                "Item No": st.column_config.NumberColumn("Item No", width="small"),
                "Quantity": st.column_config.NumberColumn("Qty", width="small"),
                "Part No": st.column_config.TextColumn("Part No", width="medium"),
                "Description": st.column_config.TextColumn("Description", width="large"),
                "Material": st.column_config.TextColumn("Material", width="medium"),
            }
        )
        
        # Export options
        st.markdown("<div class='section-header'><h3>📦 Export Data</h3></div>", unsafe_allow_html=True)
        
        col_excel, col_csv = st.columns(2)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        with col_excel:
            excel_data = create_excel_download(edited_df, yellow_header)
            st.download_button(
                label="📥 Download Excel (.xlsx)",
                data=excel_data,
                file_name=f"BOQ_Extract_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col_csv:
            csv_data = edited_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📄 Download CSV (.csv)",
                data=csv_data,
                file_name=f"BOQ_Extract_{timestamp}.csv",
                mime="text/csv",
                use_container_width=True
            )
    
    # Logs
    if st.session_state.extraction_logs:
        with st.expander("🔍 Processing Logs", expanded=False):
            for log in st.session_state.extraction_logs:
                if log.startswith("✅"):
                    st.markdown(f"<span style='color: #28a745;'>{log}</span>", unsafe_allow_html=True)
                elif log.startswith("❌"):
                    st.markdown(f"<span style='color: #dc3545;'>{log}</span>", unsafe_allow_html=True)
                elif log.startswith("⚠️"):
                    st.markdown(f"<span style='color: #ffc107;'>{log}</span>", unsafe_allow_html=True)
                else:
                    st.markdown(f"<span style='color: #17a2b8;'>{log}</span>", unsafe_allow_html=True)
    
    # Instructions
    with st.expander("📖 How to Use", expanded=False):
        st.markdown("""
        ### Step-by-Step Guide

        #### 1. Upload PDF
        Upload your BOQ PDF file using the file uploader above.

        #### 2. Define Crop Region
        Use the sliders in the sidebar to define the extraction area:
        - **X1 (Left)**: Start from left edge (0-100%)
        - **Y1 (Top)**: Start from top edge (0-100%)
        - **X2 (Right)**: End at right edge (0-100%)
        - **Y2 (Bottom)**: End at bottom edge (0-100%)

        **Default**: Right side of page (50-95% width, 10-50% height)

        #### 3. Adjust Preview Zoom
        Use the "Preview Zoom Level" slider in the sidebar to get a larger view:
        - **1x**: Normal size
        - **2x**: Double size (recommended)
        - **3x**: Triple size (for precise adjustments)

        #### 4. Preview
        Click **"👁️ Preview Region"** to:
        - See the selected area highlighted in red
        - View sample extracted text
        - Verify items are detected correctly

        #### 5. Extract
        Click **"🔍 Extract BOQ"** to process all pages.

        #### 6. Review & Export
        - Edit data in the table if needed
        - Download as Excel or CSV

        ### Supported Formats
        The parser recognizes:
        - **Item No**: 1, 2, 3... (first column)
        - **Quantity**: Integer numbers
        - **Part No**: M1-V1, V2-26-BM1, PC1-2, etc.
        - **Description**: Text between Part No and Material
        - **Material**: A240 SS316, A36, Per MSS-SP58, etc.

        ### Troubleshooting
        | Problem | Solution |
        |---------|----------|
        | No items found | Adjust crop region to include entire table |
        | Missing columns | Ensure region covers all columns |
        | Wrong parsing | Check that Item No starts each row |
        | Partial data | Increase Y2 (Bottom) to capture all rows |
        | Preview too small | Increase Zoom Level in sidebar |
        """)
    
    # Footer
    st.markdown("""
    <div class="footer">
        <p>📋 BOQ Extractor Pro v13.1 | Crop Region-based Extraction</p>
        <p style="font-size: 0.8rem; color: #999;">
            Smart parsing | Adjustable preview | Batch processing
        </p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
