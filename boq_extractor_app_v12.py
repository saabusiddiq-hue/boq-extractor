"""
BOQ Extractor Pro v17.0 - REPAIRED EDITION
Fixed: Resource leaks, column indexing, error handling, dependency checks
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime
from typing import Optional, Tuple, List, Dict, Any, Callable

# Dependency check with helpful error messages
try:
    import pdfplumber
except ImportError:
    st.error("❌ Missing dependency: `pdfplumber`. Install with: `pip install pdfplumber`")
    st.stop()

try:
    import fitz  # PyMuPDF
except ImportError:
    st.error("❌ Missing dependency: `PyMuPDF`. Install with: `pip install PyMuPDF`")
    st.stop()

try:
    from PIL import Image, ImageDraw
except ImportError:
    st.error("❌ Missing dependency: `Pillow`. Install with: `pip install Pillow`")
    st.stop()

try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("❌ Missing dependency: `openpyxl`. Install with: `pip install openpyxl`")
    st.stop()

st.set_page_config(
    page_title="BOQ Extractor Pro", 
    page_icon="📋", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

def load_css():
    st.markdown("""
    <style>
        .stApp { background-color: #0d1117 !important; }
        [data-testid="stSidebar"] { background-color: #161b22 !important; border-right: 1px solid #30363d; }
        h1, h2, h3, h4, h5, h6 { color: #e6edf3 !important; }
        .section-header { 
            background: #21262d; 
            padding: 0.75rem 1rem; 
            border-radius: 8px; 
            margin: 1rem 0 0.5rem 0; 
            border-left: 3px solid #1f6feb; 
            font-size: 0.95rem; 
            font-weight: 600; 
            color: #58a6ff !important; 
        }
        .stat-card { 
            background: #161b22; 
            border-radius: 8px; 
            padding: 1rem; 
            border: 1px solid #30363d; 
            text-align: center; 
        }
        .stat-value { 
            font-size: 1.5rem; 
            font-weight: 700; 
            color: #58a6ff; 
        }
        .stat-label { 
            font-size: 0.75rem; 
            color: #8b949e; 
            text-transform: uppercase; 
            margin-top: 0.25rem; 
        }
        .extracted-text { 
            background: #0d1117; 
            color: #e6edf3; 
            padding: 1rem; 
            border-radius: 6px; 
            font-family: monospace; 
            font-size: 0.8rem; 
            max-height: 300px; 
            overflow-y: auto; 
            border: 1px solid #30363d; 
        }
        .crop-info { 
            background: #0d1117; 
            border-radius: 6px; 
            padding: 0.75rem; 
            border: 1px solid #30363d; 
            margin-top: 0.5rem; 
        }
        .crop-dimension { 
            display: flex; 
            justify-content: space-between; 
            padding: 0.2rem 0; 
            border-bottom: 1px solid #30363d; 
            color: #8b949e; 
            font-size: 0.85rem; 
        }
        .crop-dimension:last-child { border-bottom: none; }
        .stButton>button { 
            background: #238636 !important; 
            color: white !important; 
            border-radius: 6px !important; 
        }
        .stButton>button:hover { background: #2ea043 !important; }
        .stButton>button[kind="secondary"] { 
            background: #21262d !important; 
            border: 1px solid #30363d !important; 
            color: #c9d1d9 !important; 
        }
        .stSlider>div>div>div { background-color: #1f6feb !important; }
        .stFileUploader>div>div { 
            background-color: #161b22 !important; 
            border: 2px dashed #30363d !important; 
            border-radius: 6px !important; 
        }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-track { background: #161b22; }
        ::-webkit-scrollbar-thumb { background: #1f6feb; border-radius: 3px; }
        p, span, label { color: #c9d1d9 !important; }
    </style>
    """, unsafe_allow_html=True)

class SmartTableDetector:
    def __init__(self, max_items: int = 15):
        self.max_items = max_items
        
    def parse_table_content(self, text: str) -> List[Dict[str, Any]]:
        """Parse table content from extracted text."""
        if not text or not isinstance(text, str):
            return []
            
        lines = text.strip().split("\n")
        items = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            parsed = self._parse_line_smart(line)
            if parsed and parsed.get("Item No"):
                item_no = parsed["Item No"]
                if 1 <= item_no <= self.max_items:
                    items.append(parsed)
                elif item_no > self.max_items:
                    # Stop processing if we exceed max items
                    break
                    
        return items
    
    def _parse_line_smart(self, line: str) -> Optional[Dict[str, Any]]:
        """Smart parsing of a single line."""
        # Skip header lines
        skip_keywords = ["ITEM", "NO.", "REQ'D", "FIG", "DESCRIPTION", "MATERIAL", 
                        "NO", "REQUIRED", "FIGURE", "PART", "DRAWING", "MARK", "QTY"]
        
        parts = line.split()
        if not parts:
            return None
            
        first_word = parts[0].upper()
        if any(kw in first_word for kw in skip_keywords) and len(parts) <= 3:
            return None
        
        if len(parts) < 3:
            return None
        
        result = {
            "Drawing No": "",
            "Mark No": "", 
            "Item No": None,
            "Quantity": None,
            "Part No": "",
            "Description": "",
            "Material": ""
        }
        
        # Check if line starts with item number
        if parts[0].isdigit():
            item_no = int(parts[0])
            if 1 <= item_no <= self.max_items:
                result["Item No"] = item_no
                remaining = parts[1:]
                
                # Extract quantity
                if remaining and remaining[0].isdigit():
                    qty = int(remaining[0])
                    if 1 <= qty <= 9999:  # Increased max qty for flexibility
                        result["Quantity"] = qty
                        remaining = remaining[1:]
                
                # Find part number
                fig_idx = -1
                for i, part in enumerate(remaining[:4]):
                    if self._is_part_no(part):
                        result["Part No"] = part
                        fig_idx = i
                        break
                
                # Extract description and material
                if fig_idx >= 0:
                    desc_parts = []
                    mat_parts = []
                    remaining_text = remaining[fig_idx+1:]
                    
                    for i, part in enumerate(remaining_text):
                        if self._is_material(part) or (i < len(remaining_text)-1 and 
                           self._is_material(part + " " + remaining_text[i+1])):
                            mat_parts.append(part)
                        elif mat_parts:
                            mat_parts.append(part)
                        else:
                            desc_parts.append(part)
                            
                    result["Description"] = " ".join(desc_parts)
                    result["Material"] = (" ".join(mat_parts) if mat_parts 
                                        else self._extract_material_from_end(" ".join(remaining_text)))
                else:
                    result["Description"] = " ".join(remaining)
        
        return result if result["Item No"] else None
    
    def _is_part_no(self, text: str) -> bool:
        """Check if text matches part number patterns."""
        patterns = [
            r"^[A-Z]\d+[-_][A-Z]?\d+", r"^F\d+[-_]M\d+", r"^V\d+[-_]\d+[-_][A-Z]+\d*",
            r"^PC\d+[-_]\d+", r"^F\d+[-_]TS\d+", r"^PTFE", r"^Variable", r"^M\d+[:]?\d*",
            r"^3x\d+", r"^15x\d+", r"^\d+x\d+x\d+", r"^[A-Z]{2,}\d+"
        ]
        return any(re.match(p, text, re.IGNORECASE) for p in patterns)
    
    def _is_material(self, text: str) -> bool:
        """Check if text matches material patterns."""
        patterns = [
            r"A36\b", r"A105\b", r"A193", r"A194", r"MSS[-]?SP", 
            r"SS316", r"SS304", r"A516", r"A240", r"Gr[.]?\s*\d+", 
            r"CI[.]?\d+", r"PTFE", r"Graphite", r"Stainless", r"Carbon"
        ]
        return any(re.search(p, text, re.IGNORECASE) for p in patterns)
    
    def _extract_material_from_end(self, text: str) -> str:
        """Extract material specification from end of text."""
        mat_match = re.search(
            r"(A36|A105|A193\s*GR?[.]?B7|A194\s*GR?[.]?2H|Per\s*MSS[-]?SP\d+|"
            r"SS316|SS304|A516|Gr[.]?\s*\d+[.]?\d*|Graphite|Stainless|Carbon)(.*?)$", 
            text, re.IGNORECASE
        )
        if mat_match:
            return mat_match.group(0).strip()
        return ""

def validate_crop_box(crop_box: Tuple[int, int, int, int]) -> Tuple[float, float, float, float]:
    """Validate and normalize crop box coordinates."""
    x1, y1, x2, y2 = crop_box
    
    # Ensure percentages are within bounds
    x1 = max(0, min(100, float(x1)))
    y1 = max(0, min(100, float(y1)))
    x2 = max(0, min(100, float(x2)))
    y2 = max(0, min(100, float(y2)))
    
    # Ensure x2 > x1 and y2 > y1
    if x2 <= x1:
        x2 = min(x1 + 10, 100)
    if y2 <= y1:
        y2 = min(y1 + 10, 100)
        
    return (x1, y1, x2, y2)

@st.cache_data(show_spinner=False)
def render_pdf_page_with_crop(pdf_bytes: bytes, page_num: int, 
                              crop_box: Optional[Tuple[float, float, float, float]], 
                              zoom_level: float = 2.0) -> Tuple[Optional[bytes], Optional[str]]:
    """
    Render PDF page with crop overlay.
    Returns: (image_bytes, error_message)
    """
    doc = None
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        if page_num > len(doc) or page_num < 1:
            return None, f"Invalid page number {page_num}. PDF has {len(doc)} pages."
        
        page = doc[page_num - 1]
        rect = page.rect
        width_pt, height_pt = rect.width, rect.height
        
        # Draw crop box if specified
        if crop_box:
            x1_pct, y1_pct, x2_pct, y2_pct = validate_crop_box(crop_box)
            x1 = width_pt * (x1_pct / 100)
            y1 = height_pt * (y1_pct / 100)
            x2 = width_pt * (x2_pct / 100)
            y2 = height_pt * (y2_pct / 100)
            
            crop_rect = fitz.Rect(x1, y1, x2, y2)
            
            # Draw filled rectangle with transparency
            shape = page.new_shape()
            shape.draw_rect(crop_rect)
            shape.finish(color=(0.97, 0.32, 0.29), fill=(0.97, 0.32, 0.29), fill_opacity=0.15, width=3)
            shape.commit()
            
            # Add label
            label_y = max(y1 - 5, 10)  # Ensure label is within page
            label_point = fitz.Point(x1 + 5, label_y)
            page.insert_text(label_point, "EXTRACTION AREA", fontsize=14, 
                           color=(0.97, 0.32, 0.29), fontname="helv")
        
        # Render to image
        mat = fitz.Matrix(zoom_level, zoom_level)
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")
        
        return img_data, None
        
    except Exception as e:
        return None, f"PDF rendering error: {str(e)}"
    finally:
        if doc:
            doc.close()

def extract_from_region(pdf_bytes: bytes, page_num: int, 
                       crop_box: Optional[Tuple[float, float, float, float]]) -> str:
    """Extract text from specific region of PDF page."""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            if page_num > len(pdf.pages) or page_num < 1:
                return ""
                
            page = pdf.pages[page_num - 1]
            
            if crop_box:
                width, height = page.width, page.height
                x1, y1, x2, y2 = validate_crop_box(crop_box)
                crop_area = (width*x1/100, height*y1/100, width*x2/100, height*y2/100)
                page = page.crop(crop_area)
            
            text = page.extract_text()
            return text or ""
    except Exception as e:
        st.error(f"Extraction error on page {page_num}: {e}")
        return ""

def process_all_pages(pdf_bytes: bytes, 
                     crop_box: Optional[Tuple[float, float, float, float]], 
                     detector: SmartTableDetector, 
                     progress_callback: Optional[Callable[[float, str], None]] = None) -> Tuple[List[Dict], List[str]]:
    """Process all pages and extract items."""
    all_items = []
    logs = []
    
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            total_pages = len(pdf.pages)
            logs.append(f"Processing {total_pages} pages...")
            
            for page_num in range(1, total_pages + 1):
                if progress_callback:
                    progress_callback(page_num / total_pages, f"Processing page {page_num}/{total_pages}")
                
                try:
                    page = pdf.pages[page_num - 1]
                    
                    # Apply crop if specified
                    if crop_box:
                        width, height = page.width, page.height
                        x1, y1, x2, y2 = validate_crop_box(crop_box)
                        crop_area = (width*x1/100, height*y1/100, width*x2/100, height*y2/100)
                        page = page.crop(crop_area)
                    
                    text = page.extract_text() or ""
                    items = detector.parse_table_content(text)
                    
                    # Add page number to each item
                    for item in items:
                        item["Page"] = page_num
                    
                    all_items.extend(items)
                    logs.append(f"✓ Page {page_num}: {len(items)} items")
                    
                except Exception as e:
                    logs.append(f"✗ Page {page_num}: Error - {str(e)}")
                    continue
                
    except Exception as e:
        logs.append(f"Critical Error: {str(e)}")
    
    return all_items, logs

def create_excel_download(df: pd.DataFrame, yellow_header: bool = True) -> bytes:
    """Create formatted Excel file from DataFrame."""
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ Data"
    
    # Styles
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_font = Font(bold=True, size=10, color="000000")
    cell_border = Border(
        left=Side(style="thin"), 
        right=Side(style="thin"), 
        top=Side(style="thin"), 
        bottom=Side(style="thin")
    )
    
    # Reorder columns
    preferred_order = ["Drawing No", "Mark No", "Item No", "Quantity", "Part No", 
                      "Description", "Material", "Page"]
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
        cell.border = cell_border
    
    # Write data
    for row_num, row_data in enumerate(df_export.values, 2):
        for col_num, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.border = cell_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    
    # Auto-adjust column widths
    for idx in range(1, len(df_export.columns) + 1):
        col_letter = get_column_letter(idx)
        max_length = 0
        
        for row_idx in range(1, len(df_export) + 2):
            cell_value = ws.cell(row=row_idx, column=idx).value
            if cell_value:
                max_length = max(max_length, len(str(cell_value)))
        
        # Set width with limits
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[col_letter].width = max(adjusted_width, 10)
    
    ws.freeze_panes = "A2"
    wb.save(output)
    output.seek(0)
    return output.getvalue()

def initialize_session_state():
    """Initialize session state variables."""
    defaults = {
        "extracted_data": pd.DataFrame(),
        "detector": SmartTableDetector(),
        "pdf_bytes": None,
        "crop_coords": (5, 15, 95, 60),
        "filename": None
    }
    
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

def main():
    load_css()
    initialize_session_state()
    
    # Sidebar controls
    with st.sidebar:
        st.markdown("### ⚙️ Settings")
        
        max_items = st.number_input(
            "Max Items Per Page", 
            min_value=1, 
            max_value=100, 
            value=st.session_state.detector.max_items,
            help="Maximum number of items to extract per page"
        )
        st.session_state.detector.max_items = max_items
        
        st.markdown("---")
        st.markdown("**📐 Crop Region**")
        
        # Crop sliders with validation
        x1 = st.slider("Left %", 0, 100, st.session_state.crop_coords[0], key="x1")
        y1 = st.slider("Top %", 0, 100, st.session_state.crop_coords[1], key="y1")
        x2 = st.slider("Right %", 0, 100, st.session_state.crop_coords[2], key="x2")
        y2 = st.slider("Bottom %", 0, 100, st.session_state.crop_coords[3], key="y2")
        
        # Ensure valid coordinates
        if x2 <= x1:
            st.warning("⚠️ Right must be > Left. Adjusting...")
            x2 = min(x1 + 10, 100)
        if y2 <= y1:
            st.warning("⚠️ Bottom must be > Top. Adjusting...")
            y2 = min(y1 + 10, 100)
            
        current_crop = (x1, y1, x2, y2)
        st.session_state.crop_coords = current_crop
        
        # Display crop dimensions
        st.markdown('<div class="crop-info">', unsafe_allow_html=True)
        st.markdown(f'<div class="crop-dimension"><span>Width:</span><span>{x2-x1}%</span></div>', 
                   unsafe_allow_html=True)
        st.markdown(f'<div class="crop-dimension"><span>Height:</span><span>{y2-y1}%</span></div>', 
                   unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown("---")
        zoom_level = st.select_slider(
            "🔍 Preview Zoom", 
            options=[1.0, 1.5, 2.0, 2.5, 3.0], 
            value=2.0, 
            format_func=lambda x: f"{x}x"
        )
        
        st.markdown("---")
        yellow_header = st.checkbox("📊 Yellow Excel Headers", value=True)
        
        if st.button("🗑️ Clear All", type="secondary", use_container_width=True):
            for key in ["extracted_data", "pdf_bytes", "filename"]:
                st.session_state[key] = None if key != "extracted_data" else pd.DataFrame()
            st.session_state["crop_coords"] = (5, 15, 95, 60)
            st.rerun()
    
    # Main content
    col_left, col_right = st.columns([1, 1])
    
    with col_left:
        st.markdown("<div class='section-header'>📤 Upload PDF</div>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader(
            "Select PDF file", 
            type=["pdf"], 
            label_visibility="collapsed",
            help="Upload a PDF containing BOQ tables"
        )
        
        if uploaded_file:
            try:
                # Read bytes once and store in session
                pdf_bytes = uploaded_file.read()
                st.session_state.pdf_bytes = pdf_bytes
                st.session_state.filename = uploaded_file.name
                st.success(f"✓ Loaded: {uploaded_file.name} ({len(pdf_bytes):,} bytes)")
            except Exception as e:
                st.error(f"❌ Error reading file: {e}")
                st.session_state.pdf_bytes = None
    
    with col_right:
        st.markdown("<div class='section-header'>🚀 Actions</div>", unsafe_allow_html=True)
        
        if st.session_state.pdf_bytes:
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("👁️ Preview Area", use_container_width=True):
                    st.session_state.show_preview = True
            with col_btn2:
                if st.button("🔍 Extract All Pages", use_container_width=True, type="primary"):
                    st.session_state.do_extract = True
    
    # Preview section
    if st.session_state.pdf_bytes and st.session_state.get("show_preview", False):
        st.markdown("<div class='section-header'>🎯 Preview (Red Box = Extraction Area)</div>", 
                   unsafe_allow_html=True)
        
        with st.spinner("Rendering preview..."):
            img_data, error = render_pdf_page_with_crop(
                st.session_state.pdf_bytes, 
                1, 
                current_crop, 
                zoom_level
            )
            
            if error:
                st.error(f"❌ {error}")
            elif img_data:
                st.image(img_data, use_column_width=True)
                
                # Show sample text
                sample_text = extract_from_region(
                    st.session_state.pdf_bytes, 
                    1, 
                    current_crop
                )
                if sample_text:
                    with st.expander("📝 Sample Extracted Text"):
                        display_text = sample_text[:2000] + "..." if len(sample_text) > 2000 else sample_text
                        st.markdown(f'<div class="extracted-text">{display_text}</div>', 
                                  unsafe_allow_html=True)
                else:
                    st.info("ℹ️ No text found in selected region")
    
    # Extraction section
    if st.session_state.pdf_bytes and st.session_state.get("do_extract", False):
        st.markdown("<div class='section-header'>⚙️ Processing All Pages...</div>", 
                   unsafe_allow_html=True)
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        def update_progress(progress: float, message: str):
            progress_bar.progress(min(progress, 1.0))
            status_text.info(message)
        
        items, logs = process_all_pages(
            st.session_state.pdf_bytes, 
            current_crop, 
            st.session_state.detector, 
            update_progress
        )
        
        progress_bar.empty()
        status_text.empty()
        
        if items:
            df = pd.DataFrame(items)
            st.session_state.extracted_data = df
            st.success(f"✅ Successfully extracted {len(df)} items from all pages")
            
            with st.expander("📋 Processing Logs"):
                st.code("\n".join(logs), language="text")
        else:
            st.warning("⚠️ No items found. Check crop area or PDF format.")
            with st.expander("📋 Debug Logs"):
                st.code("\n".join(logs), language="text")
        
        st.session_state.do_extract = False
    
    # Results display
    if not st.session_state.extracted_data.empty:
        df = st.session_state.extracted_data
        
        st.markdown("<div class='section-header'>📊 Extraction Results</div>", 
                   unsafe_allow_html=True)
        
        # Statistics cards
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f'<div class="stat-card"><div class="stat-value">{len(df)}</div>'
                       f'<div class="stat-label">Total Items</div></div>', unsafe_allow_html=True)
        with c2:
            pages = df["Page"].nunique() if "Page" in df.columns else 1
            st.markdown(f'<div class="stat-card"><div class="stat-value">{pages}</div>'
                       f'<div class="stat-label">Pages Processed</div></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="stat-card"><div class="stat-value">{max_items}</div>'
                       f'<div class="stat-label">Max Items/Page</div></div>', unsafe_allow_html=True)
        with c4:
            materials = df["Material"].nunique() if "Material" in df.columns else 0
            st.markdown(f'<div class="stat-card"><div class="stat-value">{materials}</div>'
                       f'<div class="stat-label">Unique Materials</div></div>', unsafe_allow_html=True)
        
        # Data editor
        st.markdown("### ✏️ Review & Edit Data")
        edited_df = st.data_editor(
            df, 
            num_rows="dynamic", 
            use_container_width=True, 
            hide_index=True,
            key="data_editor",
            column_config={
                "Drawing No": st.column_config.TextColumn("Drawing No", width="small"),
                "Mark No": st.column_config.TextColumn("Mark No", width="small"),
                "Item No": st.column_config.NumberColumn("Item", min_value=1, width="small"),
                "Quantity": st.column_config.NumberColumn("Qty", min_value=0, width="small"),
                "Part No": st.column_config.TextColumn("Part No", width="medium"),
                "Description": st.column_config.TextColumn("Description", width="large"),
                "Material": st.column_config.TextColumn("Material", width="medium"),
                "Page": st.column_config.NumberColumn("Page", width="small", disabled=True),
            }
        )
        
        # Export section
        st.markdown("<div class='section-header'>📦 Export Data</div>", unsafe_allow_html=True)
        col_excel, col_csv = st.columns(2)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_filename = st.session_state.filename.replace(".pdf", "") if st.session_state.filename else "BOQ"
        
        with col_excel:
            try:
                excel_data = create_excel_download(edited_df, yellow_header)
                st.download_button(
                    label="📥 Download Excel (.xlsx)", 
                    data=excel_data, 
                    file_name=f"{base_filename}_{timestamp}.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"❌ Excel export error: {e}")
        
        with col_csv:
            csv_data = edited_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📄 Download CSV (.csv)", 
                data=csv_data, 
                file_name=f"{base_filename}_{timestamp}.csv", 
                mime="text/csv", 
                use_container_width=True
            )
    
    # Help section
    with st.expander("❓ Help & Instructions"):
        st.markdown(f"""
        **BOQ Extractor Pro v17.0**
        
        **Features:**
        - ✅ Multi-page extraction (items 1-{max_items} per page)
        - ✅ Visual crop area selection with real-time preview
        - ✅ Smart table detection (works without grid lines)
        - ✅ Export to Excel with formatted headers or CSV
        
        **How to use:**
        1. **Upload PDF** - Select your BOQ drawing/document
        2. **Adjust Crop** - Use sliders to select the table area (red box in preview)
        3. **Preview** - Check the red box covers your data correctly
        4. **Extract All** - Process all pages at once
        5. **Review & Export** - Edit data if needed and download
        
        **Tips:**
        - Ensure the crop area includes all columns but excludes headers/footers
        - If extraction fails, try adjusting the crop area slightly larger
        - Item numbers must start from 1 and be consecutive for best results
        """)

if __name__ == "__main__":
    main()
