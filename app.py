"""
BOQ Extractor Pro v18.0 - TEMPLATE LEARNING EDITION
New: Sample image upload for structure learning and pattern matching
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime
from typing import Optional, Tuple, List, Dict, Any, Callable

# Dependency checks
try:
    import pdfplumber
except ImportError:
    st.error("❌ Install pdfplumber: `pip install pdfplumber`")
    st.stop()

try:
    import fitz  # PyMuPDF
except ImportError:
    st.error("❌ Install PyMuPDF: `pip install PyMuPDF`")
    st.stop()

try:
    from PIL import Image
except ImportError:
    st.error("❌ Install Pillow: `pip install Pillow`")
    st.stop()

try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("❌ Install openpyxl: `pip install openpyxl`")
    st.stop()

# Optional OCR for sample image analysis
try:
    import pytesseract
    from PIL import ImageEnhance, ImageFilter
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False
    st.warning("⚠️ pytesseract not installed. Sample image analysis will be limited. Install with: `pip install pytesseract`")

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
            background: linear-gradient(90deg, #21262d 0%, #161b22 100%);
            padding: 0.75rem 1rem; 
            border-radius: 8px; 
            margin: 1rem 0 0.5rem 0; 
            border-left: 3px solid #1f6feb; 
            font-size: 0.95rem; 
            font-weight: 600; 
            color: #58a6ff !important; 
        }
        .sample-box { 
            background: #161b22; 
            border: 2px solid #238636; 
            border-radius: 8px; 
            padding: 1rem; 
            margin: 0.5rem 0;
        }
        .sample-header { color: #3fb950; font-weight: 600; margin-bottom: 0.5rem; }
        .stat-card { 
            background: #161b22; 
            border-radius: 8px; 
            padding: 1rem; 
            border: 1px solid #30363d; 
            text-align: center; 
        }
        .stat-value { font-size: 1.5rem; font-weight: 700; color: #58a6ff; }
        .stat-label { font-size: 0.75rem; color: #8b949e; text-transform: uppercase; margin-top: 0.25rem; }
        .extracted-text { 
            background: #0d1117; 
            color: #e6edf3; 
            padding: 1rem; 
            border-radius: 6px; 
            font-family: 'Courier New', monospace; 
            font-size: 0.85rem; 
            max-height: 300px; 
            overflow-y: auto; 
            border: 1px solid #30363d; 
            white-space: pre-wrap;
        }
        .pattern-match { background: rgba(35, 134, 54, 0.2); color: #3fb950; padding: 2px 4px; border-radius: 3px; }
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
        .stButton>button { background: #238636 !important; color: white !important; border-radius: 6px !important; }
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
        .learned-badge { 
            display: inline-block; 
            background: rgba(31, 111, 235, 0.2); 
            color: #58a6ff; 
            padding: 0.25rem 0.75rem; 
            border-radius: 12px; 
            font-size: 0.75rem; 
            margin-left: 0.5rem;
            border: 1px solid rgba(31, 111, 235, 0.4);
        }
    </style>
    """, unsafe_allow_html=True)

class StructureLearner:
    """Learns BOQ structure from sample images"""
    def __init__(self):
        self.learned_patterns = {
            "header_keywords": [],
            "column_order": [],
            "item_pattern": None,
            "part_no_pattern": None,
            "material_keywords": []
        }
        self.is_learned = False
    
    def analyze_sample_text(self, text: str):
        """Analyze OCR text from sample to learn structure"""
        lines = text.strip().split('\n')
        
        # Detect header row
        for line in lines[:5]:  # Check first few lines
            line_upper = line.upper()
            if any(kw in line_upper for kw in ["ITEM", "NO.", "DESCRIPTION", "QTY", "MATERIAL"]):
                # Split and clean header columns
                headers = [h.strip() for h in re.split(r'\s{2,}|\t', line) if h.strip()]
                self.learned_patterns["header_keywords"] = headers
                self.learned_patterns["column_order"] = self._map_headers_to_columns(headers)
                break
        
        # Detect data patterns from first few data rows
        for line in lines[5:15]:
            if line.strip() and line[0].isdigit():
                # Learn item number pattern
                self.learned_patterns["item_pattern"] = r"^\d+"
                
                # Try to detect part number pattern
                words = line.split()
                for word in words[1:4]:  # Check first few words after item no
                    if re.match(r"^[A-Z]\d+[-_]", word):
                        self.learned_patterns["part_no_pattern"] = r"^[A-Z]\d+[-_][A-Z]?\d+"
                        break
                
                break
        
        # Common material keywords
        self.learned_patterns["material_keywords"] = [
            "A36", "A105", "A193", "A194", "SS316", "SS304", 
            "A516", "GR.", "GRADE", "MSS-SP", "PTFE", "GRAPHITE"
        ]
        
        self.is_learned = True
        return self.learned_patterns
    
    def _map_headers_to_columns(self, headers: List[str]) -> List[str]:
        """Map detected headers to standard column names"""
        mapping = []
        for h in headers:
            h_upper = h.upper()
            if any(x in h_upper for x in ["ITEM", "NO."]):
                mapping.append("Item No")
            elif any(x in h_upper for x in ["DRAWING", "DWG"]):
                mapping.append("Drawing No")
            elif any(x in h_upper for x in ["MARK"]):
                mapping.append("Mark No")
            elif any(x in h_upper for x in ["QTY", "QUANTITY", "REQ"]):
                mapping.append("Quantity")
            elif any(x in h_upper for x in ["PART", "FIG", "FIG."]):
                mapping.append("Part No")
            elif any(x in h_upper for x in ["DESC", "DESCRIPTION"]):
                mapping.append("Description")
            elif any(x in h_upper for x in ["MATERIAL", "MAT'L"]):
                mapping.append("Material")
            else:
                mapping.append(h)
        return mapping

class SmartTableDetector:
    def __init__(self, max_items: int = 15, structure_learner: Optional[StructureLearner] = None):
        self.max_items = max_items
        self.learner = structure_learner or StructureLearner()
        
    def parse_table_content(self, text: str) -> List[Dict[str, Any]]:
        """Parse table content using learned patterns if available"""
        if not text or not isinstance(text, str):
            return []
            
        lines = text.strip().split("\n")
        items = []
        
        # Use learned patterns to skip header if available
        start_idx = 0
        if self.learner.is_learned:
            for i, line in enumerate(lines[:10]):
                if any(kw in line.upper() for kw in ["ITEM", "NO."]):
                    start_idx = i + 1
                    break
        
        for line in lines[start_idx:]:
            line = line.strip()
            if not line:
                continue
                
            parsed = self._parse_line_smart(line)
            if parsed and parsed.get("Item No"):
                item_no = parsed["Item No"]
                if 1 <= item_no <= self.max_items:
                    items.append(parsed)
                elif item_no > self.max_items:
                    break
                    
        return items
    
    def _parse_line_smart(self, line: str) -> Optional[Dict[str, Any]]:
        """Smart parsing with learned pattern support"""
        # Skip headers using learned keywords or defaults
        skip_keywords = ["ITEM", "NO.", "REQ'D", "FIG", "DESCRIPTION", "MATERIAL"]
        if self.learner.is_learned:
            skip_keywords = [k.upper() for k in self.learner.learned_patterns["header_keywords"]]
        
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
        
        if parts[0].isdigit():
            item_no = int(parts[0])
            if 1 <= item_no <= self.max_items:
                result["Item No"] = item_no
                remaining = parts[1:]
                
                # Extract quantity (usually second number)
                if remaining and remaining[0].isdigit():
                    qty = int(remaining[0])
                    if 1 <= qty <= 9999:
                        result["Quantity"] = qty
                        remaining = remaining[1:]
                
                # Use learned part number pattern if available
                part_pattern = (self.learner.learned_patterns.get("part_no_pattern") 
                              if self.learner.is_learned else None)
                
                fig_idx = -1
                for i, part in enumerate(remaining[:5]):
                    if part_pattern and re.match(part_pattern, part):
                        result["Part No"] = part
                        fig_idx = i
                        break
                    elif self._is_part_no(part):
                        result["Part No"] = part
                        fig_idx = i
                        break
                
                # Extract description and material
                if fig_idx >= 0:
                    remaining_text = remaining[fig_idx+1:]
                    desc_parts = []
                    mat_parts = []
                    
                    for i, part in enumerate(remaining_text):
                        if self._is_material(part):
                            mat_parts.append(part)
                        elif mat_parts:
                            mat_parts.append(part)
                        else:
                            desc_parts.append(part)
                            
                    result["Description"] = " ".join(desc_parts)
                    result["Material"] = " ".join(mat_parts) if mat_parts else ""
                else:
                    result["Description"] = " ".join(remaining)
        
        return result if result["Item No"] else None
    
    def _is_part_no(self, text: str) -> bool:
        """Check if text matches part number patterns"""
        patterns = [
            r"^[A-Z]\d+[-_][A-Z]?\d+", r"^F\d+[-_]M\d+", r"^V\d+[-_]\d+[-_][A-Z]+\d*",
            r"^PC\d+[-_]\d+", r"^F\d+[-_]TS\d+", r"^PTFE", r"^Variable", 
            r"^M\d+[:]?\d*", r"^3x\d+", r"^\d+x\d+x\d+"
        ]
        return any(re.match(p, text, re.IGNORECASE) for p in patterns)
    
    def _is_material(self, text: str) -> bool:
        """Check if text matches material patterns"""
        if self.learner.is_learned:
            keywords = self.learner.learned_patterns.get("material_keywords", [])
            if any(kw.upper() in text.upper() for kw in keywords):
                return True
                
        patterns = [
            r"A36\b", r"A105\b", r"A193", r"A194", r"MSS[-]?SP", 
            r"SS316", r"SS304", r"A516", r"A240", r"Gr[.]?\s*\d+", 
            r"CI[.]?\d+", r"PTFE", r"Graphite", r"Stainless", r"Carbon"
        ]
        return any(re.search(p, text, re.IGNORECASE) for p in patterns)

def validate_crop_box(crop_box: Tuple[int, int, int, int]) -> Tuple[float, float, float, float]:
    """Validate crop box coordinates"""
    x1, y1, x2, y2 = map(float, crop_box)
    x1, y1 = max(0, min(100, x1)), max(0, min(100, y1))
    x2, y2 = max(0, min(100, x2)), max(0, min(100, y2))
    
    if x2 <= x1:
        x2 = min(x1 + 10, 100)
    if y2 <= y1:
        y2 = min(y1 + 10, 100)
        
    return (x1, y1, x2, y2)

@st.cache_data(show_spinner=False)
def render_pdf_page_with_crop(pdf_bytes: bytes, page_num: int, 
                              crop_box: Optional[Tuple[float, float, float, float]], 
                              zoom_level: float = 2.0) -> Tuple[Optional[bytes], Optional[str]]:
    """Render PDF page with crop overlay"""
    doc = None
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        if page_num > len(doc) or page_num < 1:
            return None, f"Invalid page {page_num}. PDF has {len(doc)} pages."
        
        page = doc[page_num - 1]
        rect = page.rect
        width_pt, height_pt = rect.width, rect.height
        
        if crop_box:
            x1_pct, y1_pct, x2_pct, y2_pct = validate_crop_box(crop_box)
            x1 = width_pt * (x1_pct / 100)
            y1 = height_pt * (y1_pct / 100)
            x2 = width_pt * (x2_pct / 100)
            y2 = height_pt * (y2_pct / 100)
            
            crop_rect = fitz.Rect(x1, y1, x2, y2)
            
            shape = page.new_shape()
            shape.draw_rect(crop_rect)
            shape.finish(color=(0.97, 0.32, 0.29), fill=(0.97, 0.32, 0.29), fill_opacity=0.15, width=3)
            shape.commit()
            
            label_y = max(y1 - 5, 10)
            page.insert_text(fitz.Point(x1 + 5, label_y), "EXTRACTION AREA", 
                           fontsize=14, color=(0.97, 0.32, 0.29), fontname="helv")
        
        mat = fitz.Matrix(zoom_level, zoom_level)
        pix = page.get_pixmap(matrix=mat)
        return pix.tobytes("png"), None
        
    except Exception as e:
        return None, str(e)
    finally:
        if doc:
            doc.close()

def extract_from_region(pdf_bytes: bytes, page_num: int, 
                       crop_box: Optional[Tuple[float, float, float, float]]) -> str:
    """Extract text from PDF region"""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            if page_num > len(pdf.pages):
                return ""
                
            page = pdf.pages[page_num - 1]
            
            if crop_box:
                width, height = page.width, page.height
                x1, y1, x2, y2 = validate_crop_box(crop_box)
                crop_area = (width*x1/100, height*y1/100, width*x2/100, height*y2/100)
                page = page.crop(crop_area)
            
            return page.extract_text() or ""
    except Exception as e:
        return ""

def process_sample_image(image_bytes: bytes) -> Tuple[str, Dict]:
    """OCR and analyze sample image"""
    if not OCR_AVAILABLE:
        return "", {"error": "OCR not available"}
    
    try:
        image = Image.open(io.BytesIO(image_bytes))
        
        # Enhance image for better OCR
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(2.0)
        image = image.filter(ImageFilter.SHARPEN)
        
        # OCR with structure preservation
        text = pytesseract.image_to_string(image, config='--psm 6')
        
        return text, {"status": "success", "size": image.size}
    except Exception as e:
        return "", {"error": str(e)}

def process_all_pages(pdf_bytes: bytes, crop_box: Optional[Tuple], 
                     detector: SmartTableDetector, 
                     progress_callback: Optional[Callable] = None) -> Tuple[List, List]:
    """Process all PDF pages"""
    all_items = []
    logs = []
    
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            total_pages = len(pdf.pages)
            logs.append(f"Processing {total_pages} pages...")
            
            for page_num in range(1, total_pages + 1):
                if progress_callback:
                    progress_callback(page_num / total_pages, f"Page {page_num}/{total_pages}")
                
                try:
                    page = pdf.pages[page_num - 1]
                    
                    if crop_box:
                        w, h = page.width, page.height
                        x1, y1, x2, y2 = validate_crop_box(crop_box)
                        page = page.crop((w*x1/100, h*y1/100, w*x2/100, h*y2/100))
                    
                    text = page.extract_text() or ""
                    items = detector.parse_table_content(text)
                    
                    for item in items:
                        item["Page"] = page_num
                    
                    all_items.extend(items)
                    logs.append(f"✓ Page {page_num}: {len(items)} items")
                    
                except Exception as e:
                    logs.append(f"✗ Page {page_num}: {e}")
                    
    except Exception as e:
        logs.append(f"Error: {e}")
    
    return all_items, logs

def create_excel_download(df: pd.DataFrame, yellow_header: bool = True) -> bytes:
    """Create Excel file"""
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ Data"
    
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_font = Font(bold=True, size=10, color="000000")
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    
    # Column ordering
    preferred = ["Drawing No", "Mark No", "Item No", "Quantity", "Part No", "Description", "Material", "Page"]
    cols = [c for c in preferred if c in df.columns] + [c for c in df.columns if c not in preferred]
    df = df[cols]
    
    # Headers
    for col_num, header in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        if yellow_header:
            cell.fill = yellow_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    
    # Data
    for row_num, row in enumerate(df.values, 2):
        for col_num, value in enumerate(row, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    
    # Widths
    for idx in range(1, len(df.columns) + 1):
        col_letter = get_column_letter(idx)
        max_len = max([len(str(ws.cell(row=r, column=idx).value or "")) for r in range(1, len(df)+2)])
        ws.column_dimensions[col_letter].width = min(max_len + 2, 50)
    
    ws.freeze_panes = "A2"
    wb.save(output)
    return output.getvalue()

def initialize_session_state():
    """Init session state"""
    defaults = {
        "extracted_data": pd.DataFrame(),
        "pdf_bytes": None,
        "sample_bytes": None,
        "sample_text": "",
        "structure_learner": StructureLearner(),
        "crop_coords": (5, 15, 95, 60),
        "filename": None,
        "sample_loaded": False
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

def main():
    load_css()
    initialize_session_state()
    
    st.title("📋 BOQ Extractor Pro v18.0")
    st.markdown("*Now with Sample Learning*")
    
    # Sidebar
    with st.sidebar:
        st.markdown("### 🎯 Settings")
        
        max_items = st.number_input("Max Items/Page", 1, 100, 15)
        
        st.markdown("---")
        st.markdown("**📐 Crop Region**")
        
        x1 = st.slider("Left %", 0, 100, st.session_state.crop_coords[0])
        y1 = st.slider("Top %", 0, 100, st.session_state.crop_coords[1])
        x2 = st.slider("Right %", 0, 100, st.session_state.crop_coords[2])
        y2 = st.slider("Bottom %", 0, 100, st.session_state.crop_coords[3])
        
        if x2 <= x1: x2 = min(x1 + 10, 100)
        if y2 <= y1: y2 = min(y1 + 10, 100)
        
        current_crop = (x1, y1, x2, y2)
        st.session_state.crop_coords = current_crop
        
        st.markdown(f'<div class="crop-info"><div class="crop-dimension"><span>Width:</span><span>{x2-x1}%</span></div><div class="crop-dimension"><span>Height:</span><span>{y2-y1}%</span></div></div>', unsafe_allow_html=True)
        
        zoom_level = st.select_slider("Zoom", [1.0, 1.5, 2.0, 2.5, 3.0], 2.0, format_func=lambda x: f"{x}x")
        yellow_header = st.checkbox("Yellow Headers", True)
        
        if st.button("🗑️ Clear All", type="secondary"):
            for k in ["extracted_data", "pdf_bytes", "sample_bytes", "sample_text", "sample_loaded"]:
                st.session_state[k] = None if k not in ["extracted_data", "sample_text"] else (pd.DataFrame() if k == "extracted_data" else "")
            st.session_state["structure_learner"] = StructureLearner()
            st.rerun()
    
    # Main layout: Two columns for uploads
    col_sample, col_target = st.columns(2)
    
    # Sample Upload Section
    with col_sample:
        st.markdown("<div class='section-header'>📸 Upload Sample/Reference</div>", unsafe_allow_html=True)
        st.markdown("*Upload a clear image showing the BOQ format to learn from*")
        
        sample_file = st.file_uploader(
            "Sample Image (PNG/JPG)", 
            type=["png", "jpg", "jpeg"],
            key="sample_upload",
            help="Upload a reference image of the BOQ format. The app will learn the column structure from this."
        )
        
        if sample_file:
            try:
                sample_bytes = sample_file.read()
                st.session_state.sample_bytes = sample_bytes
                
                st.success(f"✓ Sample loaded: {sample_file.name}")
                st.image(sample_bytes, caption="Sample Reference", use_column_width=True)
                
                if OCR_AVAILABLE:
                    if st.button("🔍 Analyze Sample Structure", key="analyze_sample"):
                        with st.spinner("Learning from sample..."):
                            text, info = process_sample_image(sample_bytes)
                            if text:
                                st.session_state.sample_text = text
                                st.session_state.structure_learner.analyze_sample_text(text)
                                st.session_state.sample_loaded = True
                                st.success("✅ Structure learned from sample!")
                                
                                with st.expander("View Detected Structure"):
                                    st.json(st.session_state.structure_learner.learned_patterns)
                                    st.markdown("**Raw OCR Text:**")
                                    st.markdown(f'<div class="extracted-text">{text[:1000]}...</div>', unsafe_allow_html=True)
                            else:
                                st.error("Failed to extract text from sample")
                else:
                    st.info("ℹ️ Install pytesseract for automatic structure learning")
                    
            except Exception as e:
                st.error(f"Error loading sample: {e}")
    
    # Target PDF Upload Section
    with col_target:
        st.markdown("<div class='section-header'>📄 Upload Target PDF</div>", unsafe_allow_html=True)
        
        if st.session_state.sample_loaded:
            st.markdown('<span class="learned-badge">✓ Using Learned Structure</span>', unsafe_allow_html=True)
        
        pdf_file = st.file_uploader(
            "Target PDF", 
            type=["pdf"],
            key="pdf_upload",
            help="Upload the PDF to extract BOQ data from"
        )
        
        if pdf_file:
            try:
                pdf_bytes = pdf_file.read()
                st.session_state.pdf_bytes = pdf_bytes
                st.session_state.filename = pdf_file.name
                st.success(f"✓ PDF loaded: {pdf_file.name}")
                
                # Quick preview
                img_data, err = render_pdf_page_with_crop(pdf_bytes, 1, current_crop, zoom_level)
                if img_data:
                    st.image(img_data, caption="Page 1 Preview with Crop Area", use_column_width=True)
                elif err:
                    st.error(err)
                    
            except Exception as e:
                st.error(f"Error loading PDF: {e}")
    
    # Action buttons
    if st.session_state.pdf_bytes:
        st.markdown("<div class='section-header'>🚀 Extraction Actions</div>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        
        with c1:
            if st.button("👁️ Preview Crop Area", use_container_width=True):
                st.session_state.show_preview = True
        
        with c2:
            if st.button("🔍 Extract Page 1 Only", use_container_width=True):
                st.session_state.extract_single = True
                
        with c3:
            if st.button("⚡ Extract All Pages", use_container_width=True, type="primary"):
                st.session_state.extract_all = True
    
    # Create detector with learned structure
    detector = SmartTableDetector(
        max_items=max_items,
        structure_learner=st.session_state.structure_learner if st.session_state.sample_loaded else StructureLearner()
    )
    
    # Single page extraction
    if st.session_state.get("extract_single", False):
        st.markdown("### 📄 Single Page Results")
        text = extract_from_region(st.session_state.pdf_bytes, 1, current_crop)
        
        if text:
            items = detector.parse_table_content(text)
            if items:
                df_single = pd.DataFrame(items)
                st.dataframe(df_single, use_container_width=True)
            else:
                st.warning("No items found on page 1")
                with st.expander("Debug: Extracted Text"):
                    st.text(text)
        st.session_state.extract_single = False
    
    # All pages extraction
    if st.session_state.get("extract_all", False):
        st.markdown("<div class='section-header'>⚙️ Processing All Pages...</div>", unsafe_allow_html=True)
        
        progress_bar = st.progress(0)
        status = st.empty()
        
        items, logs = process_all_pages(
            st.session_state.pdf_bytes,
            current_crop,
            detector,
            lambda p, m: (progress_bar.progress(min(p, 1.0)), status.info(m))
        )
        
        progress_bar.empty()
        status.empty()
        
        if items:
            st.session_state.extracted_data = pd.DataFrame(items)
            st.success(f"✅ Extracted {len(items)} items")
        else:
            st.warning("No items found")
            
        with st.expander("Processing Logs"):
            st.code("\n".join(logs))
            
        st.session_state.extract_all = False
    
    # Display final results
    if not st.session_state.extracted_data.empty:
        df = st.session_state.extracted_data
        
        st.markdown("<div class='section-header'>📊 Final Results</div>", unsafe_allow_html=True)
        
        # Stats
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Items", len(df))
        c2.metric("Pages", df["Page"].nunique() if "Page" in df.columns else 1)
        c3.metric("Materials", df["Material"].nunique() if "Material" in df.columns else 0)
        c4.metric("Using Sample", "Yes" if st.session_state.sample_loaded else "No")
        
        # Editor
        st.markdown("### ✏️ Edit Data")
        edited_df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Drawing No": st.column_config.TextColumn("Drawing No"),
                "Mark No": st.column_config.TextColumn("Mark No"),
                "Item No": st.column_config.NumberColumn("Item"),
                "Quantity": st.column_config.NumberColumn("Qty"),
                "Part No": st.column_config.TextColumn("Part No"),
                "Description": st.column_config.TextColumn("Description", width="large"),
                "Material": st.column_config.TextColumn("Material"),
                "Page": st.column_config.NumberColumn("Page", disabled=True),
            }
        )
        
        # Export
        st.markdown("<div class='section-header'>📦 Export</div>", unsafe_allow_html=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = (st.session_state.filename or "BOQ").replace(".pdf", "")
        
        c1, c2 = st.columns(2)
        with c1:
            excel_data = create_excel_download(edited_df, yellow_header)
            st.download_button(
                "📥 Download Excel",
                excel_data,
                f"{base}_{ts}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with c2:
            csv_data = edited_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📄 Download CSV",
                csv_data,
                f"{base}_{ts}.csv",
                "text/csv",
                use_container_width=True
            )

    # Help
    with st.expander("❓ How to Use Sample Learning"):
        st.markdown("""
        **1. Upload Sample Image** (Left side)
        - Take a screenshot or photo of a correctly formatted BOQ table
        - Upload it to teach the app your specific format
        - Click "Analyze Sample Structure" to learn column positions
        
        **2. Upload Target PDF** (Right side)
        - Upload the PDF containing multiple pages to extract
        - The red crop box shows the extraction area
        
        **3. Extract**
        - The app uses the learned structure from the sample to parse the PDF
        - Edit results if needed and export to Excel
        
        **Benefits:**
        - Handles custom column orders
        - Learns part number patterns
        - Adapts to different header names
        - Better accuracy for non-standard formats
        """)

if __name__ == "__main__":
    main()
