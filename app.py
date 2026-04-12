"""
BOQ Extractor Pro v19.0 - MATERIAL REVIEW EDITION
New: Material validation preview with inline correction
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import shutil
from datetime import datetime
from typing import Optional, Tuple, List, Dict, Any, Callable

# Auto-detect tesseract path
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
        .material-card {
            background: #161b22;
            border: 1px solid #30363d;
            border-radius: 8px;
            padding: 1rem;
            margin: 0.5rem 0;
        }
        .material-item {
            background: #0d1117;
            border-left: 3px solid #238636;
            padding: 0.75rem;
            margin: 0.5rem 0;
            border-radius: 0 6px 6px 0;
        }
        .material-merged {
            border-left-color: #f85149 !important;
            background: rgba(248, 81, 73, 0.1) !important;
        }
        .material-preview {
            font-family: 'Courier New', monospace;
            font-size: 0.9rem;
            color: #e6edf3;
        }
        .extracted-text { 
            background: #0d1117; 
            color: #e6edf3; 
            padding: 1rem; 
            border-radius: 6px; 
            font-family: monospace; 
            font-size: 0.85rem; 
            max-height: 300px; 
            overflow-y: auto; 
            border: 1px solid #30363d; 
            white-space: pre-wrap;
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
        .stButton>button { background: #238636 !important; color: white !important; border-radius: 6px !important; }
        .stButton>button:hover { background: #2ea043 !important; }
        .stButton>button[kind="secondary"] { 
            background: #21262d !important; 
            border: 1px solid #30363d !important; 
            color: #c9d1d9 !important; 
        }
        .stSlider>div>div>div { background-color: #1f6feb !important; }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-track { background: #161b22; }
        ::-webkit-scrollbar-thumb { background: #1f6feb; border-radius: 3px; }
        p, span, label { color: #c9d1d9 !important; }
        .badge {
            display: inline-block;
            padding: 0.2rem 0.6rem;
            border-radius: 12px;
            font-size: 0.75rem;
            font-weight: 600;
            margin-left: 0.5rem;
        }
        .badge-warning { background: rgba(248, 81, 73, 0.2); color: #f85149; border: 1px solid rgba(248, 81, 73, 0.4); }
        .badge-success { background: rgba(35, 134, 54, 0.2); color: #3fb950; border: 1px solid rgba(35, 134, 54, 0.4); }
        .split-view {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 1rem;
            background: #161b22;
            padding: 1rem;
            border-radius: 8px;
            border: 1px solid #30363d;
        }
    </style>
    """, unsafe_allow_html=True)

# Common material standards for validation
STANDARD_MATERIALS = [
    "A36", "A105", "A193 GR.B7", "A194 GR.2H", "A240", "A516",
    "SS304", "SS316", "SS316L", "SS304L",
    "Graphite", "PTFE", "Bronze", "Graphite Bronze",
    "Carbon Steel", "Stainless Steel", "Cast Iron", "CI",
    "MSS-SP", "Per MSS-SP", "Galvanized"
]

class MaterialValidator:
    """Detects and validates materials, identifies merged descriptions"""
    
    def __init__(self):
        self.detected_patterns = []
    
    def extract_materials_from_text(self, text: str) -> List[Dict]:
        """Extract potential materials and check for merging issues"""
        materials_found = []
        
        for std_mat in STANDARD_MATERIALS:
            if std_mat.upper() in text.upper():
                # Check if material is at the start (merged with description)
                pattern = rf"({re.escape(std_mat)})\s+(.+)"
                matches = re.finditer(pattern, text, re.IGNORECASE)
                for match in matches:
                    materials_found.append({
                        "material": match.group(1),
                        "rest": match.group(2),
                        "full_text": text,
                        "is_merged": True,  # Material at start = likely merged
                        "position": "start"
                    })
        
        # Also check for materials at end (correct position)
        for std_mat in STANDARD_MATERIALS:
            if text.upper().endswith(std_mat.upper()):
                materials_found.append({
                    "material": std_mat,
                    "rest": "",
                    "full_text": text,
                    "is_merged": False,
                    "position": "end"
                })
        
        return materials_found
    
    def suggest_correction(self, description: str, current_material: str) -> Dict:
        """Suggest material/description split if merged"""
        result = {
            "original_desc": description,
            "original_mat": current_material,
            "suggested_desc": description,
            "suggested_mat": current_material,
            "confidence": "high",
            "issue": None
        }
        
        # Check if description starts with material
        for std_mat in STANDARD_MATERIALS:
            pattern = rf"^({re.escape(std_mat)})\s*[/-]?\s*(.+)"
            match = re.match(pattern, description, re.IGNORECASE)
            if match:
                result["suggested_mat"] = match.group(1)
                result["suggested_desc"] = match.group(2).strip()
                result["issue"] = "material_merged_with_description"
                result["confidence"] = "medium"
                break
        
        # Check if material contains description (reverse merge)
        if current_material and len(current_material) > 20:
            for std_mat in STANDARD_MATERIALS:
                if std_mat.upper() in current_material.upper():
                    # Extract just the material part
                    idx = current_material.upper().find(std_mat.upper())
                    result["suggested_mat"] = current_material[idx:idx+len(std_mat)]
                    result["suggested_desc"] = description + " " + current_material[:idx].strip()
                    result["issue"] = "description_merged_with_material"
                    result["confidence"] = "low"
                    break
        
        return result

class SmartTableDetector:
    def __init__(self, max_items: int = 15):
        self.max_items = max_items
        self.material_validator = MaterialValidator()
        
    def parse_table_content(self, text: str) -> List[Dict[str, Any]]:
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
                    # Check for material issues
                    validation = self.material_validator.suggest_correction(
                        parsed.get("Description", ""),
                        parsed.get("Material", "")
                    )
                    parsed["_validation"] = validation
                    items.append(parsed)
                elif item_no > self.max_items:
                    break
                    
        return items
    
    def _parse_line_smart(self, line: str) -> Optional[Dict[str, Any]]:
        skip_keywords = ["ITEM", "NO.", "REQ'D", "FIG", "DESCRIPTION", "MATERIAL"]
        
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
                
                if remaining and remaining[0].isdigit():
                    qty = int(remaining[0])
                    if 1 <= qty <= 9999:
                        result["Quantity"] = qty
                        remaining = remaining[1:]
                
                fig_idx = -1
                for i, part in enumerate(remaining[:5]):
                    if self._is_part_no(part):
                        result["Part No"] = part
                        fig_idx = i
                        break
                
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
        patterns = [
            r"^[A-Z]\d+[-_][A-Z]?\d+", r"^F\d+[-_]M\d+", r"^V\d+[-_]\d+[-_][A-Z]+\d*",
            r"^PC\d+[-_]\d+", r"^F\d+[-_]TS\d+", r"^PTFE", r"^Variable", 
            r"^M\d+[:]?\d*", r"^3x\d+", r"^\d+x\d+x\d+"
        ]
        return any(re.match(p, text, re.IGNORECASE) for p in patterns)
    
    def _is_material(self, text: str) -> bool:
        patterns = [
            r"A36\b", r"A105\b", r"A193", r"A194", r"MSS[-]?SP", 
            r"SS316", r"SS304", r"A516", r"A240", r"Gr[.]?\s*\d+", 
            r"CI[.]?\d+", r"PTFE", r"Graphite", r"Bronze", r"Stainless", r"Carbon"
        ]
        return any(re.search(p, text, re.IGNORECASE) for p in patterns)

def validate_crop_box(crop_box: Tuple) -> Tuple[float, float, float, float]:
    x1, y1, x2, y2 = map(float, crop_box)
    x1, y1 = max(0, min(100, x1)), max(0, min(100, y1))
    x2, y2 = max(0, min(100, x2)), max(0, min(100, y2))
    if x2 <= x1: x2 = min(x1 + 10, 100)
    if y2 <= y1: y2 = min(y1 + 10, 100)
    return (x1, y1, x2, y2)

@st.cache_data(show_spinner=False)
def render_pdf_page_with_crop(pdf_bytes: bytes, page_num: int, 
                              crop_box: Optional[Tuple], 
                              zoom_level: float = 2.0) -> Tuple[Optional[bytes], Optional[str]]:
    doc = None
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if page_num > len(doc) or page_num < 1:
            return None, f"Invalid page {page_num}"
        
        page = doc[page_num - 1]
        rect = page.rect
        
        if crop_box:
            x1_pct, y1_pct, x2_pct, y2_pct = validate_crop_box(crop_box)
            x1 = rect.width * (x1_pct / 100)
            y1 = rect.height * (y1_pct / 100)
            x2 = rect.width * (x2_pct / 100)
            y2 = rect.height * (y2_pct / 100)
            
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

def extract_from_region(pdf_bytes: bytes, page_num: int, crop_box: Optional[Tuple]) -> str:
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            if page_num > len(pdf.pages):
                return ""
            page = pdf.pages[page_num - 1]
            if crop_box:
                w, h = page.width, page.height
                x1, y1, x2, y2 = validate_crop_box(crop_box)
                page = page.crop((w*x1/100, h*y1/100, w*x2/100, h*y2/100))
            return page.extract_text() or ""
    except Exception as e:
        return ""

def process_all_pages(pdf_bytes: bytes, crop_box: Optional[Tuple], 
                     detector: SmartTableDetector, 
                     progress_callback: Optional[Callable] = None) -> Tuple[List, List]:
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
    
    preferred = ["Drawing No", "Mark No", "Item No", "Quantity", "Part No", "Description", "Material", "Page"]
    cols = [c for c in preferred if c in df.columns] + [c for c in df.columns if c not in preferred]
    df = df[cols]
    
    for col_num, header in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        if yellow_header:
            cell.fill = yellow_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
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

def initialize_session_state():
    defaults = {
        "extracted_data": pd.DataFrame(),
        "pdf_bytes": None,
        "crop_coords": (5, 15, 95, 60),
        "filename": None,
        "material_corrections": {},  # Store user corrections
        "show_material_review": False,
        "corrected_data": None
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

def main():
    load_css()
    initialize_session_state()
    
    st.title("📋 BOQ Extractor Pro v19.0")
    st.markdown("*With Material Review & Correction*")
    
    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Settings")
        
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
            for k in ["extracted_data", "pdf_bytes", "corrected_data", "material_corrections", "show_material_review"]:
                st.session_state[k] = None if k not in ["extracted_data", "corrected_data"] else pd.DataFrame()
            st.session_state["crop_coords"] = (5, 15, 95, 60)
            st.rerun()
    
    # Upload Section
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("<div class='section-header'>📄 Upload Target PDF</div>", unsafe_allow_html=True)
        pdf_file = st.file_uploader("Select PDF", type=["pdf"], key="pdf_upload")
        
        if pdf_file:
            try:
                pdf_bytes = pdf_file.read()
                st.session_state.pdf_bytes = pdf_bytes
                st.session_state.filename = pdf_file.name
                st.success(f"✓ {pdf_file.name}")
                
                img_data, err = render_pdf_page_with_crop(pdf_bytes, 1, current_crop, zoom_level)
                if img_data:
                    st.image(img_data, caption="Preview with Crop Area", use_column_width=True)
                    
            except Exception as e:
                st.error(f"Error: {e}")
    
    with col2:
        st.markdown("<div class='section-header'>🚀 Actions</div>", unsafe_allow_html=True)
        
        if st.session_state.pdf_bytes:
            if st.button("⚡ Extract & Analyze Materials", use_container_width=True, type="primary"):
                st.session_state.extract_all = True
    
    # Extraction
    if st.session_state.get("extract_all", False):
        detector = SmartTableDetector(max_items=max_items)
        
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
            df = pd.DataFrame(items)
            st.session_state.extracted_data = df
            st.session_state.show_material_review = True
            st.success(f"✅ Extracted {len(items)} items. Review materials below.")
        else:
            st.warning("No items found")
            
        st.session_state.extract_all = False
    
    # 🎯 MATERIAL REVIEW SECTION
    if st.session_state.show_material_review and not st.session_state.extracted_data.empty:
        df = st.session_state.extracted_data
        
        st.markdown("<div class='section-header'>🔍 Material Review & Correction</div>", unsafe_allow_html=True)
        st.info("👇 Review detected materials. If 'Graphite bronze' merged with description, fix it here before finalizing.")
        
        # Find problematic materials
        issues_found = []
        for idx, row in df.iterrows():
            validation = row.get("_validation", {})
            if validation.get("issue"):
                issues_found.append({
                    "index": idx,
                    "item_no": row.get("Item No"),
                    "part_no": row.get("Part No"),
                    "current_desc": row.get("Description"),
                    "current_mat": row.get("Material"),
                    "suggested_desc": validation.get("suggested_desc"),
                    "suggested_mat": validation.get("suggested_mat"),
                    "issue": validation.get("issue")
                })
        
        # Show summary
        if issues_found:
            st.warning(f"⚠️ Found {len(issues_found)} potential material issues (e.g., merged with description)")
        else:
            st.success("✅ All materials look good! No merging issues detected.")
        
        # Material correction interface
        corrected_data = []
        
        for idx, row in df.iterrows():
            validation = row.get("_validation", {})
            has_issue = validation.get("issue") is not None
            
            with st.container():
                cols = st.columns([0.5, 1, 1.5, 1.5, 0.5])
                
                # Item info
                with cols[0]:
                    st.markdown(f"**#{int(row.get('Item No', idx))}**")
                    if has_issue:
                        st.markdown('<span class="badge badge-warning">Fix Needed</span>', unsafe_allow_html=True)
                
                # Part No
                with cols[1]:
                    st.text(row.get("Part No", "-"))
                
                # Description (editable)
                with cols[2]:
                    current_desc = row.get("Description", "")
                    # Apply any previous correction
                    if idx in st.session_state.material_corrections:
                        current_desc = st.session_state.material_corrections[idx].get("Description", current_desc)
                    
                    new_desc = st.text_area(
                        "Description",
                        value=current_desc,
                        key=f"desc_{idx}",
                        height=80,
                        label_visibility="collapsed"
                    )
                
                # Material (editable)
                with cols[3]:
                    current_mat = row.get("Material", "")
                    # Apply suggestion if available and not yet corrected
                    if has_issue and idx not in st.session_state.material_corrections:
                        current_mat = validation.get("suggested_mat", current_mat)
                    elif idx in st.session_state.material_corrections:
                        current_mat = st.session_state.material_corrections[idx].get("Material", current_mat)
                    
                    new_mat = st.text_input(
                        "Material",
                        value=current_mat,
                        key=f"mat_{idx}",
                        label_visibility="collapsed"
                    )
                    
                    # Show suggestion hint
                    if has_issue and idx not in st.session_state.material_corrections:
                        st.caption(f"💡 Was: '{validation.get('original_mat')}' in desc")
                
                # Apply button
                with cols[4]:
                    if st.button("✓", key=f"apply_{idx}"):
                        st.session_state.material_corrections[idx] = {
                            "Description": new_desc,
                            "Material": new_mat
                        }
                        st.rerun()
                
                # Divider
                st.markdown("---")
                
                # Build corrected row
                corrected_row = row.to_dict()
                corrected_row["Description"] = new_desc
                corrected_row["Material"] = new_mat
                # Remove validation metadata from final output
                corrected_row.pop("_validation", None)
                corrected_data.append(corrected_row)
        
        # Bulk actions
        col_bulk1, col_bulk2, col_bulk3 = st.columns(3)
        
        with col_bulk1:
            if st.button("🤖 Auto-Fix All Suggestions", use_container_width=True):
                for issue in issues_found:
                    idx = issue["index"]
                    st.session_state.material_corrections[idx] = {
                        "Description": issue["suggested_desc"],
                        "Material": issue["suggested_mat"]
                    }
                st.rerun()
        
        with col_bulk2:
            if st.button("✅ Finalize to BOQ", use_container_width=True, type="primary"):
                # Apply all corrections to dataframe
                final_df = pd.DataFrame(corrected_data)
                for idx, correction in st.session_state.material_corrections.items():
                    if idx < len(final_df):
                        final_df.at[idx, "Description"] = correction["Description"]
                        final_df.at[idx, "Material"] = correction["Material"]
                
                # Remove internal columns
                final_df = final_df.drop(columns=[col for col in final_df.columns if col.startswith("_")], errors="ignore")
                
                st.session_state.corrected_data = final_df
                st.session_state.show_material_review = False
                st.success("✅ Materials corrected! Scroll down to see final BOQ.")
                st.rerun()
        
        with col_bulk3:
            if st.button("⏭️ Skip Review (Use As-Is)", use_container_width=True):
                final_df = pd.DataFrame(corrected_data)
                final_df = final_df.drop(columns=[col for col in final_df.columns if col.startswith("_")], errors="ignore")
                st.session_state.corrected_data = final_df
                st.session_state.show_material_review = False
                st.rerun()
    
    # Final Results Display
    if st.session_state.corrected_data is not None and not st.session_state.corrected_data.empty:
        final_df = st.session_state.corrected_data
        
        st.markdown("<div class='section-header'>📊 Final BOQ Data</div>", unsafe_allow_html=True)
        
        # Stats
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Items", len(final_df))
        c2.metric("Pages", final_df["Page"].nunique() if "Page" in final_df.columns else 1)
        c3.metric("Materials", final_df["Material"].nunique() if "Material" in final_df.columns else 0)
        
        # Editable final table
        st.markdown("### ✏️ Final Edit (if needed)")
        edited_df = st.data_editor(
            final_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="final_editor"
        )
        
        # Export
        st.markdown("<div class='section-header'>📦 Export</div>", unsafe_allow_html=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = (st.session_state.filename or "BOQ").replace(".pdf", "")
        
        col_excel, col_csv = st.columns(2)
        with col_excel:
            excel_data = create_excel_download(edited_df, yellow_header)
            st.download_button(
                "📥 Download Excel",
                excel_data,
                f"{base}_{ts}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_csv:
            csv_data = edited_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📄 Download CSV",
                csv_data,
                f"{base}_{ts}.csv",
                "text/csv",
                use_container_width=True
            )

if __name__ == "__main__":
    main()
