"""
BOQ Extractor Pro v16.0 - FINAL EDITION
Clean Layout | Red Crop Preview | Smart Detection | Multi-Page Support
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

st.set_page_config(page_title="BOQ Extractor Pro", page_icon="📋", layout="wide", initial_sidebar_state="expanded")

def load_css():
    st.markdown("""
    <style>
        .stApp { background-color: #0d1117 !important; }
        [data-testid="stSidebar"] { background-color: #161b22 !important; border-right: 1px solid #30363d; }
        h1, h2, h3, h4, h5, h6 { color: #e6edf3 !important; }
        .section-header { background: #21262d; padding: 0.75rem 1rem; border-radius: 8px; margin: 1rem 0 0.5rem 0; border-left: 3px solid #1f6feb; font-size: 0.95rem; font-weight: 600; color: #58a6ff !important; }
        .stat-card { background: #161b22; border-radius: 8px; padding: 1rem; border: 1px solid #30363d; text-align: center; }
        .stat-value { font-size: 1.5rem; font-weight: 700; color: #58a6ff; }
        .stat-label { font-size: 0.75rem; color: #8b949e; text-transform: uppercase; margin-top: 0.25rem; }
        .extracted-text { background: #0d1117; color: #e6edf3; padding: 1rem; border-radius: 6px; font-family: monospace; font-size: 0.8rem; max-height: 300px; overflow-y: auto; border: 1px solid #30363d; }
        .slider-label { font-weight: 500; color: #c9d1d9; margin-bottom: 0.25rem; font-size: 0.85rem; }
        .crop-info { background: #0d1117; border-radius: 6px; padding: 0.75rem; border: 1px solid #30363d; margin-top: 0.5rem; }
        .crop-dimension { display: flex; justify-content: space-between; padding: 0.2rem 0; border-bottom: 1px solid #30363d; color: #8b949e; font-size: 0.85rem; }
        .crop-dimension:last-child { border-bottom: none; }
        .stButton>button { background: #238636 !important; color: white !important; border-radius: 6px !important; }
        .stButton>button:hover { background: #2ea043 !important; }
        .stButton>button[kind="secondary"] { background: #21262d !important; border: 1px solid #30363d !important; color: #c9d1d9 !important; }
        .stSlider>div>div>div { background-color: #1f6feb !important; }
        .stFileUploader>div>div { background-color: #161b22 !important; border: 2px dashed #30363d !important; border-radius: 6px !important; }
        .stSuccess { background-color: rgba(35, 134, 54, 0.1) !important; border: 1px solid #238636 !important; color: #3fb950 !important; padding: 0.5rem !important; border-radius: 6px !important; }
        .stInfo { background-color: rgba(56, 139, 253, 0.1) !important; border: 1px solid #388bfd !important; color: #58a6ff !important; padding: 0.75rem !important; border-radius: 6px !important; }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-track { background: #161b22; }
        ::-webkit-scrollbar-thumb { background: #1f6feb; border-radius: 3px; }
        p, span, label { color: #c9d1d9 !important; }
    </style>
    """, unsafe_allow_html=True)

class SmartTableDetector:
    def __init__(self):
        self.max_items = 15
        
    def parse_table_content(self, text):
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
                    break
        return items
    
    def _parse_line_smart(self, line):
        skip_keywords = ["ITEM", "NO.", "REQ'D", "FIG", "DESCRIPTION", "MATERIAL", "NO", "REQUIRED", "FIGURE", "PART", "DRAWING", "MARK", "QTY"]
        first_word = line.split()[0].upper() if line.split() else ""
        if any(kw in first_word for kw in skip_keywords) and len(line.split()) <= 3:
            return None
        
        parts = line.split()
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
            if 1 <= item_no <= 15:
                result["Item No"] = item_no
                remaining = parts[1:]
                
                if remaining and remaining[0].isdigit():
                    qty = int(remaining[0])
                    if 1 <= qty <= 20:
                        result["Quantity"] = qty
                        remaining = remaining[1:]
                
                fig_idx = -1
                for i, part in enumerate(remaining[:4]):
                    if self._is_part_no(part):
                        result["Part No"] = part
                        fig_idx = i
                        break
                
                if fig_idx >= 0:
                    desc_parts = []
                    mat_parts = []
                    for i, part in enumerate(remaining[fig_idx+1:], start=fig_idx+1):
                        if self._is_material(part) or (i < len(remaining)-1 and self._is_material(part + " " + remaining[i+1])):
                            mat_parts.append(part)
                        elif mat_parts:
                            mat_parts.append(part)
                        else:
                            desc_parts.append(part)
                    result["Description"] = " ".join(desc_parts)
                    result["Material"] = " ".join(mat_parts) if mat_parts else self._extract_material_from_end(" ".join(remaining[fig_idx+1:]))
                else:
                    result["Description"] = " ".join(remaining)
        
        return result if result["Item No"] else None
    
    def _is_part_no(self, text):
        patterns = [
            r"^[A-Z]\d+[-_][A-Z]?\d+", r"^F\d+[-_]M\d+", r"^V\d+[-_]\d+[-_][A-Z]+\d*",
            r"^PC\d+[-_]\d+", r"^F\d+[-_]TS\d+", r"^PTFE", r"^Variable", r"^M\d+[:]?\d*",
            r"^3x\d+", r"^15x\d+", r"^\d+x\d+x\d+"
        ]
        return any(re.match(p, text, re.IGNORECASE) for p in patterns)
    
    def _is_material(self, text):
        patterns = [r"A36\b", r"A105\b", r"A193", r"A194", r"MSS[-]?SP", r"SS316", r"SS304", r"A516", r"A240", r"Gr[.]?\s*\d+", r"CI[.]?\s*\d+", r"PTFE", r"Graphite"]
        return any(re.search(p, text, re.IGNORECASE) for p in patterns)
    
    def _extract_material_from_end(self, text):
        mat_match = re.search(r"(A36|A105|A193\s*GR?[.]?B7|A194\s*GR?[.]?2H|Per\s*MSS[-]?SP\d+|SS316|SS304|A516|Gr[.]?\s*\d+[.]?\d*|Graphite)(.*?)$", text, re.IGNORECASE)
        if mat_match:
            return mat_match.group(0).strip()
        return ""

def render_pdf_page_with_crop(pdf_file, page_num, crop_box, zoom_level=2.0):
    try:
        pdf_file.seek(0)
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        if page_num > len(doc):
            doc.close()
            return None, "Invalid page"
        
        page = doc[page_num - 1]
        rect = page.rect
        width_pt, height_pt = rect.width, rect.height
        
        if crop_box:
            x1_pct, y1_pct, x2_pct, y2_pct = crop_box
            x1 = width_pt * (x1_pct / 100)
            y1 = height_pt * (y1_pct / 100)
            x2 = width_pt * (x2_pct / 100)
            y2 = height_pt * (y2_pct / 100)
            
            crop_rect = fitz.Rect(x1, y1, x2, y2)
            page.draw_rect(crop_rect, color=(0.97, 0.32, 0.29), width=3)
            
            shape = page.new_shape()
            shape.draw_rect(crop_rect)
            shape.finish(color=(0.97, 0.32, 0.29), fill=(0.97, 0.32, 0.29), fill_opacity=0.15)
            shape.commit()
            
            label_point = fitz.Point(x1 + 5, y1 - 5)
            page.insert_text(label_point, "EXTRACTION AREA", fontsize=14, color=(0.97, 0.32, 0.29), fontname="helv")
        
        mat = fitz.Matrix(zoom_level, zoom_level)
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")
        doc.close()
        return img_data, None
    except Exception as e:
        return None, str(e)

def extract_from_region(pdf_file, page_num, crop_box):
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            page = pdf.pages[page_num - 1]
            width, height = page.width, page.height
            
            if crop_box:
                x1, y1, x2, y2 = crop_box
                crop_area = (width*x1/100, height*y1/100, width*x2/100, height*y2/100)
                page = page.crop(crop_area)
            
            text = page.extract_text()
            return text
    except Exception as e:
        return ""

def process_all_pages(pdf_file, crop_box, detector, progress_callback=None):
    all_items = []
    logs = []
    
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            total_pages = len(pdf.pages)
            logs.append(f"Processing {total_pages} pages...")
            
            for page_num in range(total_pages):
                if progress_callback:
                    progress_callback((page_num + 1) / total_pages, f"Page {page_num + 1}")
                
                page = pdf.pages[page_num]
                width, height = page.width, page.height
                
                if crop_box:
                    x1, y1, x2, y2 = crop_box
                    page = page.crop((width*x1/100, height*y1/100, width*x2/100, height*y2/100))
                
                text = page.extract_text()
                items = detector.parse_table_content(text)
                
                for item in items:
                    item["Page"] = page_num + 1
                
                all_items.extend(items)
                logs.append(f"✓ Page {page_num+1}: {len(items)} items")
                
    except Exception as e:
        logs.append(f"Error: {str(e)}")
    
    return all_items, logs

def create_excel_download(df, yellow_header=True):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ Data"
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_font = Font(bold=True, size=10, color="000000")
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    
    preferred_order = ["Drawing No", "Mark No", "Item No", "Quantity", "Part No", "Description", "Material", "Page"]
    available_cols = [c for c in preferred_order if c in df.columns]
    other_cols = [c for c in df.columns if c not in preferred_order]
    final_cols = available_cols + other_cols
    df_export = df[final_cols]
    
    for col_num, header in enumerate(df_export.columns, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        if yellow_header:
            cell.fill = yellow_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    
    for row_num, row_data in enumerate(df_export.values, 2):
        for col_num, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    
    for idx in range(1, len(df_export.columns) + 1):
        col_letter = chr(64 + idx) if idx <= 26 else chr(64 + (idx-1)//26) + chr(65 + (idx-1)%26)
        max_length = 0
        for row_idx in range(1, len(df_export) + 2):
            cell_value = ws.cell(row=row_idx, column=idx).value
            if cell_value:
                max_length = max(max_length, len(str(cell_value)))
        ws.column_dimensions[col_letter].width = min(max_length + 2, 50)
    
    ws.freeze_panes = "A2"
    wb.save(output)
    output.seek(0)
    return output

def main():
    load_css()
    
    if "extracted_data" not in st.session_state:
        st.session_state.extracted_data = pd.DataFrame()
    if "detector" not in st.session_state:
        st.session_state.detector = SmartTableDetector()
    if "pdf_file" not in st.session_state:
        st.session_state.pdf_file = None
    if "crop_coords" not in st.session_state:
        st.session_state.crop_coords = (5, 15, 95, 60)
    
    with st.sidebar:
        st.markdown("### Settings")
        max_items = st.number_input("Max Items Per Page", 1, 50, 15)
        st.session_state.detector.max_items = max_items
        st.markdown("---")
        st.markdown("**Crop Region**")
        st.markdown("<div class='slider-label'>Left %</div>", unsafe_allow_html=True)
        x1 = st.slider("X1", 0, 100, st.session_state.crop_coords[0], key="x1", label_visibility="collapsed")
        st.markdown("<div class='slider-label'>Top %</div>", unsafe_allow_html=True)
        y1 = st.slider("Y1", 0, 100, st.session_state.crop_coords[1], key="y1", label_visibility="collapsed")
        st.markdown("<div class='slider-label'>Right %</div>", unsafe_allow_html=True)
        x2 = st.slider("X2", 0, 100, st.session_state.crop_coords[2], key="x2", label_visibility="collapsed")
        st.markdown("<div class='slider-label'>Bottom %</div>", unsafe_allow_html=True)
        y2 = st.slider("Y2", 0, 100, st.session_state.crop_coords[3], key="y2", label_visibility="collapsed")
        if x2 <= x1: x2 = min(x1 + 10, 100)
        if y2 <= y1: y2 = min(y1 + 10, 100)
        current_crop = (x1, y1, x2, y2)
        st.session_state.crop_coords = current_crop
        st.markdown("**Crop Size**")
        st.markdown(f'<div class="crop-info"><div class="crop-dimension"><span>Width:</span><span>{x2-x1}%</span></div><div class="crop-dimension"><span>Height:</span><span>{y2-y1}%</span></div></div>', unsafe_allow_html=True)
        st.markdown("---")
        zoom_level = st.select_slider("Zoom", options=[1.0, 1.5, 2.0, 2.5, 3.0], value=2.0, format_func=lambda x: f"{x}x")
        st.markdown("---")
        yellow_header = st.checkbox("Yellow Headers", value=True)
        if st.button("Clear", type="secondary"):
            st.session_state.extracted_data = pd.DataFrame()
            st.session_state.pdf_file = None
            st.rerun()
    
    col_left, col_right = st.columns([1, 1])
    with col_left:
        st.markdown("<div class='section-header'>📤 Upload PDF</div>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Select PDF", type=["pdf"], label_visibility="collapsed")
        if uploaded_file:
            st.success(f"✓ {uploaded_file.name}")
            st.session_state.pdf_file = uploaded_file
    with col_right:
        st.markdown("<div class='section-header'>🚀 Actions</div>", unsafe_allow_html=True)
        if uploaded_file:
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("👁️ Preview", use_container_width=True):
                    st.session_state.show_preview = True
            with col_btn2:
                if st.button("🔍 Extract All", use_container_width=True, type="primary"):
                    st.session_state.do_extract = True
    
    if st.session_state.pdf_file and st.session_state.get("show_preview", False):
        st.markdown("<div class='section-header'>🎯 Preview (Red Box = Extraction Area)</div>", unsafe_allow_html=True)
        with st.spinner("Rendering..."):
            img_data, error = render_pdf_page_with_crop(st.session_state.pdf_file, 1, current_crop, zoom_level)
            if error:
                st.error(error)
            else:
                st.image(img_data, use_column_width=True)
                sample_text = extract_from_region(st.session_state.pdf_file, 1, current_crop)
                if sample_text:
                    with st.expander("Sample Text"):
                        st.markdown(f"<div class='extracted-text'>{sample_text[:1000]}</div>", unsafe_allow_html=True)
    
    if st.session_state.pdf_file and st.session_state.get("do_extract", False):
        st.markdown("<div class='section-header'>⚙️ Extracting All Pages...</div>", unsafe_allow_html=True)
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        def update_progress(progress, message):
            progress_bar.progress(progress, message)
        
        items, logs = process_all_pages(st.session_state.pdf_file, current_crop, st.session_state.detector, update_progress)
        
        progress_bar.empty()
        status_text.empty()
        
        if items:
            df = pd.DataFrame(items)
            st.session_state.extracted_data = df
            st.success(f"✓ Extracted {len(df)} total items from all pages")
            with st.expander("Processing Logs"):
                for log in logs:
                    st.text(log)
        else:
            st.warning("No items found")
        st.session_state.do_extract = False
    
    if not st.session_state.extracted_data.empty:
        df = st.session_state.extracted_data
        st.markdown("<div class='section-header'>📊 Results (Items 1-15 per page)</div>", unsafe_allow_html=True)
        
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"<div class='stat-card'><div class='stat-value'>{len(df)}</div><div class='stat-label'>Total Items</div></div>", unsafe_allow_html=True)
        with c2:
            pages = df["Page"].nunique() if "Page" in df.columns else 1
            st.markdown(f"<div class='stat-card'><div class='stat-value'>{pages}</div><div class='stat-label'>Pages</div></div>", unsafe_html=True)
        with c3:
            st.markdown(f"<div class='stat-card'><div class='stat-value'>{max_items}</div><div class='stat-label'>Items/Page</div></div>", unsafe_allow_html=True)
        with c4:
            materials = df["Material"].nunique() if "Material" in df.columns else 0
            st.markdown(f"<div class='stat-card'><div class='stat-value'>{materials}</div><div class='stat-label'>Materials</div></div>", unsafe_html=True)
        
        st.markdown("### ✏️ Edit Data")
        edited_df = st.data_editor(
            df, num_rows="dynamic", use_container_width=True, hide_index=True,
            column_config={
                "Drawing No": st.column_config.TextColumn("Drawing No", width="small"),
                "Mark No": st.column_config.TextColumn("Mark No", width="small"),
                "Item No": st.column_config.NumberColumn("Item", width="small"),
                "Quantity": st.column_config.NumberColumn("Qty", width="small"),
                "Part No": st.column_config.TextColumn("Part No", width="medium"),
                "Description": st.column_config.TextColumn("Description", width="large"),
                "Material": st.column_config.TextColumn("Material", width="medium"),
                "Page": st.column_config.NumberColumn("Page", width="small"),
            }
        )
        
        st.markdown("<div class='section-header'>📦 Export</div>", unsafe_html=True)
        col_excel, col_csv = st.columns(2)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        with col_excel:
            try:
                excel_data = create_excel_download(edited_df, yellow_header)
                st.download_button(label="📥 Download Excel", data=excel_data, file_name=f"BOQ_{timestamp}.xlsx", 
                                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except Exception as e:
                st.error(f"Export error: {e}")
        with col_csv:
            csv_data = edited_df.to_csv(index=False).encode("utf-8")
            st.download_button(label="📄 Download CSV", data=csv_data, file_name=f"BOQ_{timestamp}.csv", 
                             mime="text/csv", use_container_width=True)
    
    with st.expander("❓ Help"):
        st.markdown(f"""
        **FINISHED PRODUCT v16.0**
        
        ✓ **Multi-Page**: Extracts items 1-{max_items} from EACH page
        ✓ **Red Crop Preview**: Shows extraction area with red box
        ✓ **Smart Detection**: No grid lines needed - detects by content
        ✓ **Columns**: Drawing No | Mark No | Item | Qty | Part No | Description | Material | Page
        
        **How to use**:
        1. Upload PDF
        2. Adjust crop region (red box shows area)
        3. Click "Extract All" 
        4. Gets items 1-{max_items} from every page
        """)

if __name__ == "__main__":
    main()
