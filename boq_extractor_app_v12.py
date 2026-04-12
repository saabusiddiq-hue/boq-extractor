"""
BOQ Extractor Pro v15.1 - Clean Layout Edition
No Header Banner, Direct Functional Interface
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
        .section-header h3 { margin: 0; font-size: 1rem; color: #58a6ff !important; }
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
        .stSuccess { background-color: rgba(35, 134, 54, 0.1) !important; border: 1px solid #238636 !important; color: #3fb950 !important; padding: 0.5rem !important; border-radius: 6px !important; font-size: 0.9rem !important; }
        .stInfo { background-color: rgba(56, 139, 253, 0.1) !important; border: 1px solid #388bfd !important; color: #58a6ff !important; padding: 0.75rem !important; border-radius: 6px !important; font-size: 0.85rem !important; }
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
        skip_keywords = ["ITEM", "NO.", "REQ'D", "FIG", "DESCRIPTION", "MATERIAL", "NO", "REQUIRED", "FIGURE", "PART", "DRAWING", "MARK"]
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
                
                if remaining and remaining[0].isdigit() and 1 <= int(remaining[0]) <= 20:
                    result["Quantity"] = int(remaining[0])
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
            r"^[A-Z]\d+[-_][A-Z]?\d+",
            r"^F\d+[-_]M\d+",
            r"^V\d+[-_]\d+[-_][A-Z]+\d*",
            r"^PC\d+[-_]\d+",
            r"^F\d+[-_]TS\d+",
            r"^F\d+[-_]M\d+\s*L?=\d*",
            r"^PTFE",
            r"^Variable",
            r"^M\d+[:]?\d*",
            r"^3x\d+",
            r"^15x\d+"
        ]
        return any(re.match(p, text, re.IGNORECASE) for p in patterns)
    
    def _is_material(self, text):
        patterns = [r"A36\b", r"A105\b", r"A193", r"A194", r"MSS[-]?SP", r"SS316", r"SS304", r"A516", r"A240", r"Gr[.]?\s*\d+", r"CI[.]?\s*\d+", r"PTFE", r"Top\s*Plate", r"Base\s*Plate", r"Stud\s*Bolt", r"Nut"]
        return any(re.search(p, text, re.IGNORECASE) for p in patterns)
    
    def _extract_material_from_end(self, text):
        mat_match = re.search(r"(A36|A105|A193\s*GR?[.]?B7|A194\s*GR?[.]?2H|Per\s*MSS[-]?SP\d+|SS316|SS304|A516|Gr[.]?\s*\d+[.]?\d*|CI[.]?\s*\d+|PTFE)(.*?)$", text, re.IGNORECASE)
        if mat_match:
            return mat_match.group(0).strip()
        return ""
    
    def detect_columns_from_pdf(self, page):
        chars = page.chars
        if not chars:
            return None
        x_positions = {}
        for char in chars:
            x_rounded = round(char["x0"] / 10) * 10
            if x_rounded not in x_positions:
                x_positions[x_rounded] = []
            x_positions[x_rounded].append(char)
        sorted_x = sorted(x_positions.keys())
        columns = []
        current_col = [sorted_x[0]] if sorted_x else []
        for i in range(1, len(sorted_x)):
            if sorted_x[i] - sorted_x[i-1] < 50:
                current_col.append(sorted_x[i])
            else:
                avg_x = sum(current_col) / len(current_col)
                columns.append({"x": avg_x, "chars": sum(len(x_positions[x]) for x in current_col)})
                current_col = [sorted_x[i]]
        if current_col:
            avg_x = sum(current_col) / len(current_col)
            columns.append({"x": avg_x, "chars": sum(len(x_positions[x]) for x in current_col)})
        return [c for c in columns if c["chars"] > 20][:7]

def add_grid_overlay(img_data, column_boundaries):
    try:
        img = Image.open(io.BytesIO(img_data))
        draw = ImageDraw.Draw(img)
        width, height = img.size
        for col in column_boundaries:
            x = int(col["x"])
            if 0 <= x < width:
                draw.line([(x, 0), (x, height)], fill=(31, 111, 235), width=2)
        output = io.BytesIO()
        img.save(output, format="PNG")
        return output.getvalue()
    except:
        return img_data

def render_pdf_with_grid(pdf_file, page_num, detector, zoom_level=2.0):
    try:
        pdf_file.seek(0)
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        if page_num > len(doc):
            doc.close()
            return None, None, "Invalid page"
        page = doc[page_num - 1]
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            plum_page = pdf.pages[page_num - 1]
            columns = detector.detect_columns_from_pdf(plum_page)
        mat = fitz.Matrix(zoom_level, zoom_level)
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")
        if columns:
            img_data = add_grid_overlay(img_data, columns)
        doc.close()
        return img_data, columns, None
    except Exception as e:
        return None, None, str(e)

def extract_with_smart_detection(pdf_file, page_num, crop_box, detector):
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            page = pdf.pages[page_num - 1]
            width, height = page.width, page.height
            if crop_box:
                x1, y1, x2, y2 = crop_box
                page = page.crop((width*x1/100, height*y1/100, width*x2/100, height*y2/100))
            text = page.extract_text()
            items = detector.parse_table_content(text)
            for item in items:
                item["Page"] = page_num
            return items, text, None
    except Exception as e:
        return [], "", str(e)

def create_excel_download(df, yellow_header=True):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ Data"
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_font = Font(bold=True, size=10, color="000000")
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    # Match Excel template: Drawing No, Mark No, Item No, Quantity, Part No, Description, Material
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
    
    # SIDEBAR
    with st.sidebar:
        st.markdown("### Settings")
        max_items = st.number_input("Max Items", 1, 50, 15)
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
        st.markdown(f"""
        <div class="crop-info">
            <div class="crop-dimension"><span>Width:</span><span>{x2-x1}%</span></div>
            <div class="crop-dimension"><span>Height:</span><span>{y2-y1}%</span></div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("---")
        zoom_level = st.select_slider("Zoom", options=[1.0, 1.5, 2.0, 2.5, 3.0], value=2.0, format_func=lambda x: f"{x}x")
        st.markdown("---")
        yellow_header = st.checkbox("Yellow Headers", value=True)
        if st.button("Clear", type="secondary"):
            st.session_state.extracted_data = pd.DataFrame()
            st.session_state.pdf_file = None
            st.rerun()
    
    # MAIN CONTENT - CLEAN NO BANNER
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
                if st.button("👁️ Preview Grid", use_container_width=True):
                    st.session_state.show_preview = True
            with col_btn2:
                if st.button("🔍 Extract", use_container_width=True, type="primary"):
                    st.session_state.do_extract = True
    
    # PREVIEW
    if st.session_state.pdf_file and st.session_state.get("show_preview", False):
        st.markdown("<div class='section-header'>🎯 Preview with Grid</div>", unsafe_allow_html=True)
        with st.spinner("Detecting..."):
            img_data, columns, error = render_pdf_with_grid(st.session_state.pdf_file, 1, st.session_state.detector, zoom_level)
            if error:
                st.error(error)
            else:
                st.image(img_data, use_column_width=True, caption="Blue vertical lines = detected columns")
                if columns:
                    st.info(f"Detected {len(columns)} columns")
    
    # EXTRACTION
    if st.session_state.pdf_file and st.session_state.get("do_extract", False):
        st.markdown("<div class='section-header'>⚙️ Extraction</div>", unsafe_allow_html=True)
        with st.spinner("Extracting..."):
            items, raw_text, error = extract_with_smart_detection(st.session_state.pdf_file, 1, current_crop, st.session_state.detector)
            if error:
                st.error(error)
            elif items:
                df = pd.DataFrame(items)
                st.session_state.extracted_data = df
                st.success(f"✓ Extracted {len(df)} items (1-{max_items})")
                with st.expander("View raw text"):
                    st.markdown(f"<div class='extracted-text'>{raw_text[:1500]}</div>", unsafe_allow_html=True)
            else:
                st.warning("No items found")
        st.session_state.do_extract = False
    
    # RESULTS
    if not st.session_state.extracted_data.empty:
        df = st.session_state.extracted_data
        st.markdown("<div class='section-header'>📊 Results</div>", unsafe_allow_html=True)
        
        # Stats
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"<div class='stat-card'><div class='stat-value'>{len(df)}</div><div class='stat-label'>Items</div></div>", unsafe_allow_html=True)
        with c2:
            total_qty = int(df["Quantity"].sum()) if "Quantity" in df.columns else 0
            st.markdown(f"<div class='stat-card'><div class='stat-value'>{total_qty}</div><div class='stat-label'>Quantity</div></div>", unsafe_allow_html=True)
        with c3:
            materials = df["Material"].nunique() if "Material" in df.columns else 0
            st.markdown(f"<div class='stat-card'><div class='stat-value'>{materials}</div><div class='stat-label'>Materials</div></div>", unsafe_allow_html=True)
        with c4:
            st.markdown(f"<div class='stat-card'><div class='stat-value'>{len(df.columns)}</div><div class='stat-label'>Columns</div></div>", unsafe_allow_html=True)
        
        # Data editor with correct column order
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
            }
        )
        
        # Export
        st.markdown("<div class='section-header'>📦 Export</div>", unsafe_allow_html=True)
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
        st.markdown("""
        **Auto Detection**: Upload PDF → Preview Grid → Extract. App detects columns by content position.
        
        **Excel Template Columns**:
        1. Drawing No | 2. Mark No | 3. Item No | 4. Quantity | 5. Part No | 6. Description | 7. Material
        
        **Limits**: Extracts only items 1-15 (stops before Spring details section)
        """)

if __name__ == "__main__":
    main()
