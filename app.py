"""
BOQ Extractor Pro v21.0 - TRIPLE VALIDATION EDITION
Structure Learning + Sample Image + Sample PDF + 3-Step Review
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import shutil
from datetime import datetime
from typing import Optional, Tuple, List, Dict, Any, Callable

# Setup OCR
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

st.set_page_config(page_title="BOQ Extractor Pro", page_icon="📋", layout="wide")

# CSS
st.markdown("""
<style>
    .stApp { background-color: #0d1117; }
    .main-header { 
        background: linear-gradient(90deg, #1f6feb 0%, #238636 100%);
        padding: 1rem;
        border-radius: 8px;
        color: white;
        text-align: center;
        margin-bottom: 1rem;
    }
    .step-box {
        background: #161b22;
        border: 2px solid #30363d;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
    }
    .step-active {
        border-color: #1f6feb !important;
        box-shadow: 0 0 10px rgba(31, 111, 235, 0.3);
    }
    .step-complete {
        border-color: #238636 !important;
    }
    .validation-card {
        background: #0d1117;
        border-left: 4px solid #f85149;
        padding: 1rem;
        margin: 0.5rem 0;
        border-radius: 0 6px 6px 0;
    }
    .validation-good {
        border-left-color: #238636 !important;
    }
    .material-highlight {
        background: rgba(31, 111, 235, 0.2);
        color: #58a6ff;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: 600;
    }
    .section-header { 
        background: #21262d; 
        padding: 0.75rem 1rem; 
        border-radius: 8px; 
        margin: 1rem 0 0.5rem 0; 
        border-left: 3px solid #1f6feb; 
        color: #58a6ff !important; 
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# Standard materials
STANDARD_MATERIALS = [
    "Graphite bronze", "Graphite", "Bronze", "PTFE", "Lead",
    "SS304", "SS316", "SS316L", "SS304L", "Stainless Steel",
    "A36", "A105", "A193 GR.B7", "A194 GR.2H", "A240", "A516",
    "Carbon Steel", "Cast Iron", "CI", "MSS-SP", "Galvanized",
    "Rubber", "Neoprene", "EPDM", "Viton"
]

class StructureLearner:
    """Learn BOQ structure from sample image or PDF"""
    def __init__(self):
        self.learned_patterns = {
            "header_keywords": [],
            "column_positions": {},
            "part_no_pattern": r"^[A-Z]\d+[-_][A-Z]?\d+",
            "material_keywords": STANDARD_MATERIALS.copy()
        }
        self.is_learned = False
        self.sample_text = ""

    def learn_from_text(self, text: str, source: str = "unknown"):
        """Learn structure from extracted text"""
        self.sample_text = text
        lines = text.strip().split("\n")

        # Find header row
        for i, line in enumerate(lines[:10]):
            line_upper = line.upper()
            if any(kw in line_upper for kw in ["ITEM", "NO.", "DESCRIPTION", "QTY", "MATERIAL", "PART"]):
                headers = [h.strip() for h in re.split(r"\s{2,}|\t", line) if h.strip()]
                self.learned_patterns["header_keywords"] = headers

                # Detect column positions
                for j, h in enumerate(headers):
                    h_upper = h.upper()
                    if any(x in h_upper for x in ["ITEM", "NO."]):
                        self.learned_patterns["column_positions"]["item"] = j
                    elif any(x in h_upper for x in ["PART", "FIG", "MARK"]):
                        self.learned_patterns["column_positions"]["part"] = j
                    elif any(x in h_upper for x in ["DESC"]):
                        self.learned_patterns["column_positions"]["desc"] = j
                    elif any(x in h_upper for x in ["MAT", "MATERIAL"]):
                        self.learned_patterns["column_positions"]["material"] = j
                break

        # Learn part number patterns from first few data rows
        for line in lines[5:15]:
            if line.strip() and line[0].isdigit():
                words = line.split()
                for word in words[1:4]:
                    if re.match(r"^[A-Z]\d+[-_][A-Z]?\d+|^\d+x\d+|^F\d+[-_]M\d+", word, re.IGNORECASE):
                        self.learned_patterns["part_no_pattern"] = word
                        break
                break

        # Extract unique materials from sample
        found_materials = []
        for mat in STANDARD_MATERIALS:
            if mat.upper() in text.upper():
                found_materials.append(mat)
        if found_materials:
            self.learned_patterns["material_keywords"] = found_materials

        self.is_learned = True
        return self.learned_patterns

    def learn_from_image(self, image_bytes: bytes) -> Dict:
        """OCR sample image and learn"""
        if not OCR_AVAILABLE:
            return {"error": "OCR not available"}

        try:
            image = Image.open(io.BytesIO(image_bytes))
            # Enhance for OCR
            enhancer = ImageEnhance.Contrast(image)
            image = enhancer.enhance(2.0)
            image = image.filter(ImageFilter.SHARPEN)

            text = pytesseract.image_to_string(image, config="--psm 6")
            return self.learn_from_text(text, "image")
        except Exception as e:
            return {"error": str(e)}

    def learn_from_pdf(self, pdf_bytes: bytes, page: int = 1) -> Dict:
        """Extract from sample PDF and learn"""
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                if page > len(pdf.pages):
                    page = 1
                text = pdf.pages[page - 1].extract_text() or ""
                return self.learn_from_text(text, f"pdf_page_{page}")
        except Exception as e:
            return {"error": str(e)}

class TripleValidator:
    """3-step validation: Extract -> Review -> Finalize"""
    def __init__(self, learner: StructureLearner):
        self.learner = learner
        self.validations = []

    def validate_extraction(self, items: List[Dict]) -> List[Dict]:
        """Step 1: Initial validation after extraction"""
        validations = []
        for item in items:
            issues = []

            # Check 1: Material merged with description
            desc = item.get("Description", "")
            mat = item.get("Material", "")

            if not mat and desc:
                # Try to extract material from description
                for std_mat in self.learner.learned_patterns.get("material_keywords", STANDARD_MATERIALS):
                    pattern = rf"^({re.escape(std_mat)})\s*[-,:/]?\s*(.*)"
                    match = re.match(pattern, desc, re.IGNORECASE)
                    if match:
                        issues.append({
                            "type": "material_merged",
                            "suggested_mat": match.group(1),
                            "suggested_desc": match.group(2),
                            "confidence": "high"
                        })
                        break

            # Check 2: Material at end of description
            if not mat and desc:
                for std_mat in self.learner.learned_patterns.get("material_keywords", STANDARD_MATERIALS):
                    pattern = rf"(.*?)\s*[-,:/]?\s*({re.escape(std_mat)})$"
                    match = re.match(pattern, desc, re.IGNORECASE)
                    if match:
                        issues.append({
                            "type": "material_at_end",
                            "suggested_mat": match.group(2),
                            "suggested_desc": match.group(1),
                            "confidence": "medium"
                        })
                        break

            # Check 3: Empty critical fields
            if not item.get("Part No"):
                issues.append({"type": "missing_part_no", "confidence": "low"})

            if not item.get("Description"):
                issues.append({"type": "missing_description", "confidence": "high"})

            validations.append({
                "item": item,
                "issues": issues,
                "status": "needs_review" if issues else "ok"
            })

        self.validations = validations
        return validations

    def apply_auto_fixes(self, validations: List[Dict]) -> List[Dict]:
        """Apply high-confidence fixes automatically"""
        fixed_items = []
        for val in validations:
            item = val["item"].copy()

            for issue in val["issues"]:
                if issue["confidence"] == "high" and issue["type"] in ["material_merged", "material_at_end"]:
                    item["Material"] = issue["suggested_mat"]
                    item["Description"] = issue["suggested_desc"]
                    item["_auto_fixed"] = True

            fixed_items.append(item)

        return fixed_items

def extract_material_from_text(text: str, materials: List[str]) -> Tuple[str, str]:
    """Extract material from text"""
    if not text:
        return "", ""

    # Check start
    for mat in sorted(materials, key=len, reverse=True):
        pattern = rf"^({re.escape(mat)})\s*[-,:/]?\s*(.*)"
        match = re.match(pattern, text, re.IGNORECASE)
        if match:
            return match.group(2).strip(), match.group(1)

    # Check end
    for mat in sorted(materials, key=len, reverse=True):
        pattern = rf"(.*?)\s*[-,:/]?\s*({re.escape(mat)})$"
        match = re.match(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip(), match.group(2)

    return text, ""

def render_pdf_page(pdf_bytes: bytes, page_num: int, crop: Optional[Tuple], zoom: float = 2.0):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if page_num > len(doc):
            return None, "Invalid page"

        page = doc[page_num - 1]

        if crop:
            x1, y1, x2, y2 = crop
            rect = page.rect
            x1, y1 = rect.width * x1/100, rect.height * y1/100
            x2, y2 = rect.width * x2/100, rect.height * y2/100

            shape = page.new_shape()
            shape.draw_rect(fitz.Rect(x1, y1, x2, y2))
            shape.finish(color=(0.97, 0.32, 0.29), fill=(0.97, 0.32, 0.29), fill_opacity=0.15, width=3)
            shape.commit()
            page.insert_text(fitz.Point(x1 + 5, max(y1 - 5, 10)), "CROP AREA", 
                           fontsize=14, color=(0.97, 0.32, 0.29))

        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        return pix.tobytes("png"), None
    except Exception as e:
        return None, str(e)
    finally:
        if doc:
            doc.close()

def create_excel(df: pd.DataFrame, yellow_header: bool = True):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ"

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_font = Font(bold=True, size=10)
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                   top=Side(style="thin"), bottom=Side(style="thin"))

    # Reorder columns
    preferred = ["Drawing No", "Mark No", "Item No", "Quantity", "Part No", "Description", "Material", "Page"]
    cols = [c for c in preferred if c in df.columns] + [c for c in df.columns if c not in preferred and not c.startswith("_")]
    df = df[cols]

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
if "learner" not in st.session_state:
    st.session_state.learner = StructureLearner()
if "step" not in st.session_state:
    st.session_state.step = 1  # 1=Upload/Learn, 2=Extract, 3=Review, 4=Finalize
if "data" not in st.session_state:
    st.session_state.data = None
if "validations" not in st.session_state:
    st.session_state.validations = None

st.markdown("<div class='main-header'><h1>📋 BOQ Extractor Pro v21.0</h1><p>Triple Validation: Sample → Extract → Review → Export</p></div>", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.header("⚙️ Settings")
    max_items = st.number_input("Max Items/Page", 1, 50, 15)

    st.divider()
    st.write("📐 Crop Region")
    x1 = st.slider("Left %", 0, 100, 5)
    y1 = st.slider("Top %", 0, 100, 15)
    x2 = st.slider("Right %", 0, 100, 95)
    y2 = st.slider("Bottom %", 0, 100, 60)
    if x2 <= x1: x2 = min(x1 + 10, 100)
    if y2 <= y1: y2 = min(y1 + 10, 100)
    crop = (x1, y1, x2, y2)
    st.write(f"Size: {x2-x1}% × {y2-y1}%")

    zoom = st.select_slider("Zoom", [1.0, 1.5, 2.0, 2.5, 3.0], 2.0)
    yellow_header = st.checkbox("Yellow Headers", True)

    if st.button("🗑️ Reset All"):
        for k in ["learner", "step", "data", "validations", "sample_img", "sample_pdf", "target_pdf"]:
            st.session_state[k] = None if k != "learner" else StructureLearner()
        st.session_state.step = 1
        st.rerun()

# Progress indicator
col_steps = st.columns(4)
steps = ["1. Learn", "2. Extract", "3. Review", "4. Export"]
for i, (col, step_name) in enumerate(zip(col_steps, steps)):
    css_class = "step-box"
    if st.session_state.step == i + 1:
        css_class += " step-active"
    elif st.session_state.step > i + 1:
        css_class += " step-complete"
    with col:
        st.markdown(f"<div class='{css_class}'><center><b>{step_name}</b></center></div>", unsafe_allow_html=True)

# STEP 1: UPLOAD SAMPLES
if st.session_state.step == 1:
    st.markdown("<div class='section-header'>📚 Step 1: Upload Reference Samples</div>", unsafe_allow_html=True)
    st.info("Upload a sample BOQ (image or PDF) to teach the system your format. Then upload the target PDF to extract.")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Sample Reference (Image)")
        sample_img = st.file_uploader("Upload JPG/PNG sample", type=["jpg", "jpeg", "png"], key="sample_img")

        if sample_img:
            st.image(sample_img, caption="Sample Image", use_column_width=True)
            if OCR_AVAILABLE and st.button("🔍 Learn from Image", use_container_width=True):
                with st.spinner("Analyzing..."):
                    result = st.session_state.learner.learn_from_image(sample_img.read())
                    if "error" not in result:
                        st.success(f"✅ Learned {len(result.get('material_keywords', []))} materials")
                        with st.expander("View Learned Structure"):
                            st.json(result)
                    else:
                        st.error(result["error"])

    with col2:
        st.subheader("Sample Reference (PDF)")
        sample_pdf = st.file_uploader("Upload PDF sample", type=["pdf"], key="sample_pdf")

        if sample_pdf:
            pdf_bytes = sample_pdf.read()
            st.success(f"✓ {sample_pdf.name}")

            # Show page selector
            try:
                with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                    total_pages = len(pdf.pages)
                sample_page = st.number_input("Sample page", 1, total_pages, 1)
            except:
                sample_page = 1

            # Preview
            img, _ = render_pdf_page(pdf_bytes, sample_page, crop, zoom)
            if img:
                st.image(img, caption="Sample Preview", use_column_width=True)

            if st.button("🔍 Learn from PDF", use_container_width=True):
                with st.spinner("Analyzing..."):
                    result = st.session_state.learner.learn_from_pdf(pdf_bytes, sample_page)
                    if "error" not in result:
                        st.success(f"✅ Learned structure from PDF")
                        with st.expander("View Learned Structure"):
                            st.json(result)
                    else:
                        st.error(result["error"])

    # Target PDF upload
    st.divider()
    st.subheader("🎯 Target PDF (to Extract)")
    target_pdf = st.file_uploader("Upload target PDF", type=["pdf"], key="target_pdf")

    if target_pdf:
        pdf_bytes = target_pdf.read()
        st.session_state.target_pdf = pdf_bytes
        st.session_state.target_name = target_pdf.name

        img, _ = render_pdf_page(pdf_bytes, 1, crop, zoom)
        if img:
            st.image(img, caption="Target Preview", use_column_width=True)

        if st.button("➡️ Proceed to Extraction", type="primary", use_container_width=True):
            st.session_state.step = 2
            st.rerun()

# STEP 2: EXTRACTION
if st.session_state.step == 2:
    st.markdown("<div class='section-header'>⚡ Step 2: Extract BOQ Data</div>", unsafe_allow_html=True)

    if st.session_state.get("target_pdf"):
        if st.button("🚀 Extract All Pages", type="primary", use_container_width=True):
            with st.spinner("Processing..."):
                # Simple extraction logic
                items = []
                pdf_bytes = st.session_state.target_pdf

                with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                    for page_num in range(1, len(pdf.pages) + 1):
                        page = pdf.pages[page_num - 1]
                        if crop != (0, 0, 100, 100):
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

                                qty = 1
                                remaining = parts[1:]
                                if remaining[0].isdigit():
                                    qty = int(remaining[0])
                                    remaining = remaining[1:]

                                # Find part number
                                part_no = ""
                                desc_start = 0
                                for i, part in enumerate(remaining[:4]):
                                    if re.match(r"^[A-Z0-9][-A-Z0-9x]+", part, re.IGNORECASE):
                                        part_no = part
                                        desc_start = i + 1
                                        break

                                raw_desc = " ".join(remaining[desc_start:]) if desc_start < len(remaining) else ""

                                # Extract material using learned patterns
                                materials = st.session_state.learner.learned_patterns.get("material_keywords", STANDARD_MATERIALS)
                                clean_desc, material = extract_material_from_text(raw_desc, materials)

                                items.append({
                                    "Item No": item_no,
                                    "Quantity": qty,
                                    "Part No": part_no,
                                    "Description": clean_desc,
                                    "Material": material,
                                    "_raw": raw_desc,
                                    "Page": page_num
                                })
                            except:
                                continue

                if items:
                    st.session_state.data = items
                    st.success(f"✅ Extracted {len(items)} items")

                    # Run validation
                    validator = TripleValidator(st.session_state.learner)
                    validations = validator.validate_extraction(items)
                    st.session_state.validations = validations

                    # Count issues
                    issues_count = sum(1 for v in validations if v["status"] == "needs_review")
                    if issues_count > 0:
                        st.warning(f"⚠️ Found {issues_count} items needing review")

                    st.session_state.step = 3
                    st.rerun()
                else:
                    st.error("No items found")
    else:
        st.error("No target PDF uploaded. Go back to Step 1.")
        if st.button("← Back to Step 1"):
            st.session_state.step = 1
            st.rerun()

# STEP 3: REVIEW (3-TIME CHECKING)
if st.session_state.step == 3:
    st.markdown("<div class='section-header'>🔍 Step 3: Review & Correct (Check 1 of 3)</div>", unsafe_allow_html=True)

    if st.session_state.validations:
        # Summary
        total = len(st.session_state.validations)
        needs_review = sum(1 for v in st.session_state.validations if v["status"] == "needs_review")
        auto_fixed = sum(1 for v in st.session_state.validations for i in v["issues"] if i.get("confidence") == "high")

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Items", total)
        c2.metric("Needs Review", needs_review, delta=None if needs_review == 0 else f"{needs_review} items")
        c3.metric("Auto-Fixed", auto_fixed)

        st.divider()

        # Apply high-confidence fixes automatically
        if st.button("🤖 Apply High-Confidence Fixes", use_container_width=True):
            validator = TripleValidator(st.session_state.learner)
            fixed = validator.apply_auto_fixes(st.session_state.validations)
            st.session_state.data = fixed
            st.success("✅ Applied automatic fixes")
            st.rerun()

        # Manual review table
        st.subheader("Manual Review Required")

        edited_data = []
        for idx, validation in enumerate(st.session_state.validations):
            item = validation["item"]
            issues = validation["issues"]

            # Show only items with issues or all if user wants
            show_all = st.checkbox("Show all items (not just issues)", value=False, key="show_all")

            if issues or show_all:
                with st.container():
                    css_class = "validation-card"
                    if not issues:
                        css_class += " validation-good"

                    st.markdown(f"<div class='{css_class}'>", unsafe_allow_html=True)

                    cols = st.columns([0.5, 0.8, 1.2, 2, 1.5, 0.5])

                    with cols[0]:
                        st.write(f"**#{item['Item No']}"**)
                        if issues:
                            st.caption(f"⚠️ {len(issues)} issue(s)")

                    with cols[1]:
                        st.write(item.get("Part No", "-"))

                    with cols[2]:
                        new_qty = st.number_input("Qty", value=int(item.get("Quantity", 1)), key=f"qty_{idx}", min_value=1)

                    with cols[3]:
                        current_desc = item.get("Description", "")
                        # Show hint if material was extracted
                        if item.get("_raw") and not item.get("Material"):
                            materials = st.session_state.learner.learned_patterns.get("material_keywords", STANDARD_MATERIALS)
                            _, suggested_mat = extract_material_from_text(item["_raw"], materials)
                            if suggested_mat:
                                st.caption(f"💡 Material: {suggested_mat}")

                        new_desc = st.text_area("Description", value=current_desc, key=f"desc_{idx}", height=60, label_visibility="collapsed")

                    with cols[4]:
                        current_mat = item.get("Material", "")
                        new_mat = st.text_input("Material", value=current_mat, key=f"mat_{idx}", label_visibility="collapsed")

                        # Show raw text hint
                        if item.get("_raw") and not current_mat:
                            st.caption(f"Raw: {item['_raw'][:25]}...")

                    with cols[5]:
                        if st.button("✓", key=f"ok_{idx}"):
                            pass

                    st.markdown("</div>", unsafe_allow_html=True)

                    # Build edited item
                    edited_item = item.copy()
                    edited_item["Quantity"] = new_qty
                    edited_item["Description"] = new_desc
                    edited_item["Material"] = new_mat
                    edited_data.append(edited_item)

        # Navigation
        st.divider()
        col1, col2, col3 = st.columns(3)

        with col1:
            if st.button("← Back to Extraction", use_container_width=True):
                st.session_state.step = 2
                st.rerun()

        with col2:
            if st.button("💾 Save Progress", use_container_width=True):
                st.session_state.data = edited_data
                st.success("Progress saved!")

        with col3:
            if st.button("➡️ Final Check & Export →", type="primary", use_container_width=True):
                st.session_state.data = edited_data
                st.session_state.step = 4
                st.rerun()

# STEP 4: FINAL CHECK & EXPORT
if st.session_state.step == 4:
    st.markdown("<div class='section-header'>✅ Step 4: Final Check & Export</div>", unsafe_allow_html=True)

    if st.session_state.data:
        # Convert to DataFrame
        df = pd.DataFrame(st.session_state.data)

        # Remove internal columns for display
        display_df = df.drop(columns=[c for c in df.columns if c.startswith("_")], errors="ignore")

        # Final stats
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Items", len(display_df))
        c2.metric("Pages", display_df["Page"].nunique() if "Page" in display_df.columns else 1)
        c3.metric("Materials", display_df["Material"].nunique() if "Material" in display_df.columns else 0)
        c4.metric("With Part No", (display_df["Part No"] != "").sum())

        # Final editable table
        st.subheader("Final BOQ Data (Last Check)")
        final_df = st.data_editor(display_df, num_rows="dynamic", use_container_width=True, hide_index=True, key="final_edit")

        # Validation summary
        empty_materials = (final_df["Material"] == "").sum() if "Material" in final_df.columns else 0
        empty_descriptions = (final_df["Description"] == "").sum() if "Description" in final_df.columns else 0

        if empty_materials > 0 or empty_descriptions > 0:
            st.warning(f"⚠️ {empty_materials} empty materials, {empty_descriptions} empty descriptions")
        else:
            st.success("✅ All fields populated!")

        # Export
        st.divider()
        st.subheader("📦 Export")

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = st.session_state.get("target_name", "BOQ").replace(".pdf", "")

        col1, col2, col3 = st.columns(3)

        with col1:
            excel_data = create_excel(final_df, yellow_header)
            st.download_button("📥 Download Excel", excel_data, f"{base}_{ts}.xlsx",
                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             use_container_width=True)

        with col2:
            csv_data = final_df.to_csv(index=False).encode("utf-8")
            st.download_button("📄 Download CSV", csv_data, f"{base}_{ts}.csv",
                             "text/csv", use_container_width=True)

        with col3:
            if st.button("🔄 Start New Extraction", use_container_width=True):
                st.session_state.step = 1
                st.session_state.data = None
                st.session_state.validations = None
                st.rerun()
    else:
        st.error("No data to export")
        if st.button("← Go Back"):
            st.session_state.step = 3
            st.rerun()
