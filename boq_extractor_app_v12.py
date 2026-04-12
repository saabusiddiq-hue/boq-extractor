# ==================== ADD THESE IMPORTS AT THE TOP ====================
import base64
from io import BytesIO

# ==================== ADD THIS SAMPLE DATA FUNCTION ====================
def get_sample_boq_pdf() -> bytes:
    """Generate a realistic sample BOQ PDF in memory for demonstration."""
    # This creates a simple but realistic PDF that works well with your SmartTableDetector
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)  # A4 size

    # Title
    page.insert_text((50, 60), "BILL OF QUANTITIES (BOQ) - SAMPLE", fontsize=16, fontname="helv", bold=True)
    page.insert_text((50, 85), "Project: Piping Spool Fabrication - Phase 2", fontsize=11, fontname="helv")

    # Table Header
    header_y = 130
    page.insert_text((50, header_y), "ITEM NO   QTY   PART NO          DESCRIPTION                          MATERIAL", 
                     fontsize=10, fontname="helv", bold=True)

    # Sample Data (carefully crafted to match your parser)
    sample_data = [
        "1     4    F123-M45     90° Elbow 6\" SCH40                    A234 WPB",
        "2     12   V45-2        Gate Valve 4\" 150# RF                 A105",
        "3     8    PC-789       Pipe 6\" SCH80 seamless                A106 Gr.B",
        "4     6    15x300       Stud Bolt with 2 Nuts                A193 B7 / A194 2H",
        "5     25   F67-TS12     Tee 4\" x 4\" x 2\"                     A234 WPB",
        "6     3    M12:45       Flange WN 6\" 300# RF                 A105",
        "7     18   A36-PLT      Plate 10mm x 2000 x 1000             A36",
    ]

    y = 155
    for line in sample_data:
        page.insert_text((50, y), line, fontsize=10, fontname="helv")
        y += 22

    # Footer note
    page.insert_text((50, 720), "Note: This is a sample for demonstration. Item numbers 1-15 supported.", 
                     fontsize=9, fontname="helv", color=(0.5, 0.5, 0.5))

    # Convert to bytes
    pdf_bytes = doc.write()
    doc.close()
    return pdf_bytes


def get_sample_image() -> bytes:
    """Return a placeholder sample image (you can replace with real base64 later)."""
    # For now we create a simple colored rectangle with text using PIL
    img = Image.new('RGB', (800, 400), color='#1e2937')
    draw = ImageDraw.Draw(img)
    
    draw.text((100, 150), "SAMPLE BOQ TABLE IMAGE", fill="#58a6ff", font=None, size=40)
    draw.text((100, 220), "Upload your own image of BOQ for reference", fill="#94a3b8", size=20)
    draw.text((100, 260), "Or click 'Load Sample BOQ' to test the extractor", fill="#94a3b8", size=18)

    buf = BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# ==================== MODIFY YOUR MAIN() FUNCTION ====================

def main():
    load_css()
    initialize_session_state()
    
    # Sidebar (keep your existing sidebar, just add one line at the bottom if you want)
    with st.sidebar:
        st.markdown("### ⚙️ Settings")
        # ... your existing sidebar code ...

        st.markdown("---")
        if st.button("🗑️ Clear All", type="secondary", use_container_width=True):
            # your existing clear code
            st.rerun()

    # ==================== NEW UPLOAD SECTION ====================
    st.markdown("<div class='section-header'>📤 Upload or Try Sample</div>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([2, 1.2, 1.2])

    with col1:
        uploaded_file = st.file_uploader(
            "Upload your BOQ PDF",
            type=["pdf"],
            label_visibility="collapsed",
            help="Upload the actual BOQ drawing PDF"
        )

    with col2:
        if st.button("📋 Load Sample BOQ", use_container_width=True, type="primary"):
            with st.spinner("Loading sample BOQ..."):
                sample_pdf = get_sample_boq_pdf()
                st.session_state.pdf_bytes = sample_pdf
                st.session_state.filename = "Sample_BOQ_v17.pdf"
                st.success("✅ Sample BOQ loaded successfully! Click 'Extract All Pages' to test.")
                st.rerun()

    with col3:
        uploaded_image = st.file_uploader(
            "Upload Sample Image (PNG/JPG)",
            type=["png", "jpg", "jpeg"],
            label_visibility="collapsed",
            help="Upload an image of BOQ table as reference"
        )

    # Handle PDF Upload
    if uploaded_file and uploaded_file != st.session_state.get("last_uploaded_file"):
        try:
            pdf_bytes = uploaded_file.read()
            st.session_state.pdf_bytes = pdf_bytes
            st.session_state.filename = uploaded_file.name
            st.success(f"✓ Loaded: {uploaded_file.name}")
            st.session_state.last_uploaded_file = uploaded_file  # prevent re-trigger
        except Exception as e:
            st.error(f"Error reading PDF: {e}")

    # Handle Sample Image Upload (for future expansion - display + optional OCR)
    if uploaded_image:
        st.info("📸 Image uploaded. In future versions this can be used for image-based extraction.")
        st.image(uploaded_image, caption="Uploaded BOQ Reference Image", use_column_width=True)

    # Rest of your existing code remains the same...
    # (Preview, Extract, Results, etc.)

    # Show filename if loaded
    if st.session_state.get("filename"):
        st.caption(f"Current file: **{st.session_state.filename}**")

    # ====================== YOUR EXISTING CODE CONTINUES HERE ======================
    # col_left, col_right = st.columns([1, 1])   ← You can keep or merge with new layout

    # ... paste all your existing preview, extraction, results, and help sections here ...
