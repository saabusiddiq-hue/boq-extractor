# BOQ Extractor Pro v21.0 - Triple Validation

**Structure Learning + 3-Step Validation System**

## Features

1. **📚 Step 1: Learn** - Upload sample image OR sample PDF to teach the system
2. **⚡ Step 2: Extract** - Process target PDF with learned patterns
3. **🔍 Step 3: Review** - Triple-check materials (auto-fix + manual + final check)
4. **✅ Step 4: Export** - Download Excel/CSV

## What's Fixed

- **Graphite bronze merging**: Auto-detected and separated
- **3-time checking**: Learn → Extract → Review → Final Check
- **Dual sample input**: Image (JPG/PNG) + PDF samples supported
- **Smart material extraction**: Uses learned patterns from your samples

## Deploy

```bash
# Local
pip install -r requirements.txt
streamlit run app.py

# Streamlit Cloud
# Push to GitHub → Deploy with main file: app.py
```

## Usage

1. Upload sample BOQ (image or PDF) → Click "Learn"
2. Upload target PDF → Click "Extract"
3. Review materials (red borders = needs fix)
4. Final check → Export
