# 📋 BOQ Extractor Pro v12.0 - Interactive Preview Edition

**Material-Anchored PDF to Excel/CSV Converter with Visual Preview**

A professional web application for extracting Bill of Quantities (BOQ) from PDF engineering drawings, featuring **interactive visual preview** with adjustable extraction regions and material anchor point recognition.

![Preview Demo](demo_preview.png)

## 🚀 Quick Start

### Windows
1. Download all files to a folder
2. Double-click `launch_v12.bat`
3. App opens in browser at `http://localhost:8501`

### Linux/Mac
```bash
chmod +x launch_v12.sh
./launch_v12.sh
```

### Manual
```bash
pip install -r requirements_v12.txt
streamlit run boq_extractor_app_v12.py
```

## ✨ Key Features

### 🎯 Interactive Visual Preview
- **Real-time region adjustment** with X/Y sliders (0-100%)
- **Visual anchor points** - See exactly where materials are detected
- **Color-coded regions**:
  - 🔵 **Blue box** = Text extraction region
  - 🔴 **Red boxes** = Auto-detected material anchors
  - ⚓ **Labels** = Material codes (A240 SS316, etc.)

### ⚓ Material-Anchor Engine
Uses Material column as parsing anchor, working right-to-left:
```
[Item] [Qty] [Part No] [Description] ⚓[Material]
   1      1   3x240x240  SS Plate      A240 SS316
```

### 📐 Adjustable Extraction Regions
| Setting | Default | Range | Purpose |
|---------|---------|-------|---------|
| X Range | 50%-98% | 0%-100% | Left-right table position |
| Y Range | 8%-45% | 0%-100% | Top-bottom table position |

### 🔍 Smart Material Detection
Auto-detects these material anchors:
- ✅ **A240 SS316** - Stainless steel
- ✅ **A36, A105** - Carbon steel grades  
- ✅ **A193 GR.B7, A194 GR.2H** - Bolt/nut grades
- ✅ **Per MSS-SP58** - Standard reference
- ✅ **CARBON STEEL, A516 GR.60** - Material specs
- ✅ **Graphite bronze** - Special materials

## 📖 Step-by-Step Guide

### Step 1: Upload PDF
Drag & drop your BOQ PDF into the upload area.

### Step 2: Adjust Extraction Region
Use the **sidebar sliders** to position the blue extraction box:
- Drag X Range to cover the table horizontally
- Drag Y Range to cover the table vertically
- **Default** (50%-98% X, 8%-45% Y) works for most right-side tables

### Step 3: Preview with Anchor Points
Click **"👁️ Preview"** to see:
- Your extraction region (blue box)
- Detected material anchors (red boxes with ⚓ labels)
- Extracted text with highlighted materials
- Parsing analysis (how many items detected)

### Step 4: Verify & Extract
If preview shows ✅ items detected, click **"🔍 Extract BOQ"** to process all pages.

### Step 5: Export
Download as Excel (yellow headers) or CSV.

## 🎯 Understanding the Visual Preview

```
┌─────────────────────────────────────────┐
│         PDF Page                        │
│  ┌─────────────────────┐               │
│  │                     │               │
│  │   Table Region      │ ← 🔵 Blue Box │
│  │   (Extraction)      │   (Sliders    │
│  │                     │    control    │
│  │   3x240x240 SS Plate│    this)      │
│  │   ⚓ A240 SS316 ←───┼── 🔴 Red Box   │
│  │                     │   (Material   │
│  │   V2-26-BM1         │    Anchor)    │
│  │   ⚓ Per MSS-SP58 ←─┼── ⚓ Label      │
│  └─────────────────────┘               │
│                                         │
│   POS-... (Mark No)                    │
│   Q24250 SUPP 05-021 (Drawing No)      │
└─────────────────────────────────────────┘
```

## 🛠️ Troubleshooting

### No materials detected in preview?
| Cause | Solution |
|-------|----------|
| Region too small | Expand X Range to include rightmost column |
| Region misaligned | Adjust Y Range to cover table rows |
| PDF is scanned image | Use OCR-enabled PDF or text-based PDF |

### Wrong item count?
- Ensure ⚓ material anchors appear on every row
- Check that all rows have complete material codes
- Adjust region to avoid header/footer text

### Missing some drawings?
- Check Processing Logs for "⚠️" warnings
- Verify page range covers all pages
- Ensure title blocks are in bottom-right

## 📋 Input Format Supported

```
Item No  Qty  Part No        Description              Material
1        1    3x240x240      SS Plate(Mirror Finish)  A240 SS316
2        1    V2-26-BM1      Variable Effort Support  Per MSS-SP58
3        1    GraphiteBronze Graphite bronze-190SQ    (blank)
```

The system uses the **Material column** (rightmost) as the anchor to parse each row.

## 📦 File Structure

```
boq_extractor_v12/
├── boq_extractor_app_v12.py    # Main application (36 KB)
├── requirements_v12.txt        # Dependencies
├── launch_v12.bat             # Windows launcher
├── launch_v12.sh              # Linux/Mac launcher
└── README_v12.md              # This file
```

## 🔧 System Requirements

- **OS**: Windows 10/11, macOS 10.15+, Ubuntu 18.04+
- **Python**: 3.8 or higher
- **RAM**: 4GB minimum (8GB for large PDFs)
- **Browser**: Chrome 90+, Firefox 88+, Safari 14+, Edge 90+

## 🔄 Version History

- **v12.0**: Interactive preview with visual anchor points and adjustable regions
- **v11.0**: Complete web interface with material-anchored extraction
- **v10.0**: Material-anchored parsing engine

---

**Made for engineers who need precise BOQ extraction with visual verification.**
