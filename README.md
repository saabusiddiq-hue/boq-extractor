# BOQ Extractor Pro v27 - Material Column Fix

## Problem Identified
Material column was empty because:
1. Crop area (Z/Right slider) was too narrow, excluding Material column
2. Material detection was looking in wrong position

## Solution
1. **Warning system**: Shows alert if Z < 85% (Material column likely cut off)
2. **Better parsing**: Material column detected from END of line
3. **Debug mode**: View raw extracted text to verify columns
4. **Stats**: Shows how many items have materials detected

## How to Fix Empty Material Column

### Step 1: Check Crop Area
- **Z (Right) slider**: Must be 90-95% to include Material column
- If Z < 85%, you'll see red warning: "Expand Z to include Material column"

### Step 2: Verify with Debug
- Enable "Debug" checkbox
- Check "Raw Text" to see if Material column text appears

### Step 3: Material Detection
Material keywords detected:
- A36, A105, A193, A194, A240, A516
- SS316, SS316L, SS304, SS304L
- Per MSS-SP58, MSS-SP58
- Gr. 8.8, Grade 8.8
- CI. 8, Cast Iron
- Bronze, Graphite, PTFE

## Expected Output
| Item | Qty | Fig No | Description | Material |
|------|-----|--------|-------------|----------|
| 1 | 1 | PTFE | PTFE-140SQx3mm | (empty) |
| 2 | 1 | V1-22-BM1 | Variable Effort Support | Per MSS-SP58 |
| 3 | 1 | 3x190x190 | SS Plate(mirror finish) | A240 SS316 |
