# BOQ Extractor Pro v28 - Material Separation

## Problem Solved
**Material was merging into Description column**

### Before (Wrong):
| Description | Material |
|-------------|----------|
| bronze Graphite bronze-190SQx5mm | (empty) |
| Variable Effort Support Per MSS-SP58 | (empty) |
| Inverted Beam Welding Attachment A36 | (empty) |

### After (Correct):
| Description | Material |
|-------------|----------|
| bronze Graphite bronze-190SQx5mm | (empty - no material spec) |
| Variable Effort Support | Per MSS-SP58 |
| Inverted Beam Welding Attachment | A36 |
| Full Nut | A194 GR.2H |
| Weldless Eye Nut | A105 |
| Pipe Clamp 2 Bolt | A36 |

## How It Works

### Material Separation Logic:
1. Extract full text after Fig No
2. Search for material patterns from **END** of text:
   - `Per MSS-SP58`
   - `A194 GR.2H`
   - `A36`
   - `SS316`
   - etc.
3. Split text at material position
4. Everything before = Description
5. Material pattern = Material column

### Material Patterns Detected:
- Per MSS-SP58, MSS-SP58
- A194 GR.2H, A193 GR.B7
- A240 SS316, SS316, SS316L
- A36, A105, A516
- Gr. 8.8, Grade 8.8
- CI. 8, Cast Iron
- And more...

## Usage
1. Upload PDF
2. Adjust crop to include all columns
3. Click EXTRACT
4. Material automatically separated from Description
