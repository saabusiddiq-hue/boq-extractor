# BOQ Extractor Pro v26 - Correct Understanding

## Important Clarification

**FIG NO is NOT removed from DESCRIPTION!**

### Your Format:
| FIG NO | DESCRIPTION | MATERIAL |
|--------|-------------|----------|
| PTFE | PTFE-140SQx3mm | (empty) |
| Graphite bronze | Graphite bronze-140SQx5mm | (empty) |
| V1-22-BM1 | Variable Effort Support | Per MSS-SP58 |

### Understanding:
- **FIG NO** = What the item is (identifier)
- **DESCRIPTION** = FIG NO + Specifications (this is INTENTIONAL, not duplicate)
- **MATERIAL** = Material grade (only if different from FIG NO)

### Extraction Logic:
1. Extract FIG NO (PTFE, Graphite bronze, V1-22-BM1, etc.)
2. Keep DESCRIPTION as-is (includes FIG NO prefix)
3. Extract MATERIAL only if found at end (A36, SS316, Per MSS-SP58)

### Example:
**Input line:** `1  1  PTFE  PTFE-140SQx3mm`

**Output:**
- Item: 1
- Qty: 1
- Fig No: PTFE
- Description: PTFE-140SQx3mm (kept as-is!)
- Material: (empty)
