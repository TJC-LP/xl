# NumFmt Preservation for Byte-Perfect Surgical Modification

**Status**: Planned (Not Yet Implemented)
**Priority**: P0 - BLOCKING for surgical modification
**Estimated Effort**: 2-3 hours

## Problem Statement

Surgical modification is losing number format styling, causing Excel to display numbers with incorrect decimal precision.

### Observable Symptoms

**Original Excel file:**
- Cell displays: `2.72` (2 decimals)
- Formula bar: `2.72622981509072` (full precision)
- Format tooltip: "Number" (custom format applied)
- XML: `<c r="M13" s="283"><v>2.7262298150907198</v></c>`
- Style 283: `numFmtId="39"` (accounting format `#,##0.00;(#,##0.00)`)

**After surgical write:**
- Cell displays: `2.7262298` (wrong precision!)
- Formula bar: `2.72622981509072` (correct)
- Format tooltip: "General" (format LOST!)
- XML: `<c r="M13" s="283"><v>2.7262298150907198</v></c>` (styleId preserved)
- Style 283: `numFmtId="0"` (General - CORRUPTED!)

### Impact

- **Visual corruption**: Numbers display with wrong decimal places
- **User confusion**: Spreadsheet looks "broken" even though values are correct
- **Defeats surgical modification goal**: Not byte-perfect, not format-preserving
- **Critical for financial models**: Beta, WACC, ratios all need specific precision

---

## Root Cause Analysis

### Discovery Process

1. **Initial hypothesis**: `BigDecimal.toString` losing precision
   - **Finding**: Value precision IS preserved (`2.7262298150907198` matches exactly)
   - **Conclusion**: Not the issue

2. **Second hypothesis**: Number format codes not preserved in styles.xml
   - **Finding**: Custom formats ARE preserved (66 numFmts in both source and output)
   - **Conclusion**: Styles.xml has the formats, but not applied correctly

3. **Third hypothesis**: StyleId not preserved for cells
   - **Finding**: Cell `s="283"` attribute IS preserved
   - **Conclusion**: Cell points to correct style index

4. **ROOT CAUSE DISCOVERED**: Style at index 283 is **completely different** in output!

```
SOURCE Style 283: numFmtId="39" fontId="6" applyNumberFormat="1"
OUTPUT Style 283: numFmtId="0"  fontId="1" (completely replaced!)
```

### Why This Happens

**The Bug Chain:**

1. **Parsing** (`Styles.scala:689-691`):
   ```scala
   val numFmtId = xfElem.attribute("numFmtId").flatMap(_.text.toIntOption)  // Some(39)
   val numFmt = numFmtId
     .flatMap(id => numFmts.get(id)              // Not in custom formats
     .orElse(NumFmt.fromId(id)))                 // Tries built-in map
     .getOrElse(NumFmt.General)                  // Falls back to General!
   ```

2. **NumFmt.fromId(39)** (`NumFmt.scala:91-92`):
   ```scala
   def fromId(id: Int, formatCode: Option[String] = None): Option[NumFmt] =
     builtInById.get(id).orElse(formatCode.map(parse))
   ```

3. **builtInById map** (`NumFmt.scala:29-53`):
   ```scala
   private val builtInById: Map[Int, NumFmt] = Map(
     0 -> General,
     1 -> Integer,
     ...
     22 -> DateTime,
     49 -> Text
     // ❌ MISSING: IDs 23-48 (including 39!)
   )
   ```

4. **Writing** (`Styles.scala:349-353`):
   ```scala
   val numFmtId = NumFmt.builtInId(style.numFmt)  // General → Some(0)
     .getOrElse(
       index.numFmts.find(_._2 == style.numFmt).map(_._1).getOrElse(0)
     )
   // Writes numFmtId="0" even though original was "39"!
   ```

### Missing Built-In Formats

Excel has **164 built-in format IDs** (0-163), but XL only maps ~20 of them:

| ID Range | Formats | Status in XL |
|----------|---------|--------------|
| 0-4 | General, Integer, Decimal, Thousands | ✅ Mapped |
| 5-8 | Currency variants | ✅ Mapped |
| 9-13 | Percent, Scientific, Fraction | ✅ Mapped |
| 14-22 | Date/Time variants | ✅ Mapped |
| **23-36** | Reserved | ❌ Missing |
| **37-44** | **Accounting formats** | ❌ **MISSING (Bug!)** |
| **45-48** | Time/Currency variants | ❌ Missing |
| 49 | Text | ✅ Mapped |
| 50-163 | Locale-specific, reserved | ❌ Missing |
| 164+ | Custom (user-defined) | ✅ Handled separately |

---

## Recommended Solution

### Architecture: Dual Representation (Like Number.originalString)

**Current:**
```scala
case class CellStyle(
  font: Font,
  fill: Fill,
  border: Border,
  numFmt: NumFmt,  // Enum: General, Currency, Date, Custom(code)
  align: Align
)
```

**Proposed:**
```scala
case class CellStyle(
  font: Font,
  fill: Fill,
  border: Border,
  numFmt: NumFmt,              // Semantic meaning (for programmatic use)
  numFmtId: Option[Int] = None, // Raw Excel ID (for byte-perfect preservation)
  align: Align
)
```

### Read Logic (Styles.scala:667-693)

**Current:**
```scala
val numFmtId = xfElem.attribute("numFmtId").flatMap(_.text.toIntOption)
val numFmt = numFmtId
  .flatMap(id => numFmts.get(id).orElse(NumFmt.fromId(id)))
  .getOrElse(NumFmt.General)  // ❌ Loses ID 39!

CellStyle(font, fill, border, numFmt, align)
```

**Fixed:**
```scala
val numFmtIdOpt = xfElem.attribute("numFmtId").flatMap(_.text.toIntOption)
val numFmt = numFmtIdOpt
  .flatMap(id => numFmts.get(id).orElse(NumFmt.fromId(id)))
  .getOrElse(NumFmt.General)  // Best-effort semantic mapping

// Preserve BOTH the enum AND the raw ID
CellStyle(font, fill, border, numFmt, numFmtIdOpt, align)
```

### Write Logic (Styles.scala:349-353)

**Current:**
```scala
val numFmtId = NumFmt.builtInId(style.numFmt).getOrElse(
  index.numFmts.find(_._2 == style.numFmt).map(_._1).getOrElse(0)
)
```

**Fixed:**
```scala
val numFmtId = style.numFmtId.getOrElse {
  // No raw ID → derive from NumFmt enum (programmatic creation)
  NumFmt.builtInId(style.numFmt).getOrElse(
    index.numFmts.find(_._2 == style.numFmt).map(_._1).getOrElse(0)
  )
}
```

---

## Implementation Plan

### Phase 1: Extend CellStyle (Core Changes)

**1. Update CellStyle definition** (`xl-core/src/com/tjclp/xl/style/CellStyle.scala`)
```scala
case class CellStyle(
  font: Font = Font.default,
  fill: Fill = Fill.default,
  border: Border = Border.none,
  numFmt: NumFmt = NumFmt.General,
  numFmtId: Option[Int] = None,  // NEW
  align: Align = Align.default
) derives CanEqual
```

**2. Update builders** (same file)
```scala
def withNumFmt(fmt: NumFmt): CellStyle =
  copy(numFmt = fmt, numFmtId = None)  // Clear ID when changing format

// NEW: Set explicit ID (advanced use case)
def withNumFmtId(id: Int, fmt: NumFmt): CellStyle =
  copy(numFmt = fmt, numFmtId = Some(id))
```

**3. Update canonicalKey** (`CellStyle.scala`)
- **Decision**: Should numFmtId be part of canonical key?
- **Answer**: No - two styles with same visual appearance but different IDs should deduplicate
- **Implementation**: Canonical key uses `numFmt` only, not `numFmtId`

### Phase 2: Update Parsing (OOXML Layer)

**4. Update Styles.scala:667-693** (parseCellStyle)
```scala
val numFmtIdOpt = xfElem.attribute("numFmtId").flatMap(_.text.toIntOption)
val numFmt = numFmtIdOpt
  .flatMap(id => numFmts.get(id).orElse(NumFmt.fromId(id)))
  .getOrElse(NumFmt.General)

CellStyle(
  font = font,
  fill = fill,
  border = border,
  numFmt = numFmt,
  numFmtId = numFmtIdOpt,  // NEW: Preserve raw ID
  align = align
)
```

### Phase 3: Update Serialization (OOXML Layer)

**5. Update Styles.scala:349-353** (toXml)
```scala
val numFmtId = style.numFmtId.getOrElse {
  // Derive from NumFmt enum for programmatic styles
  NumFmt.builtInId(style.numFmt).getOrElse(
    index.numFmts.find(_._2 == style.numFmt).map(_._1).getOrElse(0)
  )
}
```

### Phase 4: Update Pattern Matches

**6. Find and update all CellStyle pattern matches** (similar to Number fix)
- Use grep to find: `case CellStyle(font, fill, border, numFmt, align)`
- Update to: `case CellStyle(font, fill, border, numFmt, _, align)` (ignore numFmtId)
- Estimated: 10-15 pattern matches across codebase

### Phase 5: Testing

**7. Add regression test** (`XlsxWriterCorruptionRegressionSpec.scala`)
```scala
test("preserves built-in number format IDs (including accounting formats 37-44)") {
  // Create file with numFmtId=39 (accounting)
  // Surgical write
  // Verify output has numFmtId="39" not "0"
  // Verify Excel displays with correct precision (2 decimals)
}
```

**8. Validate with Syndigo file**
- Run surgical-demo.sc
- Check style 283 has numFmtId=39
- Open in Excel, verify "Number" not "General"

---

## Alternative Solutions (Rejected)

### Option A: Add All Built-In IDs to Map
**Rejected because:**
- Incomplete (Excel has 164+ built-in IDs)
- Fragile (new Excel versions add more)
- Maintenance burden (large map to maintain)
- Still doesn't handle unknown IDs

### Option C: Surgical-Only Fix
**Rejected because:**
- Band-aid solution
- Doesn't fix non-surgical writes
- Creates code complexity (two code paths)

---

## Backward Compatibility

**Breaking Changes:** None!

- `numFmtId` is optional with default `None`
- Existing code: `CellStyle(font, fill, border, numFmt, align)` → works (uses default)
- Pattern matches: Update to ignore numFmtId with `_`
- Programmatic creation: Auto-derives ID (existing behavior)
- Excel reads: Preserves exact ID (new behavior, improves fidelity)

---

## Expected Outcomes

### Before Fix
```
Source: numFmtId=39 (accounting) → displays as "2.72"
Output: numFmtId=0 (general) → displays as "2.7262298" ❌
```

### After Fix
```
Source: numFmtId=39 → CellStyle(..., numFmt=Currency, numFmtId=Some(39))
Output: numFmtId=39 → displays as "2.72" ✅
```

### Programmatic Usage (Unchanged)
```scala
val style = CellStyle.default.withNumFmt(NumFmt.Currency)
// numFmtId = None → auto-derives to 5 (built-in Currency ID)
// Displays correctly with currency formatting ✅
```

---

## Implementation Checklist

- [ ] Update `CellStyle` case class with `numFmtId: Option[Int] = None`
- [ ] Update `withNumFmt` builder to clear numFmtId
- [ ] Update `parseCellStyle` to capture numFmtId
- [ ] Update `toXml` write logic to use numFmtId when available
- [ ] Find and update all `CellStyle` pattern matches (~10-15 locations)
- [ ] Add regression test for built-in format preservation
- [ ] Validate with Syndigo file (style 283 should have numFmtId=39)
- [ ] Run full test suite (712+ tests should pass)
- [ ] Document in CHANGELOG

---

## Files to Modify

### Core Changes
1. `xl-core/src/com/tjclp/xl/style/CellStyle.scala` (add numFmtId field)

### OOXML Changes
2. `xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala` (parse + write)

### Pattern Match Updates (~10-15 files)
3. `xl-core/src/com/tjclp/xl/codec/CellCodec.scala`
4. `xl-core/src/com/tjclp/xl/html/HtmlRenderer.scala`
5. `xl-core/src/com/tjclp/xl/optics/Lenses.scala`
6. Test files (OpticsSpec, StyleSpec, etc.)

### Tests
7. `xl-ooxml/test/src/com/tjclp/xl/ooxml/XlsxWriterCorruptionRegressionSpec.scala` (new test)

---

## Risk Assessment

**Risk Level**: Medium

**Risks:**
- Pattern match exhaustiveness warnings if any matches are missed
- Canonical key behavior change (mitigated: numFmtId not in key)
- Test suite churn (many pattern matches to update)

**Mitigation:**
- Use compiler to find all pattern matches (will error if incomplete)
- Backward compatible default parameter
- Comprehensive test coverage

**Confidence**: High - mirrors successful `Number.originalString` pattern

---

## Similar Patterns in Codebase

This follows the **dual representation pattern** established for `CellValue.Number`:

| Type | Semantic Value | Original Representation | Use Case |
|------|----------------|------------------------|----------|
| `CellValue.Number` | `value: BigDecimal` | `originalString: Option[String]` | Preserve exact numeric string |
| `CellStyle` | `numFmt: NumFmt` | `numFmtId: Option[Int]` | Preserve exact format ID |

**Design Principle**: Store both the **semantic meaning** (for programmatic use) and the **original representation** (for byte-perfect preservation).

---

## Post-Implementation Validation

### Test Cases

1. **Built-in formats 0-22** (currently mapped)
   - Should continue working (ID derived from NumFmt)

2. **Built-in formats 37-44** (accounting - currently broken)
   - Should preserve exact ID
   - Excel should show "Number" not "General"

3. **Custom formats 164+** (currently working)
   - Should continue working (already in index.numFmts)

4. **Programmatic creation**
   - `withNumFmt(NumFmt.Currency)` → derives ID=5
   - No change to user experience

5. **Surgical modification**
   - Read → Write → Byte-perfect preservation
   - All numFmtIds preserved exactly

### Syndigo Validation

```bash
# After fix:
scala-cli run data/surgical-demo.sc

# Verify style 283 has numFmtId=39:
python3 <<'PYEOF'
import xml.etree.ElementTree as ET, zipfile
with zipfile.ZipFile('data/syndigo-surgical-output.xlsx') as zf:
    root = ET.parse(zf.open('xl/styles.xml')).getroot()
    ns = {'': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    xf283 = list(root.find('.//cellXfs', ns))[283]
    print(f"Style 283: numFmtId={xf283.get('numFmtId')}")
    assert xf283.get('numFmtId') == '39', "Should preserve accounting format!"
PYEOF

# Open in Excel → hover over M13 → should show "Number" not "General"
```

---

## Timeline

**Estimated Effort:** 2-3 hours

- Phase 1: Core changes (30 min)
- Phase 2-3: Parse/write updates (30 min)
- Phase 4: Pattern match updates (45 min)
- Phase 5: Testing (30 min)
- Validation: (15 min)
- Buffer: (15 min)

**Blocking For:** PR #16 merge

**Next Steps:**
1. Implement tomorrow morning
2. Validate with Syndigo file
3. Update PR #16 with fix
4. Request final review

---

## Notes

- This bug affects **all** Excel files with non-standard built-in formats (IDs 23-48)
- Common in financial models (accounting formats 37-44)
- Surgical modification revealed this issue (normal writes might have it too)
- Fix improves both surgical AND normal writes (better fidelity overall)

**Key Insight**: Excel's number formatting is **not just about the format code**, but about the **exact format ID**. Different IDs can have similar codes but different Excel behavior (locale, negative handling, etc.). Preserving the ID is critical.

---

**Status**: Ready for implementation (2025-11-18)
**Author**: Claude Code
**Reviewer**: (pending)
