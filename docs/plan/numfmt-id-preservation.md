# NumFmt ID Preservation Plan

**Status**: Ready for Implementation
**Priority**: P0 (Blocks surgical modification correctness)
**Complexity**: Medium
**Estimated Effort**: 4-6 hours

---

## Problem Statement

Surgical modification of XLSX files corrupts number formatting in two ways:

1. **Format ID loss**: Style 283 has `numFmtId="0"` (General) instead of `numFmtId="39"` (Accounting)
   - Cells display with wrong precision: "2.7262298" instead of "2.72"
   - Excel shows "General" format instead of "Number" format

2. **Style duplication**: Output has 648 styles instead of 647
   - One extra style created during surgical write
   - Indicates deduplication failure

**Impact**: Real-world files (Syndigo) show incorrect number formatting after surgical modification, breaking financial reports.

---

## Background: What We Tried and Why It Failed

### Failed Attempt (Commits c1d62db, 35f38a8)

**Changes made**:
1. Added `numFmtId: Option[Int]` field to `CellStyle`
2. Updated parse logic to capture numFmtId from XML
3. Updated write logic to use preserved numFmtId
4. **CRITICAL ERROR**: Modified `canonicalKey` to include numFmtId

```scala
// The mistake (35f38a8):
val numFmtKey = s"N:${numFmt},${numFmtId.fold("_")(_.toString)}"
```

**Why it failed**:

Including `numFmtId` in `canonicalKey` broke style deduplication:

```scala
// Style A (from source): numFmt=General, numFmtId=Some(0)
// Style B (programmatic): numFmt=General, numFmtId=None

// With numFmtId in canonicalKey:
//   canonicalKey A: "N:General,0"
//   canonicalKey B: "N:General,_"
//   ❌ Different keys → NO deduplication → TWO styles created

// Without numFmtId in canonicalKey (correct):
//   canonicalKey A: "N:General"
//   canonicalKey B: "N:General"
//   ✅ Same key → Deduplication → ONE style
```

**Result**: 648 styles instead of 647, AND complete formatting corruption (cells reference wrong style indices).

### Why Tests Didn't Catch It

1. **Existing tests only checked determinism**, not deduplication semantics:
   ```scala
   test("canonicalKey is deterministic") {
     assertEquals(style.canonicalKey, style.canonicalKey) // ✅ Passes even when broken
   }
   ```

2. **No test for style count stability** during surgical write

3. **No test verifying deduplication works** (same visual appearance → same key)

4. **Manual inspection only** (Syndigo output checked by hand, not automated)

---

## Root Cause Analysis

### The Fundamental Design Principle

**`canonicalKey` represents VISUAL EQUIVALENCE, not byte-perfect equivalence.**

Purpose from CellStyle.scala:
> "Two styles with the same key are structurally equivalent and should map to the same style index in styles.xml."

### Why numFmtId Doesn't Belong in canonicalKey

1. **numFmtId is an implementation detail**, not a visual property:
   - `numFmtId=0` (General) and `numFmtId=None` (default) look IDENTICAL in Excel
   - `numFmtId=39` and custom format 164 can both represent "#,##0.00" (same visual)

2. **Deduplication must be based on appearance**, not provenance:
   - Programmatic style should deduplicate with identical style from source
   - The fact that one has `numFmtId=Some(39)` and other has `None` is irrelevant

3. **Preservation is a separate concern** from deduplication:
   - canonicalKey → deduplication (visual equivalence)
   - numFmtId → preservation (byte-perfect surgical write)

### The Real Bug

The actual problem is **incomplete `NumFmt.builtInById` map**:

```scala
// Current state (db456c8):
val builtInById: Map[Int, NumFmt] = Map(
  0 -> General,
  1 -> NumFmt.Custom("0"),
  2 -> NumFmt.Custom("0.00"),
  // ... IDs 3-22 mapped ...
  // ❌ IDs 23-48 MISSING (including 39, 40, 41, 42)
  49 -> NumFmt.Custom("@")
)
```

**What happens when we parse `numFmtId="39"`**:

```scala
// Styles.scala parseCellStyle():
val numFmt = numFmtIdOpt
  .flatMap(id => numFmts.get(id).orElse(NumFmt.fromId(id))) // fromId(39) returns None
  .getOrElse(NumFmt.General) // ❌ Falls back to General!
```

**Result**: Style with `numFmtId=39` gets parsed as `NumFmt.General`, then written back as `numFmtId=0`.

---

## The Correct Solution

### Strategy

**Separate concerns**:
1. **Deduplication** (canonicalKey): Based on visual properties ONLY
2. **Preservation** (numFmtId): Stored separately, used during write

**Fix the root cause**:
3. **Complete NumFmt.builtInById map**: Add missing IDs 23-48

### Implementation Plan

#### Phase 1: Add numFmtId Field (Without Breaking canonicalKey)

**File**: `xl-core/src/com/tjclp/xl/style/CellStyle.scala`

```scala
case class CellStyle(
  font: Font = Font.default,
  fill: Fill = Fill.default,
  border: Border = Border.none,
  numFmt: NumFmt = NumFmt.General,
  numFmtId: Option[Int] = None,  // NEW: preserve raw Excel format ID
  align: Align = Align.default
):
  // ... existing methods ...

  def withNumFmt(n: NumFmt): CellStyle =
    copy(numFmt = n, numFmtId = None)  // Clear numFmtId when changing numFmt

  def withNumFmtId(id: Int, fmt: NumFmt): CellStyle =
    copy(numFmt = fmt, numFmtId = Some(id))  // Advanced: set both

  // CRITICAL: canonicalKey does NOT include numFmtId
  lazy val canonicalKey: String =
    val numFmtKey = s"N:${numFmt}"  // ✅ Only numFmt, not numFmtId
    // ... rest unchanged ...
```

**Rationale**:
- `numFmtId` stored for preservation
- `canonicalKey` unchanged (visual equivalence only)
- `withNumFmt` clears `numFmtId` (programmatic changes lose provenance)

#### Phase 2: Update Parse Logic

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala`

```scala
private def parseCellStyle(
  xfElem: Elem,
  fonts: Vector[Font],
  fills: Vector[Fill],
  borders: Vector[Border],
  numFmts: Map[Int, NumFmt]
): CellStyle =
  // ... existing font/fill/border parsing ...

  val numFmtIdOpt = xfElem.attribute("numFmtId")
    .flatMap(_.text.toIntOption)  // ✅ Capture raw ID

  val numFmt = numFmtIdOpt
    .flatMap(id => numFmts.get(id).orElse(NumFmt.fromId(id)))
    .getOrElse(NumFmt.General)

  CellStyle(
    font = font,
    fill = fill,
    border = border,
    numFmt = numFmt,
    numFmtId = numFmtIdOpt,  // ✅ Preserve for surgical write
    align = align
  )
```

**Key point**: Capture numFmtId even if we don't recognize it (for byte-perfect preservation).

#### Phase 3: Update Write Logic

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala`

```scala
val cellXfsElem = elem("cellXfs", "count" -> index.cellStyles.size.toString)(
  index.cellStyles.map { style =>
    val fontIdx = fontMap.getOrElse(style.font, -1)
    val fillIdx = fillMap.getOrElse(style.fill, -1)
    val borderIdx = borderMap.getOrElse(style.border, -1)

    val numFmtId = style.numFmtId.getOrElse {  // ✅ Use preserved ID if available
      // No raw ID → derive from NumFmt enum (programmatic creation)
      NumFmt.builtInId(style.numFmt).getOrElse(
        index.numFmts.find(_._2 == style.numFmt).map(_._1).getOrElse(0)
      )
    }

    elem("xf",
      "numFmtId" -> numFmtId.toString,
      "fontId" -> fontIdx.toString,
      // ... other attributes ...
    )(alignmentChild*)
  }*
)
```

**Rationale**:
- Preserved numFmtId used when available (surgical write)
- Fallback to derived ID for programmatic styles (normal write)

#### Phase 4: Complete NumFmt.builtInById Map

**File**: `xl-core/src/com/tjclp/xl/style/numfmt/NumFmt.scala`

Add missing Excel built-in format IDs 23-48:

```scala
val builtInById: Map[Int, NumFmt] = Map(
  0 -> General,
  1 -> NumFmt.Custom("0"),
  // ... existing 2-22 ...

  // ✅ Add missing IDs 23-48 (from ECMA-376 Part 1, §18.8.30):
  23 -> NumFmt.Custom("General"),
  // Date formats
  27 -> NumFmt.Custom("[$-404]e/m/d"),
  28 -> NumFmt.Custom("[$-404]e\"年\"m\"月\"d\"日\""),
  // ... (see ECMA-376 spec for complete list) ...

  // Accounting formats (THE CRITICAL ONES):
  39 -> NumFmt.Custom("#,##0.00;(#,##0.00)"),  // ✅ Fixes the bug!
  40 -> NumFmt.Custom("#,##0.00;[Red](#,##0.00)"),
  41 -> NumFmt.Custom("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"),
  42 -> NumFmt.Custom("_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)"),
  43 -> NumFmt.Custom("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"),
  44 -> NumFmt.Custom("_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)"),
  // ... 45-48 for time formats ...

  49 -> NumFmt.Custom("@")
)
```

**Reference**: ECMA-376 Part 1, §18.8.30 "numFmt (Number Format)"

#### Phase 5: Update CellValue.Number Pattern Matches

**Files affected**:
- `xl-core/src/com/tjclp/xl/codec/CellCodec.scala`
- `xl-core/src/com/tjclp/xl/macros/FormattedLiterals.scala`
- `xl-core/src/com/tjclp/xl/html/HtmlRenderer.scala`
- `xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala`

```scala
// Before:
case CellValue.Number(n) => // handle number

// After:
case CellValue.Number(n, _) => // ignore originalString in most cases

// In Worksheet.scala write logic:
case CellValue.Number(num, originalStr) =>
  val strValue = originalStr.getOrElse(num.toString)  // ✅ Preserve original
  Seq(elem("v")(Text(strValue)))
```

---

## Test Strategy

### Critical New Tests

#### Test 1: canonicalKey Ignores numFmtId

**File**: `xl-core/test/src/com/tjclp/xl/style/StyleSpec.scala`

```scala
test("canonicalKey: ignores numFmtId (ensures deduplication)") {
  val style1 = CellStyle(
    numFmt = NumFmt.General,
    numFmtId = Some(0)
  )
  val style2 = CellStyle(
    numFmt = NumFmt.General,
    numFmtId = None
  )
  val style3 = CellStyle(
    numFmt = NumFmt.General,
    numFmtId = Some(39)  // Different ID, same visual (after completing builtInById)
  )

  assertEquals(style1.canonicalKey, style2.canonicalKey,
    "Styles with same visual properties must have same canonicalKey regardless of numFmtId")

  // Note: style3 would have different canonicalKey because numFmt would be Accounting,
  // not General, once we complete the builtInById map
}
```

#### Test 2: Surgical Write Preserves Exact Style Count

**File**: `xl-ooxml/test/src/com/tjclp/xl/ooxml/XlsxWriterCorruptionRegressionSpec.scala`

```scala
test("surgical write: preserves exact style count (no duplicates)") {
  // Create source with known style count
  val source = createWorkbookWithAccountingFormats()  // Has 3 styles

  // Read, modify one cell value (not style), write
  val wb = XlsxReader.read(source).toOption.get
  val modified = wb("Sheet1").flatMap(
    _.put(ref"A1" -> "Modified text")  // Value change only
  ).flatMap(wb.put).toOption.get

  val output = Files.createTempFile("test", ".xlsx")
  XlsxWriter.write(modified, output)

  // Extract style counts
  val sourceStyleCount = extractStyleCount(source)
  val outputStyleCount = extractStyleCount(output)

  assertEquals(outputStyleCount, sourceStyleCount,
    s"Style count must remain stable (got $outputStyleCount, expected $sourceStyleCount)")
}

private def extractStyleCount(path: Path): Int = {
  val zip = new ZipFile(path.toFile)
  try {
    val stylesXml = readEntryString(zip, zip.getEntry("xl/styles.xml"))
    val doc = parseXml(stylesXml)
    val cellXfs = doc.getElementsByTagName("cellXfs").item(0).asInstanceOf[Element]
    cellXfs.getAttribute("count").toInt
  } finally zip.close()
}
```

#### Test 3: NumFmtId Preserved for Unmodified Cells

**File**: `xl-ooxml/test/src/com/tjclp/xl/ooxml/XlsxWriterCorruptionRegressionSpec.scala`

```scala
test("surgical write: preserves numFmtId for unmodified cells") {
  val source = createWorkbookWithAccountingFormats()

  // Read, modify DIFFERENT cell, write
  val wb = XlsxReader.read(source).toOption.get
  val modified = wb("Sheet1").flatMap(
    _.put(ref"C1" -> "Other cell")  // Don't touch A1, B1
  ).flatMap(wb.put).toOption.get

  val output = Files.createTempFile("test", ".xlsx")
  XlsxWriter.write(modified, output)

  // Verify styles preserved
  val outputZip = new ZipFile(output.toFile)
  val stylesXml = readEntryString(outputZip, outputZip.getEntry("xl/styles.xml"))
  val doc = parseXml(stylesXml)
  val cellXfs = doc.getElementsByTagName("cellXfs")
  val xfNodes = cellXfs.item(0).asInstanceOf[Element].getElementsByTagName("xf")

  // Style 1 should still have numFmtId=39 (accounting)
  val xf1 = xfNodes.item(1).asInstanceOf[Element]
  assertEquals(xf1.getAttribute("numFmtId"), "39",
    "Style 1 numFmtId must be preserved after surgical write")

  // Style 2 should still have numFmtId=40
  val xf2 = xfNodes.item(2).asInstanceOf[Element]
  assertEquals(xf2.getAttribute("numFmtId"), "40",
    "Style 2 numFmtId must be preserved after surgical write")
}
```

#### Test 4: NumFmt.fromId Recognizes All Built-In IDs

**File**: `xl-core/test/src/com/tjclp/xl/style/StyleSpec.scala`

```scala
test("NumFmt.fromId recognizes all built-in format IDs") {
  // Test critical accounting formats
  assert(NumFmt.fromId(39).isDefined, "ID 39 (accounting) must be recognized")
  assert(NumFmt.fromId(40).isDefined, "ID 40 (accounting red) must be recognized")
  assert(NumFmt.fromId(41).isDefined, "ID 41 (accounting with $) must be recognized")

  // Test that unrecognized IDs return None (as expected)
  assert(NumFmt.fromId(200).isEmpty, "Custom ID 200 should not be built-in")
}
```

---

## Validation Plan

### Automated Tests

1. Run full test suite: `./mill __.test`
   - All existing tests must pass
   - New tests must pass

### Manual Validation: Syndigo File

```bash
scala-cli run data/surgical-demo.sc
```

**Expected results**:
- ✅ Write successful
- ✅ **Style count: 647** (not 648)
- ✅ **Style 283: numFmtId="39"** (not 0)
- ✅ Cell displays "2.72" (not "2.7262298")
- ✅ Excel shows "Number" format (not "General")

### Verification Commands

```bash
# Check style count
unzip -p syndigo-surgical-output.xlsx xl/styles.xml | xmllint --format - | grep '<cellXfs count='

# Check style 283 numFmtId
unzip -p syndigo-surgical-output.xlsx xl/styles.xml | xmllint --format - | sed -n '/<cellXfs/,/<\/cellXfs>/p' | grep -n '<xf' | sed -n '284p'
```

---

## Implementation Checklist

- [ ] Phase 1: Add numFmtId field to CellStyle (without breaking canonicalKey)
- [ ] Phase 2: Update Styles.scala parse logic
- [ ] Phase 3: Update Styles.scala write logic
- [ ] Phase 4: Complete NumFmt.builtInById map (add IDs 23-48)
- [ ] Phase 5: Update CellValue.Number pattern matches
- [ ] Test 1: canonicalKey ignores numFmtId
- [ ] Test 2: Surgical write preserves exact style count
- [ ] Test 3: NumFmtId preserved for unmodified cells
- [ ] Test 4: NumFmt.fromId recognizes all built-in IDs
- [ ] Run full test suite
- [ ] Manual validation with Syndigo file
- [ ] Commit with message: "Fix: Preserve numFmt IDs in surgical modification (647 styles, correct formatting)"

---

## Success Criteria

1. **Surgical write style count stable**: 647 styles in → 647 styles out (not 648)
2. **NumFmt IDs preserved**: Style 283 has `numFmtId="39"` after surgical write
3. **Visual formatting correct**: Cells display with proper precision/formatting in Excel
4. **All tests pass**: Existing + new tests all green
5. **No regression**: Fonts, fills, borders, alignment all preserved correctly

---

## References

- **ECMA-376 Part 1, §18.8.30**: Complete list of built-in number format IDs
- **XL Purity Charter**: `docs/design/purity-charter.md` (totality, determinism)
- **Style Guide**: `docs/design/style-guide.md` (opaque types, law-governed design)
- **Surgical Modification**: Commit 2681f95 (unified write path)
