# P4.5: OOXML Quality & Spec Compliance

**Status**: ⬜ Not Started
**Priority**: High (Spec violations cause data loss and Excel incompatibility)
**Estimated Effort**: 6-8 hours with comprehensive testing
**Test Impact**: +10 new tests (698 → 708 total)

## Executive Summary

This phase addresses 6 quality and correctness issues identified in the style review (`docs/reviews/style-review.md`). While the core OOXML read/write pipeline is functional, there are several spec violations and data loss scenarios that must be fixed before the library can be considered production-ready for strict OOXML compliance.

**Why this matters:**
- **Spec conformance**: Excel and other OOXML consumers expect specific structures
- **Data integrity**: Current implementation loses alignment and whitespace data
- **Interoperability**: Non-conformant XML may trigger warnings or errors in strict parsers
- **AI readability**: Adding structured contracts improves LLM reasoning about code behavior

## Critical Issues (High Priority)

### Issue 1: Default Fills (Gray125) ⚠️ Spec Violation

**Location**: `xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala:139`

**Current State**:
```scala
val defaultFills = Vector(Fill.None, Fill.Solid(Color.Rgb(0x00000000)))
```

**Problem**:
- Second default fill uses solid black (`0x00000000`)
- OOXML spec (ECMA-376) requires:
  - Index 0: `patternType="none"`
  - Index 1: `patternType="gray125"`

**Impact**:
- Excel may reject or misinterpret styles.xml
- Cell fills may render incorrectly
- Spec violation blocks strict validation

**Fix Complexity**: Simple (single-line change)

**Fix**:
```scala
// Define gray125 pattern with appropriate colors
val defaultFills = Vector(
  Fill.None,
  Fill.Pattern(
    fgColor = Some(Color.Rgb(0xFF000000)), // Black foreground
    bgColor = Some(Color.Rgb(0xFFC0C0C0)), // Silver background
    patternType = PatternType.Gray125
  )
)
```

**Test Plan**:
- Test: "Styles.defaultFills has None at index 0"
- Test: "Styles.defaultFills has Gray125 at index 1"
- Test: "Styles.toXml emits gray125 pattern in fills section"

**AI Contract Requirements**:
```scala
/** Default fills required by OOXML spec (ECMA-376 Part 1, §18.8.21)
  *
  * REQUIRES: None
  * ENSURES: Vector contains exactly 2 fills:
  *   - Index 0: Fill.None (patternType="none")
  *   - Index 1: Fill.Pattern with patternType="gray125"
  * DETERMINISTIC: Yes (immutable constant)
  * ERROR CASES: None (compile-time constant)
  */
private val defaultFills: Vector[Fill] = ...
```

---

### Issue 2: SharedStrings Count vs UniqueCount ⚠️ Spec Violation

**Location**: `xl-ooxml/src/com/tjclp/xl/ooxml/SharedStrings.scala:32-56`

**Current State**:
```scala
def toXml: Elem =
  elem(
    "sst",
    "xmlns" -> nsSpreadsheetML,
    "count" -> strings.size.toString,        // ❌ WRONG
    "uniqueCount" -> strings.size.toString   // ✅ Correct
  )(siElems*)
```

**Problem**:
- Both attributes set to `strings.size`
- OOXML spec requires:
  - `count` = total number of string cell instances across workbook
  - `uniqueCount` = number of unique strings in SST

**Impact**:
- Excel displays incorrect reference counts
- Spec violation may trigger warnings in strict validators

**Fix Complexity**: Moderate (data model change required)

**Fix**:
1. Update `SharedStrings` case class:
```scala
case class SharedStrings(
  strings: Vector[String],
  indexMap: Map[String, Int],
  totalCount: Int  // NEW: track total instances
)
```

2. Update `fromWorkbook` to count all text cells:
```scala
def fromWorkbook(wb: Workbook): SharedStrings =
  val allTextCells = wb.sheets.flatMap { sheet =>
    sheet.cells.values.collect {
      case cell if cell.value.isInstanceOf[CellValue.Text] =>
        cell.value.asInstanceOf[CellValue.Text].text
    }
  }
  val totalCount = allTextCells.size  // Count all instances
  val unique = allTextCells.distinct.toVector
  SharedStrings(unique, unique.zipWithIndex.toMap, totalCount)
```

3. Update `toXml`:
```scala
def toXml: Elem =
  elem(
    "sst",
    "xmlns" -> nsSpreadsheetML,
    "count" -> totalCount.toString,      // Total instances
    "uniqueCount" -> strings.size.toString  // Unique strings
  )(siElems*)
```

**Test Plan**:
- Test: "SharedStrings.fromWorkbook counts total instances"
- Test: "SharedStrings.toXml emits correct count and uniqueCount"
- Test: "Workbook with duplicate strings has count > uniqueCount"

**AI Contract Requirements**:
```scala
/** Create SharedStrings table from workbook
  *
  * REQUIRES: wb contains only valid CellValue.Text cells
  * ENSURES:
  *   - strings contains unique text values (deduped)
  *   - indexMap maps each string to its SST index
  *   - totalCount = number of text cell instances in workbook
  *   - totalCount >= strings.size (equality when no duplicates)
  * DETERMINISTIC: Yes (iteration order is stable)
  * ERROR CASES: None (total function)
  */
def fromWorkbook(wb: Workbook): SharedStrings = ...
```

---

### Issue 3: Inline String Whitespace Preservation ⚠️ Data Loss

**Locations**:
- `xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala:24-25, 52-56`
- `xl-cats-effect/src/com/tjclp/xl/io/StreamingXmlWriter.scala:42-51, 150-154`

**Current State**:

**RichText**: ✅ Correctly handles `xml:space="preserve"` (lines 52-56, 150-154)

**Plain Text (Worksheet.scala)**: ❌ Does not add `xml:space`
```scala
case CellValue.Text(text) if cellType == "inlineStr" =>
  Seq(elem("is")(elem("t")(Text(text))))  // Missing xml:space check
```

**Plain Text (StreamingXmlWriter.scala)**: ❌ Does not add `xml:space`
```scala
case CellValue.Text(s) =>
  ("inlineStr", List(
    XmlEvent.StartTag(QName("is"), Nil, false),
    XmlEvent.StartTag(QName("t"), Nil, false),  // Missing xml:space
    XmlEvent.XmlString(s, false),
    ...
  ))
```

**Problem**:
- OOXML requires `xml:space="preserve"` for text with:
  - Leading spaces
  - Trailing spaces
  - Multiple consecutive spaces
- Current implementation only handles RichText, not plain Text

**Impact**:
- Leading/trailing/multiple spaces lost in plain text cells
- Data corruption on round-trip

**Fix Complexity**: Simple (add conditional attribute)

**Fix (Worksheet.scala)**:
```scala
case CellValue.Text(text) if cellType == "inlineStr" =>
  val needsPreserve = text.startsWith(" ") || text.endsWith(" ") || text.contains("  ")
  val tElem = if needsPreserve then
    elem("t", "xml:space" -> "preserve")(Text(text))
  else elem("t")(Text(text))
  Seq(elem("is")(tElem))
```

**Fix (StreamingXmlWriter.scala)**:
```scala
case CellValue.Text(s) =>
  val needsPreserve = s.startsWith(" ") || s.endsWith(" ") || s.contains("  ")
  val tAttrs = if needsPreserve then
    List(Attr(QName("xml:space"), List(XmlString("preserve", false))))
  else Nil

  ("inlineStr", List(
    XmlEvent.StartTag(QName("is"), Nil, false),
    XmlEvent.StartTag(QName("t"), tAttrs, false),
    XmlEvent.XmlString(s, false),
    XmlEvent.EndTag(QName("t")),
    XmlEvent.EndTag(QName("is"))
  ))
```

**Test Plan**:
- Test: "Plain text with leading space preserves xml:space (OOXML)"
- Test: "Plain text with trailing space preserves xml:space (OOXML)"
- Test: "Plain text with double spaces preserves xml:space (OOXML)"
- Test: "Plain text with leading space preserves xml:space (streaming)"
- Test: "Plain text with trailing space preserves xml:space (streaming)"
- Test: "Plain text with double spaces preserves xml:space (streaming)"
- Test: "Whitespace round-trips through write/read cycle"

**AI Contract Requirements**:
```scala
/** Convert CellValue.Text to inline string XML events
  *
  * REQUIRES: value is CellValue.Text(s) where s is valid XML text
  * ENSURES:
  *   - Emits <is><t>text</t></is> structure
  *   - Adds xml:space="preserve" when s has leading/trailing/multiple spaces
  *   - Whitespace is preserved byte-for-byte on round-trip
  * DETERMINISTIC: Yes (pure transformation)
  * ERROR CASES: None (total function over CellValue.Text)
  */
private def textToXml(text: String): Seq[XmlEvent] = ...
```

---

### Issue 4: Alignment Serialization ⚠️ Data Loss

**Location**: `xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala:150-171, 429-454`

**Current State**:

**Parsing**: ✅ Correctly reads `<alignment>` elements (lines 429-454)

**Serialization**: ❌ Never writes `<alignment>` elements
```scala
val cellXfsElem = elem("cellXfs", "count" -> index.cellStyles.size.toString)(
  index.cellStyles.map { style =>
    elem(
      "xf",
      "borderId" -> borderIdx.toString,
      ...
    )()  // ❌ NO CHILDREN - alignment never serialized
  }*
)
```

**Problem**:
- Alignment is parsed from XML
- Alignment is stored in `CellStyle.align`
- Alignment is NEVER written back to XML
- All alignment data lost on write → read round-trip

**Impact**:
- High: Complete data loss for alignment properties
- Breaks round-trip fidelity
- Asymmetric read/write behavior

**Fix Complexity**: Moderate (need conditional serialization)

**Fix**:
1. Add helper to serialize alignment:
```scala
private def alignmentToXml(align: Align): Option[Elem] =
  // Only emit if non-default
  if align == Align.default then None
  else
    val attrs = Seq.newBuilder[(String, String)]

    if align.horizontal != HAlign.General then
      attrs += ("horizontal" -> align.horizontal.toString.toLowerCase(Locale.ROOT))

    if align.vertical != VAlign.Bottom then
      attrs += ("vertical" -> align.vertical.toString.toLowerCase(Locale.ROOT))

    if align.wrapText then
      attrs += ("wrapText" -> "1")

    if align.indent != 0 then
      attrs += ("indent" -> align.indent.toString)

    val attrSeq = attrs.result()
    if attrSeq.isEmpty then None
    else Some(elem("alignment", attrSeq*)())
```

2. Update cellXfs generation:
```scala
val cellXfsElem = elem("cellXfs", "count" -> index.cellStyles.size.toString)(
  index.cellStyles.map { style =>
    val alignChild = alignmentToXml(style.align).toSeq

    elem(
      "xf",
      "applyAlignment" -> (if alignChild.nonEmpty then "1" else "0"),
      "borderId" -> borderIdx.toString,
      "fillId" -> fillIdx.toString,
      "fontId" -> fontIdx.toString,
      "numFmtId" -> numFmtId.toString,
      "xfId" -> "0"
    )(alignChild*)  // Add alignment as child
  }*
)
```

**Test Plan**:
- Test: "Alignment serializes to <alignment> in cellXfs"
- Test: "Non-default alignment includes applyAlignment=1"
- Test: "Default alignment omits <alignment> element"
- Test: "Alignment round-trips through write/read cycle"

**AI Contract Requirements**:
```scala
/** Serialize alignment to XML (ECMA-376 Part 1, §18.8.1)
  *
  * REQUIRES: align is valid Align instance
  * ENSURES:
  *   - Returns None if align == Align.default (optimization)
  *   - Returns Some(<alignment .../>) if any property differs from default
  *   - Emits only non-default attributes (horizontal, vertical, wrapText, indent)
  *   - Attribute values match OOXML spec enum names (lowercase)
  * DETERMINISTIC: Yes (pure transformation based on Align equality)
  * ERROR CASES: None (total function)
  */
private def alignmentToXml(align: Align): Option[Elem] = ...
```

---

## Quality Issues (Medium Priority)

### Issue 5: Scala Version Mismatch ⚠️ Build Inconsistency

**Locations**:
- `build.mill:6` → `val scala3Version = "3.7.3"`
- `project.scala:1` → `//> using scala 3.7.4`

**Problem**:
- Mill build (CI/production) uses 3.7.3
- scala-cli (developer tool) uses 3.7.4
- Potential behavioral differences between dev and CI

**Impact**:
- Medium: May cause subtle compatibility issues
- Confusion for contributors

**Fix Complexity**: Trivial (single character change)

**Fix**:
```scala
// project.scala
//> using scala 3.7.3  // Match CI version
```

**Test Plan**: None (build configuration)

**Additional Actions**:
- Verify all docs mention 3.7.3 consistently
- Consider adding CI check to enforce version parity

---

### Issue 6: xml:space Construction ⚠️ Code Quality

**Locations**:
- `xl-ooxml/src/com/tjclp/xl/ooxml/SharedStrings.scala:41`
- `xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala:55`
- `xl-cats-effect/src/com/tjclp/xl/io/StreamingXmlWriter.scala:153`

**Current State**:

**SharedStrings.scala**:
```scala
new UnprefixedAttribute("xml:space", "preserve", Null)  // ⚠️ String-based
```

**Worksheet.scala**:
```scala
elem("t", "xml:space" -> "preserve")(Text(run.text))  // ⚠️ String-based (via XmlUtil)
```

**StreamingXmlWriter.scala**:
```scala
Attr(QName("xml:space"), List(XmlString("preserve", false)))  // ⚠️ String-based
```

**Problem**:
- All use string `"xml:space"` rather than proper namespace handling
- Works in practice but not technically correct per XML namespaces spec
- Inconsistent with other namespace usage (e.g., `PrefixedAttribute("r", "id", ...)`)

**Impact**:
- Low: Output is correct but approach is non-standard
- Potential issues with strict XML parsers
- Code inconsistency

**Fix Complexity**: Simple (change attribute construction)

**Fix**:

**Scala XML**:
```scala
// SharedStrings.scala, Worksheet.scala
PrefixedAttribute("xml", "space", "preserve", Null)
```

**fs2-data-xml**:
```scala
// StreamingXmlWriter.scala
Attr(QName(Some("xml"), "space"), List(XmlString("preserve", false)))
```

**Test Plan**:
- Existing tests should pass (functional behavior unchanged)
- Consider adding XML validation test against OOXML schema (future)

**AI Contract Requirements**:
```scala
/** Create xml:space="preserve" attribute using proper namespace
  *
  * REQUIRES: None
  * ENSURES:
  *   - Returns attribute with xml prefix properly bound to XML namespace
  *   - Value is "preserve" (OOXML spec requirement)
  *   - Compatible with strict XML parsers
  * DETERMINISTIC: Yes (constant attribute)
  * ERROR CASES: None (compile-time constant)
  */
private def xmlSpacePreserve: Attribute =
  PrefixedAttribute("xml", "space", "preserve", Null)
```

---

## Implementation Order

### Phase 1: Quick Wins (30 minutes)
1. ✅ Fix Scala version mismatch (Issue 5)
2. ✅ Fix xml:space construction (Issue 6)

### Phase 2: OOXML Spec Compliance (2 hours)
3. ✅ Fix default fills to gray125 (Issue 1)
   - Modify defaultFills definition
   - Add 3 tests
   - Add AI contract
4. ✅ Fix SharedStrings count (Issue 2)
   - Update data model
   - Update fromWorkbook
   - Update toXml
   - Add 3 tests
   - Add AI contracts

### Phase 3: Data Loss Prevention (3 hours)
5. ✅ Fix whitespace preservation (Issue 3)
   - Update Worksheet.scala
   - Update StreamingXmlWriter.scala
   - Add 7 tests (both writers + round-trip)
   - Add AI contracts
6. ✅ Fix alignment serialization (Issue 4)
   - Add alignmentToXml helper
   - Update cellXfs generation
   - Add 4 tests
   - Add AI contracts

### Phase 4: Documentation Enhancement (1.5 hours)
7. ✅ Add AI contracts to all modified functions
   - REQUIRES/ENSURES/DETERMINISTIC/ERROR CASES
   - Follow format in `docs/reference/ai-contracts-guide.md`

### Phase 5: Verification (30 minutes)
8. ✅ Run full test suite (708 tests)
9. ✅ Verify formatting
10. ✅ Commit with detailed message

---

## Definition of Done

- [ ] All 6 issues fixed with code changes
- [ ] 10 new tests added (698 → 708 total)
- [ ] All tests passing (708/708)
- [ ] Zero compilation warnings
- [ ] AI contracts added to all affected functions
- [ ] Code formatted (`./mill __.checkFormat`)
- [ ] Changes committed with detailed message referencing this plan
- [ ] `docs/plan/roadmap.md` updated to mark P4.5 as complete
- [ ] Full OOXML spec compliance verified
- [ ] Zero data loss on round-trip for all supported features

---

## Related Documentation

- **Style Review**: `docs/reviews/style-review.md` - Original review identifying these issues
- **AI Contracts Guide**: `docs/reference/ai-contracts-guide.md` - Documentation standard
- **OOXML Research**: `docs/reference/ooxml-research.md` - Spec details
- **Testing Guide**: `docs/reference/testing-guide.md` - Test coverage standards
- **Roadmap**: `docs/plan/roadmap.md` - Phase tracking

---

## Success Metrics

**Before P4.5**:
- ❌ 4 spec violations (gray125, SharedStrings, whitespace, alignment)
- ❌ 2 data loss scenarios (whitespace, alignment)
- ❌ 2 quality issues (Scala version, xml:space)
- 698 tests passing

**After P4.5**:
- ✅ Zero spec violations
- ✅ Zero data loss scenarios
- ✅ Full code quality and consistency
- ✅ Comprehensive AI contracts
- 708 tests passing
