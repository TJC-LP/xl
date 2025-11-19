# Surgical Modification Corruption Fixes

**Status**: In Progress (Phase A Complete)
**Priority**: P0 (Critical)
**Created**: 2025-11-17
**Updated**: 2025-11-17 (after workbook namespace fix)
**Parent**: surgical-modification-full-fidelity.md

## Current Status

### ✅ Completed (Phase A)
- **Workbook.xml namespace handling** - FIXED
  - `mc:Ignorable` attribute preserved
  - Namespace pollution eliminated (10 xmlns vs 50+)
  - 2 passing tests in `WorkbookNamespaceSpec.scala`
  - All 585 tests passing

### 🔄 In Progress
- Worksheet namespace pollution (81 xmlns in sheet3.xml)
- Missing worksheet elements (conditionalFormatting, rowBreaks, etc.)
- Row attribute preservation

### 📋 Remaining
- SST verification (Agent 5 reported corruption - needs manual check)
- ZIP packaging fixes
- Whitespace preservation

## Problem Statement

Despite implementing full workbook and worksheet metadata preservation, Excel still shows a corruption warning when opening surgically modified files:

> "We found a problem with some content in 'syndigo-surgical-output.xlsx'. Do you want us to try to recover as much as we can?"

However, the file is **significantly improved**:
- ✅ "Syndigo - Valuation" (modified sheet) looks mostly good
- ✅ 850 defined names preserved
- ✅ Hidden sheets remain hidden
- ✅ All 36 unknown parts byte-perfect
- ⚠️ Minor corruption warning remains

## Research Findings

Six parallel agents investigated the corruption. Key findings:

### Agent Results Summary

| Agent | Focus | Status | Critical Findings |
|-------|-------|--------|-------------------|
| Agent 1 | Workbook.xml | ✅ Complete | **P0**: Missing `mc:Ignorable`, namespace pollution |
| Agent 2 | Modified sheet3.xml | ✅ Complete | **P0**: Missing row attrs, missing elements, namespace pollution |
| Agent 3 | Unmodified sheets | ✅ Complete | ✅ **ALL byte-identical** (8/8 perfect) |
| Agent 4 | "Syndigo - Comps" | ✅ Complete | ✅ **Byte-identical** (false alarm) |
| Agent 5 | SST & Styles | ✅ Complete | 🔴 **SST corruption reported** (needs verification) |
| Agent 6 | ZIP structure | ✅ Complete | P1: ZIP version mismatch, entry order |

### Unmodified Sheet Preservation: **100% Success** ✅

**Critical validation**: All 8 unmodified sheets are byte-for-byte identical to the source. This confirms:
- `copyPreservedPart()` is working perfectly
- Surgical modification strategy is sound
- Issues are isolated to **regenerated files only**

---

## Root Cause Analysis

### Issue 1: Namespace Handling (P0 - CRITICAL)

**Problem**: Scala XML library is not properly handling namespace scope inheritance.

#### 1.1 Missing `mc:Ignorable` Attribute

**What's Wrong**:
```xml
<!-- Input (correct) -->
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
          xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
          xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
          xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6"
          xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"
          xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
          mc:Ignorable="x15 xr xr6 xr10 xr2">
  <!-- content -->
</workbook>

<!-- Output (broken) -->
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <!-- content -->
</workbook>
```

**Why It Matters**:
- `mc:Ignorable` tells Excel: "ignore these Microsoft extension namespaces if you don't understand them"
- Without it, Excel treats unknown namespaces (x15, xr, etc.) as **errors**
- **This is the PRIMARY trigger for the corruption warning**

**Impact**:
- Excel flags file as corrupted on open
- User must click "Yes" to repair
- Excel strips unknown namespace content during repair

#### 1.2 Namespace Pollution

**What's Wrong**:
Every child element re-declares all 8 namespaces:

```xml
<!-- Input (correct) - namespaces on root only -->
<workbook xmlns="..." xmlns:r="..." xmlns:mc="..." xmlns:x15="..." xmlns:xr="..." xmlns:xr6="..." xmlns:xr10="..." xmlns:xr2="...">
  <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="16130"/>
  <workbookPr codeName="ThisWorkbook" defaultThemeVersion="124226"/>
  <definedNames>
    <definedName name="Company_Name">Sheet1!$A$1</definedName>
  </definedNames>
</workbook>

<!-- Output (polluted) - namespaces repeated on EVERY child -->
<workbook xmlns="..." xmlns:r="...">
  <fileVersion xmlns="..." xmlns:xr2="..." xmlns:xr10="..." xmlns:xr6="..." xmlns:xr="..." xmlns:x15="..." xmlns:mc="..." xmlns:r="..." appName="xl" lastEdited="7" lowestEdited="7" rupBuild="16130"/>
  <workbookPr xmlns="..." xmlns:xr2="..." xmlns:xr10="..." xmlns:xr6="..." xmlns:xr="..." xmlns:x15="..." xmlns:mc="..." xmlns:r="..." codeName="ThisWorkbook" defaultThemeVersion="124226"/>
  <definedNames xmlns="..." xmlns:xr2="..." xmlns:xr10="..." xmlns:xr6="..." xmlns:xr="..." xmlns:x15="..." xmlns:mc="..." xmlns:r="...">
    <definedName xmlns="..." xmlns:xr2="..." xmlns:xr10="..." xmlns:xr6="..." xmlns:xr="..." xmlns:x15="..." xmlns:mc="..." xmlns:r="..." name="Company_Name">Sheet1!$A$1</definedName>
  </definedNames>
</workbook>
```

**Why It Happens**:
When we do `elem.child.collectFirst { case e: Elem if e.label == "definedNames" => e }`, Scala's XML library preserves the element WITH its namespace context. When we later emit it, those namespaces get serialized.

**Impact**:
- 12KB of redundant namespace declarations (15% bloat)
- Violates OOXML spec (namespaces should inherit from root)
- Confuses Excel's parser

**Affected Files**:
- `xl/workbook.xml`: 8 child elements polluted
- `xl/worksheets/sheet3.xml`: 50+ elements polluted

---

### Issue 2: Missing Worksheet Elements (P0 - DATA LOSS)

**Agent 2 found 4 critical missing elements** in regenerated sheet3.xml:

#### 2.1 Conditional Formatting (12 instances)

**What's Missing**:
```xml
<conditionalFormatting sqref="H2:H4">
  <cfRule type="expression" dxfId="0" priority="1">
    <formula>$H2&lt;0</formula>
  </cfRule>
</conditionalFormatting>
```

**Impact**:
- All red/green cell highlighting based on formulas lost
- Financial models rely heavily on conditional formatting
- User-visible data loss

#### 2.2 Print Options

**What's Missing**:
```xml
<printOptions headings="1"/>
```

**Impact**: Print settings lost (minor)

#### 2.3 Row Breaks

**What's Missing**:
```xml
<rowBreaks count="1" manualBreakCount="1">
  <brk id="93" max="16383" man="1"/>
</rowBreaks>
```

**Impact**: Manual page break at row 93 lost

#### 2.4 Custom Properties

**What's Missing**:
```xml
<customProperties>
  <customPr name="{00000000-0000-0000-0000-000000000000}"/>
</customProperties>
```

**Impact**: Worksheet GUID lost (may break references)

**Root Cause**: Our `knownElements` set doesn't include these elements, so they're filtered out instead of preserved.

---

### Issue 3: Row-Level Metadata Loss (P0 - FORMATTING LOSS)

**Agent 2 found ALL rows lost critical attributes**:

**What's Missing**:
```xml
<!-- Input (correct) -->
<row r="2" spans="2:16" s="7" customFormat="1" ht="24.95" customHeight="1">
  <c r="B2" s="221" t="s"><v>47</v></c>
</row>

<!-- Output (broken) -->
<row r="2">
  <c r="B2" s="221" t="s"><v>47</v></c>
</row>
```

**Lost Attributes**:
- `spans` - Cell coverage hint (optimization)
- `s` - Row-level style ID (formatting)
- `ht` - Custom row height (layout)
- `customHeight`, `customFormat` - Flags
- `thickBot`, `thickTop` - Border styles
- `x14ac:dyDescent` - Font descent adjustment

**Impact**:
- Row heights incorrect → visual layout broken
- Row-level formatting lost
- Excel repair may restore defaults

**Root Cause**: `OoxmlRow` doesn't store these attributes (we planned this as "Phase 3" but haven't implemented yet).

---

### Issue 4: ZIP Packaging Format (P1 - EXCEL PARANOIA)

**Agent 6 found packaging inconsistencies**:

#### 4.1 ZIP Version Mismatch
- Input: All entries use **ZIP version 4.5** (create_version=45)
- Output: Mixed **ZIP 2.0** (text) and **ZIP 1.0** (images)

**Impact**: Excel expects consistent modern ZIP versioning

#### 4.2 Flag Bits Differ
- Input: 0x0006 (UTF-8 + data descriptor)
- Output: 0x0808 (different encoding flags)

**Impact**: Excel may interpret encodings differently

#### 4.3 Entry Order Changed
- Input: Worksheets → theme → styles → SST
- Output: Styles → SST → worksheets → theme

**Impact**: Unclear if Excel cares, but deviation from norm

---

### Issue 5: Whitespace Normalization (P2 - MINOR)

**Agent 1 found 2 defined names with trimmed whitespace**:
- `Spaces`: `"         "` → `" "` (9 spaces → 1 space)
- `_GSRATES_1`: `"CT300001Latest          "` → `"CT300001Latest "`

**Impact**: Formulas like `LEN(Spaces)` will return wrong values

---

## Implementation Plan

### Fix 1: Namespace Handling (P0 - 2-3 hours) - ✅ WORKBOOK COMPLETE

**Status**: Workbook-level fix implemented and tested. Worksheet-level fix pending.

**Objective**: Preserve root element attributes and prevent namespace pollution.

#### Step 1.1: Preserve Root Attributes - ✅ IMPLEMENTED

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/Workbook.scala`

**Implementation Complete**: Commit 2304cec + local changes

**What Was Done**:
```scala
def toXml: Elem =
  val children = Seq.newBuilder[Node]
  // ... build children ...

  val nsBindings = NamespaceBinding(null, nsSpreadsheetML,
                     NamespaceBinding("r", nsRelationships, TopScope))

  Elem(null, "workbook", Null, nsBindings, minimizeEmpty = false, children.result()*)
```

**Problem**:
- Only declares `xmlns` and `xmlns:r`
- Loses `mc:Ignorable` and other xmlns declarations
- `Null` for attributes loses all original attributes

**Implemented Solution** (simpler than originally spec'd):
```scala
case class OoxmlWorkbook(
  sheets: Seq[SheetRef],
  // ... existing fields ...

  // NEW: Preserve root element metadata
  rootAttributes: MetaData = Null,
  rootScope: NamespaceBinding = defaultWorkbookScope
)

def toXml: Elem =
  val children = Seq.newBuilder[Node]
  // ... build children ...

  // Use preserved scope and attributes
  val scope = Option(rootScope).getOrElse(defaultWorkbookScope)
  val attrs = Option(rootAttributes).getOrElse(Null)

  Elem(null, "workbook", attrs, scope, minimizeEmpty = false, children.result()*)
```

**Changes Made**:
1. ✅ Added `rootAttributes: MetaData` to preserve `mc:Ignorable`
2. ✅ Added `rootScope: NamespaceBinding` to preserve xmlns declarations
3. ✅ Captured during `fromXml()`: `elem.attributes`, `elem.scope`
4. ✅ Used in `toXml()` for root element
5. ✅ Added `defaultWorkbookScope` constant for fallback

**Test Coverage**:
- ✅ `WorkbookNamespaceSpec.scala` (2 tests passing)
- ✅ Verifies `mc:Ignorable` preserved
- ✅ Verifies namespace pollution prevented (checks xmlns:mc count = 1)

#### Step 1.2: Strip Namespace Pollution - ✅ WORKBOOK DONE, WORKSHEET PENDING

**Status**: Workbook child elements are clean. Worksheet elements still polluted.

**Problem**: When we preserve elements like `<definedNames>`, Scala XML keeps their namespace context, causing duplication.

**Current Results**:
- ✅ Workbook.xml: 10 xmlns declarations total (clean)
- ❌ Worksheet sheet3.xml: 81 xmlns declarations (polluted)

**Solution**: Create a utility to strip redundant namespaces:

```scala
/** Strip redundant namespace declarations from element (they'll inherit from parent). */
private def stripRedundantNamespaces(elem: Elem): Elem =
  // Only keep namespaces not already declared on parent (workbook root)
  val cleanedChildren = elem.child.map {
    case e: Elem => stripRedundantNamespaces(e)
    case other => other
  }

  // Create new element with TopScope (no local namespaces)
  elem.copy(
    scope = TopScope,
    child = cleanedChildren
  )
```

**Apply to all preserved elements**:
```scala
fileVersion.foreach(children += stripRedundantNamespaces(_))
workbookPr.foreach(children += stripRedundantNamespaces(_))
definedNames.foreach(children += stripRedundantNamespaces(_))
// ... etc ...
```

**Alternative (Safer)**: Use XML rewriting with proper scope control:

```scala
import scala.xml.transform.{RewriteRule, RuleTransformer}

private def cleanNamespaces(elem: Elem): Elem =
  val rule = new RewriteRule {
    override def transform(n: Node): Seq[Node] = n match
      case e: Elem =>
        // Remove local namespace bindings (will inherit from parent)
        val cleanScope = TopScope
        Elem(e.prefix, e.label, e.attributes, cleanScope, e.minimizeEmpty, transform(e.child)*)
      case other => other
  }

  new RuleTransformer(rule).transform(elem).head.asInstanceOf[Elem]
```

#### Step 1.3: Preserve mc:Ignorable Exactly

**Implementation**:

```scala
// In OoxmlWorkbook.fromXml():
def fromXml(elem: Elem): Either[String, OoxmlWorkbook] =
  for
    // ... existing parsing ...
  yield OoxmlWorkbook(
    sheets,
    fileVersion,
    workbookPr,
    alternateContent,
    revisionPtr,
    bookViews,
    definedNames,
    calcPr,
    extLst,
    otherElements.toSeq,
    sourceRootElem = Some(elem)  // NEW: Store full root element
  )
```

**Testing**:
```scala
test("workbook.xml preserves mc:Ignorable attribute"):
  val input = parseXml("""
    <workbook xmlns="..." mc:Ignorable="x15 xr">
      <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
    </workbook>
  """)

  val wb = OoxmlWorkbook.fromXml(input).toOption.get
  val output = wb.toXml

  // Verify mc:Ignorable is preserved
  val ignorable = output.attribute("http://schemas.openxmlformats.org/markup-compatibility/2006", "Ignorable")
  assert(ignorable.isDefined)
  assertEquals(ignorable.get.text, "x15 xr")
```

**Estimated Effort**: 2-3 hours

---

### Fix 2: Missing Worksheet Elements (P0 - 1-2 hours) - 📋 NOT STARTED

**Objective**: Preserve conditionalFormatting, printOptions, rowBreaks, customProperties.

#### Step 2.1: Update knownElements Set

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala` (line 331-371)

**Current**:
```scala
knownElements = Set(
  "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols",
  "sheetData", "mergeCells",
  "pageMargins", "pageSetup", "headerFooter",
  "drawing", "legacyDrawing", "picture", "oleObjects", "controls",
  "extLst",
  // Additional elements from OOXML spec
  "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios",
  "autoFilter", "sortState", "dataConsolidate", "customSheetViews",
  "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks",
  "printOptions", "rowBreaks", "colBreaks", "customProperties",
  "cellWatches", "ignoredErrors", "smartTags",
  "legacyDrawingHF", "webPublishItems", "tableParts"
)
```

**Problem**: These ARE in the set, but we're not explicitly extracting them!

**Fix**: Add explicit extraction for critical elements:

```scala
def fromXmlWithSST(elem: Elem, sst: Option[SharedStrings]): Either[String, OoxmlWorksheet] =
  for
    // ... existing parsing ...

    // Add missing critical elements
    conditionalFormatting = elem.child.collect {
      case e: Elem if e.label == "conditionalFormatting" => e
    }
    printOptions = (elem \ "printOptions").headOption.collect { case e: Elem => e }
    rowBreaks = (elem \ "rowBreaks").headOption.collect { case e: Elem => e }
    colBreaks = (elem \ "colBreaks").headOption.collect { case e: Elem => e }
    customPropertiesWs = (elem \ "customProperties").headOption.collect { case e: Elem => e }

    // Update otherElements to exclude these
    knownElements = Set(
      "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
      "mergeCells", "conditionalFormatting", "printOptions", "rowBreaks", "colBreaks",
      "customProperties", "pageMargins", "pageSetup", "headerFooter",
      "drawing", "legacyDrawing", "picture", "oleObjects", "controls", "extLst"
    )
    otherElements = elem.child.collect {
      case e: Elem if !knownElements.contains(e.label) => e
    }

  yield OoxmlWorksheet(
    rows, mergedRanges,
    sheetPr, dimension, sheetViews, sheetFormatPr, cols,
    conditionalFormatting.toSeq,  // NEW: Can have multiple instances
    printOptions, rowBreaks, colBreaks, customPropertiesWs,
    pageMargins, pageSetup, headerFooter,
    drawing, legacyDrawing, picture, oleObjects, controls,
    extLst, otherElements.toSeq
  )
```

#### Step 2.2: Update OoxmlWorksheet Data Model

**Add fields**:
```scala
case class OoxmlWorksheet(
  rows: Seq[OoxmlRow],
  mergedRanges: Set[CellRange] = Set.empty,
  sheetPr: Option[Elem] = None,
  dimension: Option[Elem] = None,
  sheetViews: Option[Elem] = None,
  sheetFormatPr: Option[Elem] = None,
  cols: Option[Elem] = None,

  // NEW: Critical missing elements
  conditionalFormatting: Seq[Elem] = Seq.empty,  // Can have multiple
  printOptions: Option[Elem] = None,
  rowBreaks: Option[Elem] = None,
  colBreaks: Option[Elem] = None,
  customPropertiesWs: Option[Elem] = None,

  // Existing fields
  pageMargins: Option[Elem] = None,
  // ... rest ...
)
```

#### Step 2.3: Update toXml Emission Order

**Per OOXML §18.3.1.99 schema order**:

```scala
def toXml: Elem =
  val children = Seq.newBuilder[Node]

  sheetPr.foreach(children += cleanNamespaces(_))
  dimension.foreach(children += cleanNamespaces(_))
  sheetViews.foreach(children += cleanNamespaces(_))
  sheetFormatPr.foreach(children += cleanNamespaces(_))
  cols.foreach(children += cleanNamespaces(_))

  // sheetData (regenerated)
  children += elem("sheetData")(rows.sortBy(_.rowIndex).map(_.toXml)*)

  // Merge cells
  if mergedRanges.nonEmpty then /* ... */

  // NEW: Conditional formatting (BEFORE page layout per schema)
  conditionalFormatting.foreach(cf => children += cleanNamespaces(cf))

  // Data validations, hyperlinks, etc. would go here

  printOptions.foreach(children += cleanNamespaces(_))
  pageMargins.foreach(children += cleanNamespaces(_))
  pageSetup.foreach(children += cleanNamespaces(_))
  headerFooter.foreach(children += cleanNamespaces(_))

  rowBreaks.foreach(children += cleanNamespaces(_))
  colBreaks.foreach(children += cleanNamespaces(_))

  customPropertiesWs.foreach(children += cleanNamespaces(_))

  // ... rest of elements ...
```

**CRITICAL**: Element order must match OOXML schema exactly or Excel may reject the file.

**Estimated Effort**: 1-2 hours

---

### Fix 3: Row-Level Attributes (P1 - 3-4 hours) - 📋 NOT STARTED

**Objective**: Preserve row heights, styles, and formatting attributes.

This is the "Phase 3" we documented but didn't implement. From the spec (surgical-modification-full-fidelity.md):

#### Step 3.1: Extend OoxmlRow

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala`

**Add fields**:
```scala
case class OoxmlRow(
  rowIndex: Int,
  cells: Seq[OoxmlCell],

  // Preserve row-level attributes
  spans: Option[String] = None,        // "2:16"
  style: Option[Int] = None,           // s="7"
  height: Option[Double] = None,       // ht="12.95"
  customHeight: Boolean = false,       // customHeight="1"
  customFormat: Boolean = false,       // customFormat="1"
  hidden: Boolean = false,             // hidden="1"
  outlineLevel: Option[Int] = None,    // outlineLevel="1"
  collapsed: Boolean = false,          // collapsed="1"
  thickBot: Boolean = false,           // thickBot="1"
  thickTop: Boolean = false,           // thickTop="1"
  dyDescent: Option[Double] = None     // x14ac:dyDescent="0.2"
)
```

#### Step 3.2: Parse Row Attributes

**Update parseRows()**:

```scala
private def parseRows(
  elems: Seq[Elem],
  sst: Option[SharedStrings]
): Either[String, Seq[OoxmlRow]] =
  val parsed = elems.map { e =>
    for
      rStr <- getAttr(e, "r")
      rowIdx <- rStr.toIntOption.toRight(s"Invalid row index: $rStr")

      // Extract ALL row attributes
      spans = getAttrOpt(e, "spans")
      style = getAttrOpt(e, "s").flatMap(_.toIntOption)
      height = getAttrOpt(e, "ht").flatMap(_.toDoubleOption)
      customHeight = getAttrOpt(e, "customHeight").contains("1")
      customFormat = getAttrOpt(e, "customFormat").contains("1")
      hidden = getAttrOpt(e, "hidden").contains("1")
      outlineLevel = getAttrOpt(e, "outlineLevel").flatMap(_.toIntOption)
      collapsed = getAttrOpt(e, "collapsed").contains("1")
      thickBot = getAttrOpt(e, "thickBot").contains("1")
      thickTop = getAttrOpt(e, "thickTop").contains("1")
      dyDescent = getAttrOpt(e, "x14ac:dyDescent").flatMap(_.toDoubleOption)

      cellElems = getChildren(e, "c")
      cells <- parseCells(cellElems, sst)

    yield OoxmlRow(
      rowIdx, cells,
      spans, style, height, customHeight, customFormat,
      hidden, outlineLevel, collapsed, thickBot, thickTop, dyDescent
    )
  }

  // ... error handling ...
```

#### Step 3.3: Emit Row Attributes

**Update OoxmlRow.toXml()**:

```scala
def toXml: Elem =
  val baseAttrs = Seq("r" -> rowIndex.toString)

  val optionalAttrs = Seq.newBuilder[(String, String)]
  spans.foreach(s => optionalAttrs += ("spans" -> s))
  style.foreach(s => optionalAttrs += ("s" -> s.toString))
  height.foreach(h => optionalAttrs += ("ht" -> h.toString))
  if customHeight then optionalAttrs += ("customHeight" -> "1")
  if customFormat then optionalAttrs += ("customFormat" -> "1")
  if hidden then optionalAttrs += ("hidden" -> "1")
  outlineLevel.foreach(l => optionalAttrs += ("outlineLevel" -> l.toString))
  if collapsed then optionalAttrs += ("collapsed" -> "1")
  if thickBot then optionalAttrs += ("thickBot" -> "1")
  if thickTop then optionalAttrs += ("thickTop" -> "1")
  dyDescent.foreach(d => optionalAttrs += ("x14ac:dyDescent" -> d.toString))

  val allAttrs = baseAttrs ++ optionalAttrs.result()

  elem("row", allAttrs*)(
    cells.sortBy(_.ref.col.index0).map(_.toXml)*
  )
```

**Estimated Effort**: 3-4 hours (parsing + serialization + testing)

---

### Fix 4: SharedStrings Investigation (P0 - URGENT - 30 min) - ⚠️ NEEDS VERIFICATION

**Agent 5 reported catastrophic SST corruption**, but user says file "looks mostly quite good."

**Action Required**: Manual verification before implementing any fix.

#### Verification Steps:

1. **Extract and inspect sharedStrings.xml**:
   ```bash
   unzip -p data/syndigo-surgical-output.xlsx xl/sharedStrings.xml | xmllint --format - | head -50
   ```

2. **Check first 5 entries**:
   - Are they readable text strings?
   - Or binary garbage as agent reported?

3. **Check if cells display correctly**:
   - Open file in Excel
   - Check if cell A1 in any sheet shows text or garbage
   - If text is readable → SST is fine (agent false alarm)
   - If garbage → SST is broken (agent correct)

**Possible Outcomes**:

**A. SST is fine (False Alarm)**:
- Agent misinterpreted XML encoding
- UTF-8 characters look like garbage in ASCII
- No action needed

**B. SST is corrupted**:
- Critical bug in `SharedStrings.fromWorkbook()` or `SharedStrings.toXml()`
- Need to debug SST writer
- Estimated fix: 2-4 hours

**Estimated Effort**: 30 min investigation + 0-4 hours fix (depending on outcome)

---

### Fix 5: ZIP Packaging Format (P1 - 2-3 hours) - 📋 NOT STARTED

**Objective**: Match Excel's expected ZIP packaging format exactly.

#### Step 5.1: ZIP Version Consistency

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala` (writePart method)

**Current Issue**: No explicit version control

**Solution**:
```scala
private def writePart(
  zip: ZipOutputStream,
  path: String,
  content: Elem,
  config: WriterConfig
): Unit =
  val entry = new ZipEntry(path)
  entry.setTime(0L)

  // NEW: Set ZIP version to 4.5 (45 in decimal)
  entry.setMethod(ZipEntry.DEFLATED)
  // Note: ZipEntry doesn't expose setVersion() in Java API
  // May need to use Apache Commons Compress for finer control

  zip.putNextEntry(entry)
  // ... rest of method ...
```

**Alternative**: Use Apache Commons Compress library for full ZIP control:
```scala
libraryDeps += "org.apache.commons:commons-compress:1.27.1"
```

#### Step 5.2: Preserve Entry Order

**Current**: Entries written in code order (structural files → sheets → preserved parts)

**Solution**: Write in original ZIP order

**Implementation**:
1. During read, store entry order in `PartManifest`:
   ```scala
   case class PartManifestEntry(
     path: String,
     size: Option[Long],
     compressedSize: Option[Long],
     method: Option[Int],
     sheetIndex: Option[Int],
     parsed: Boolean,
     entryOrder: Int  // NEW: Position in original ZIP
   )
   ```

2. During write, sort parts by `entryOrder` before emitting:
   ```scala
   val orderedParts = ctx.partManifest.entries.toSeq.sortBy(_._2.entryOrder)
   orderedParts.foreach { case (path, entry) =>
     if shouldRegenerate(path) then
       writePart(zip, path, regeneratedContent, config)
     else
       copyPreservedPart(ctx.sourcePath, path, zip)
   }
   ```

**Estimated Effort**: 2-3 hours

---

### Fix 6: Whitespace Preservation (P2 - 1 hour) - 📋 NOT STARTED

**Objective**: Preserve exact text content in definedNames (don't trim).

**Current Issue**: Scala XML's `elem.text` trims whitespace by default.

**Solution**: Use `xml:space="preserve"` or manual text extraction:

```scala
private def parseDefinedNameValue(elem: Elem): String =
  // Don't use elem.text (trims whitespace)
  elem.child.collect { case Text(s) => s }.mkString
```

**Apply during**: `definedNames` element preservation in fromXml/toXml.

**Estimated Effort**: 1 hour

---

## Implementation Priority

### Phase A: Namespace Fixes (P0 - 2-3 hours)
**Immediate Impact**: Will likely eliminate corruption warning

1. Fix 1.1: Preserve root element attributes (mc:Ignorable)
2. Fix 1.2: Strip namespace pollution from child elements
3. Test: Verify Excel opens without warning

### Phase B: SST Verification (P0 - 30 min - 4 hours)
**Blocking**: Must verify before proceeding

1. Manual inspection of output sharedStrings.xml
2. If corrupted: Debug SST writer
3. If fine: Agent false alarm, no action needed

### Phase C: Missing Elements (P0 - 1-2 hours)
**Data Loss Prevention**

1. Fix 2: Add conditionalFormatting, printOptions, rowBreaks, customProperties
2. Test: Verify elements preserved after regeneration

### Phase D: Row Attributes (P1 - 3-4 hours)
**Formatting Fidelity**

1. Fix 3: Extend OoxmlRow with attribute fields
2. Parse and emit row attributes
3. Test: Verify row heights preserved

### Phase E: ZIP Packaging (P1 - 2-3 hours)
**Polish**

1. Fix 5.1: ZIP version consistency
2. Fix 5.2: Preserve entry order
3. Test: Verify zipinfo output matches original

### Phase F: Whitespace (P2 - 1 hour)
**Edge Case**

1. Fix 6: Preserve exact whitespace in definedNames
2. Test: `LEN(Spaces)` formula returns correct value

---

## Testing Strategy

### Unit Tests

**New test file**: `xl-ooxml/test/src/com/tjclp/xl/ooxml/XlsxWriterCorruptionFixesSpec.scala`

```scala
class XlsxWriterCorruptionFixesSpec extends munit.FunSuite:

  test("workbook.xml preserves mc:Ignorable attribute"):
    // Parse workbook with mc:Ignorable
    // Regenerate
    // Verify attribute present in output

  test("workbook.xml does not pollute child namespaces"):
    // Parse workbook with xmlns on root only
    // Regenerate
    // Count xmlns occurrences in output (should be 1, not 9)

  test("worksheet preserves conditionalFormatting"):
    // Parse sheet with 12 conditional formatting rules
    // Regenerate
    // Verify all 12 rules present in output

  test("worksheet preserves row attributes (spans, height, style)"):
    // Parse sheet with custom row heights
    // Regenerate
    // Verify attributes preserved

  test("worksheet preserves printOptions and rowBreaks"):
    // Parse sheet with print settings and page breaks
    // Regenerate
    // Verify elements present
```

### Integration Tests

**Syndigo File Round-Trip**:

```scala
test("Syndigo file surgical modification - no corruption warning"):
  val input = Paths.get("data/Syndigo Valuation_Q3 2025_2025.10.15_VALUES.xlsx")

  // 1. Read
  val wb = XlsxReader.read(input).toOption.get

  // 2. Modify cell
  val modified = wb("Syndigo - Valuation")
    .flatMap(_.put(ref"Z100" -> "Test"))
    .flatMap(wb.put)
    .toOption.get

  // 3. Write
  val output = Files.createTempFile("syndigo-test", ".xlsx")
  XlsxWriter.write(modified, output)

  // 4. Re-read
  val reloaded = XlsxReader.read(output).toOption.get

  // 5. Assertions
  assert(reloaded.sheets.size == 9)

  // Extract and verify workbook.xml
  val workbookXml = extractFromZip(output, "xl/workbook.xml")
  assert(workbookXml.contains("mc:Ignorable"), "mc:Ignorable must be present")

  // Count namespace declarations (should be ~10 on root, not 50+ total)
  val nsCount = workbookXml.split("xmlns").length - 1
  assert(nsCount < 15, s"Too many xmlns declarations: $nsCount (namespace pollution)")

  // Verify hidden sheets
  assert(workbookXml.contains("veryHidden"), "Hidden sheets must stay hidden")

  // Verify conditional formatting preserved
  val sheet3Xml = extractFromZip(output, "xl/worksheets/sheet3.xml")
  assert(sheet3Xml.contains("conditionalFormatting"), "Conditional formatting lost")

  // MANUAL: Open in Excel - should NOT show corruption warning ✅
```

### Golden File Test

**Byte-for-byte comparison** (after fixes):

```scala
test("Syndigo file surgical modification - golden file"):
  // Known-good output (manually verified in Excel, no warnings)
  val golden = Paths.get("data/golden/syndigo-surgical-golden.xlsx")

  // Regenerate from input
  val input = Paths.get("data/Syndigo Valuation_Q3 2025_2025.10.15_VALUES.xlsx")
  val wb = XlsxReader.read(input).flatMap(wb =>
    wb("Syndigo - Valuation")
      .flatMap(_.put(ref"Z100" -> "Test"))
      .flatMap(wb.put)
  ).toOption.get

  val output = Files.createTempFile("test", ".xlsx")
  XlsxWriter.write(wb, output)

  // Compare key files
  val goldenWorkbook = extractFromZip(golden, "xl/workbook.xml")
  val outputWorkbook = extractFromZip(output, "xl/workbook.xml")

  assertEquals(goldenWorkbook, outputWorkbook, "Workbook XML must match golden file")

  // Sheet3 should match (modulo the Z100 cell we added)
  // Can use XML diff to ignore specific cell
```

---

## Success Criteria

### Must Have (Blocking for Merge)

- ✅ No Excel corruption warning when opening output file
- ✅ Hidden sheets remain hidden (veryHidden preserved)
- ✅ Conditional formatting rules intact (12 in Syndigo file)
- ✅ Row heights preserved (visual layout matches original)
- ✅ All 850 defined names intact (formulas work)
- ✅ Column widths preserved (already working)
- ✅ All tests passing (583 current + new tests)

### Should Have (Quality)

- ✅ Namespace pollution eliminated (<15 xmlns in entire workbook.xml)
- ✅ ZIP version consistent (4.5 for all entries)
- ✅ Entry order preserved (matches original)
- ✅ Print settings and page breaks preserved

### Nice to Have (Polish)

- ✅ Whitespace exact (spaces in defined names preserved)
- ✅ Cell attribute ordering per spec (r, s, t, ...)
- ✅ Byte-for-byte golden file match (except modified cell)

---

## Risk Assessment

### High Risk

**Namespace Stripping**:
- **Risk**: Removing namespaces may break elements that need them (xr:revisionPtr, mc:AlternateContent)
- **Mitigation**: Only strip redundant declarations; preserve necessary prefixed namespaces
- **Test**: Round-trip files with complex namespace usage

**Row Attribute Parsing**:
- **Risk**: Missing attributes (x14ac:dyDescent) may have special namespace handling
- **Mitigation**: Preserve as raw `MetaData` if can't parse individually
- **Test**: Files with extended attributes

### Medium Risk

**Element Ordering**:
- **Risk**: OOXML schema has 36 possible worksheet elements; order must be exact
- **Mitigation**: Follow §18.3.1.99 schema order precisely
- **Test**: Validate with Microsoft's Open XML SDK validator

**SST Corruption** (if real):
- **Risk**: Fundamental serialization bug may require major refactoring
- **Mitigation**: Manual verification first; may need to preserve SST verbatim
- **Test**: Files with RichText, special characters, formulas

### Low Risk

**ZIP Packaging**:
- **Risk**: Java's ZipOutputStream may not support all options (version control)
- **Mitigation**: Use Apache Commons Compress if needed
- **Test**: Compare zipinfo output

---

## Timeline

| Phase | Description | Effort | Cumulative |
|-------|-------------|--------|------------|
| **Phase A** | Namespace fixes | 2-3 hours | 3 hours |
| **Phase B** | SST verification + fix | 0.5-4 hours | 4-7 hours |
| **Phase C** | Missing worksheet elements | 1-2 hours | 6-9 hours |
| **Phase D** | Row attributes | 3-4 hours | 10-13 hours |
| **Phase E** | ZIP packaging | 2-3 hours | 13-16 hours |
| **Phase F** | Whitespace | 1 hour | 14-17 hours |
| **Testing** | Integration + manual | 2-3 hours | **16-20 hours** |

**Realistic Estimate**: 2-3 working days for complete implementation

**Fast Path** (Phases A+B+C only): 4-9 hours (1 working day)

---

## Alternative Approaches

### Option 1: Minimal Namespace Fix (Phase A Only)

**Effort**: 2-3 hours
**Impact**: May eliminate corruption warning entirely
**Risk**: Low
**Recommendation**: **Try this first** - if it fixes Excel warning, defer other phases

### Option 2: Copy workbook.xml Verbatim

**Effort**: 15 minutes
**Impact**: Guaranteed to preserve all workbook metadata perfectly
**Risk**: Can't modify sheet names, can't add/remove sheets
**Recommendation**: Emergency fallback if namespace fixes fail

### Option 3: Preserve Entire Modified Sheet as XML

Instead of parsing → domain → regenerating, preserve the original sheet XML and patch only the changed cell(s):

**Effort**: 4-6 hours (XML manipulation library)
**Impact**: Perfect fidelity, minimal changes
**Risk**: Medium (complex XML patching logic)
**Recommendation**: Long-term goal (Phase 6?)

---

## Progress Report (2025-11-17)

### What's Been Implemented

**Commit**: Local changes (not yet committed)
**Files Modified**:
- `xl-ooxml/src/com/tjclp/xl/ooxml/Workbook.scala` (+10 lines)
- `xl-ooxml/test/src/com/tjclp/xl/ooxml/WorkbookNamespaceSpec.scala` (new file, 60 lines)

**Implementation**:
1. ✅ Added `rootAttributes: MetaData` field to `OoxmlWorkbook`
2. ✅ Added `rootScope: NamespaceBinding` field to `OoxmlWorkbook`
3. ✅ Captured `elem.attributes` and `elem.scope` during `fromXml()`
4. ✅ Used preserved attributes/scope in `toXml()`
5. ✅ Created comprehensive tests (2 passing)

**Results**:
- Workbook.xml: `mc:Ignorable` attribute now preserved ✅
- Workbook.xml: Namespace pollution eliminated (10 xmlns vs 50+) ✅
- All 585 tests passing ✅

**Remaining Issues**:
- Worksheet namespace pollution (81 xmlns in sheet3.xml)
- Missing worksheet elements
- Row attributes not preserved
- SST needs verification

### Current Test Results

```bash
$ ./mill xl-ooxml.test.testOnly com.tjclp.xl.ooxml.WorkbookNamespaceSpec
✅ round-trip workbook preserves mc:Ignorable attribute and namespace bindings
✅ generated minimal workbook exposes default spreadsheet and relationship namespaces

$ ./mill __.test
✅ All 585 tests passing
```

**Syndigo File Test**:
- Modified: Cell Z100 in "Syndigo - Valuation"
- Output: 320KB (input: 334KB, delta: -14KB)
- Workbook.xml: 89KB with full metadata
- Sheet3.xml: 32KB with metadata (but namespace pollution remains)

## Immediate Next Steps

### Step 1: Verify SST Status (30 min) - ⚠️ URGENT

**Action**: Manual inspection of sharedStrings.xml

```bash
# Extract and view first 10 SST entries
unzip -p data/syndigo-surgical-output.xlsx xl/sharedStrings.xml | xmllint --format - | grep -A2 "<si>" | head -40

# Check cell A1 content in sheet 0
unzip -p data/syndigo-surgical-output.xlsx xl/worksheets/sheet1.xml | xmllint --format - | grep "A1" -A2
```

**Decision Point**:
- If SST is fine → Proceed with Phase A (namespace fixes)
- If SST is broken → Fix SST first, then Phase A

### Step 2: Implement Phase A (Namespace Fixes)

**Focus**: Fix the PRIMARY cause of corruption warning

1. Preserve `mc:Ignorable` attribute
2. Strip namespace pollution from child elements
3. Test with Syndigo file
4. Verify Excel opens without warning

**If successful**: Stop here, document limitations, ship it ✅
**If not successful**: Continue with Phases B-F

### Step 3: Incremental Testing

After **each fix**, test with Syndigo file:
- Open in Excel
- Check for corruption warning
- Verify data looks correct
- Document improvements

---

## Success Metrics

### After Phase A (Namespace Fixes)

**Target**:
- Excel opens without corruption warning
- All data readable
- Hidden sheets hidden
- Formulas work

**Acceptance Criteria**:
- Manual Excel test: No repair prompt
- File size: ~325KB (similar to current)
- All 583 tests passing

### After All Phases

**Target**:
- Byte-for-byte fidelity except for modified cells
- Perfect visual match in Excel
- No warnings, no repair needed
- All formatting intact

**Acceptance Criteria**:
- Excel validation: No warnings ✅
- Visual test: Side-by-side with original (identical layout) ✅
- Formula test: All calculations work ✅
- Round-trip test: Read → Write → Read (data preserved) ✅
- Golden file test: Matches known-good output ✅

---

## References

- **Agent Reports**: Full details from parallel investigation
  - Agent 1: Workbook namespace analysis
  - Agent 2: Sheet3 deep diff (23KB data loss identified)
  - Agent 3: Unmodified sheet verification (100% success)
  - Agent 4: Sheet4 false alarm investigation
  - Agent 5: SST corruption warning (needs verification)
  - Agent 6: ZIP packaging analysis

- **ECMA-376 Part 1**:
  - §18.2: Workbook Part
  - §18.3.1.99: Worksheet schema ordering
  - §22.4: Markup Compatibility (mc:Ignorable)

- **Related Specs**:
  - `surgical-modification-full-fidelity.md`: Parent implementation spec
  - `surgical-modification.md`: Original phases 1-4

---

## Appendix: Agent Findings Summary

### Agent 1: Workbook.xml Forensics
- ❌ Missing `mc:Ignorable` (PRIMARY CAUSE)
- ❌ Namespace pollution (8 elements × 8 namespaces = 64 redundant declarations)
- ⚠️ 2 defined names trimmed (minor)
- ✅ All 850 defined names preserved

### Agent 2: Modified Sheet Deep Diff
- ❌ Namespace pollution (50+ elements)
- ❌ Missing Row 1 (empty formatted row)
- ❌ ALL row attributes lost (spans, s, ht, customHeight, etc.)
- ❌ Missing: conditionalFormatting (12), printOptions, rowBreaks, customProperties
- ✅ Z100 cell added successfully
- ✅ Column definitions preserved

### Agent 3: Unmodified Sheets Verification
- ✅ **8/8 sheets byte-identical** (100% preservation)
- ✅ **5/5 worksheet .rels byte-identical**
- ✅ `copyPreservedPart()` working perfectly

### Agent 4: "Syndigo - Comps" Investigation
- ✅ Sheet4 byte-identical (false alarm)
- ✅ All dependent files preserved
- ℹ️ Corruption report likely from other sheets

### Agent 5: SST & Styles Comparison
- 🔴 **SST corruption reported** (entry 0 binary garbage) - **NEEDS VERIFICATION**
- ⚠️ Styles 43% smaller (deduplicated)
- ⚠️ May be false alarm based on user feedback

### Agent 6: ZIP Structural Integrity
- ⚠️ ZIP version mismatch (4.5 → 2.0/1.0)
- ⚠️ Flag bits differ (0x0006 → 0x0808)
- ⚠️ Entry order changed
- ✅ All 51 entries present
- ✅ Content types complete
- ✅ Relationships correct

---

**Document**: `/Users/rcaputo3/git/xl/docs/plan/surgical-modification-corruption-fixes.md`
