# Full-Fidelity Surgical Modification Specification

**Status**: In Progress
**Priority**: P0 (Critical)
**Estimated Effort**: 8-12 hours
**Started**: 2025-11-17

## Overview

This specification details the implementation of **full-fidelity surgical modification** for XL's XLSX writer. The goal is to enable modification of cell values while preserving ALL Excel metadata with byte-perfect fidelity for unchanged content, eliminating Excel corruption warnings.

## Problem Statement

### Current Behavior

The existing surgical modification implementation (Phases 1-4) successfully:
- ✅ Preserves unknown ZIP parts (charts, images, comments) byte-for-byte
- ✅ Copies unmodified sheets verbatim
- ✅ Preserves ContentTypes.xml declarations
- ✅ Preserves Relationships (theme, calcChain links)

However, when modifying a cell in a real-world Excel file, **Excel displays a corruption warning**:
> "We found a problem with some content in 'syndigo-surgical-output.xlsx'. Do you want us to try to recover as much as we can?"

### Root Cause Analysis

Detailed investigation revealed two critical issues:

#### Issue 1: workbook.xml Complete Regeneration

**Original file** (77KB, minified to 1 line):
```xml
<workbook xmlns="..." xmlns:r="..." xmlns:mc="..." xmlns:x15="..." xmlns:xr="...">
  <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="16130"/>
  <workbookPr codeName="ThisWorkbook" defaultThemeVersion="166925"/>
  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
    <!-- File path metadata -->
  </mc:AlternateContent>
  <xr:revisionPtr revIDLastSave="0" .../>
  <bookViews>
    <workbookView activeTab="2" firstSheet="2" windowHeight="23040" windowWidth="41760" .../>
  </bookViews>
  <sheets>
    <sheet name="__snloffice" sheetId="872" state="veryHidden" r:id="rId1"/>
    <sheet name="_CIQHiddenCacheSheet" sheetId="871" state="veryHidden" r:id="rId2"/>
    <sheet name="Syndigo - Valuation" sheetId="699" r:id="rId3"/>
    <sheet name="Syndigo - Comps" sheetId="423" r:id="rId4"/>
    <sheet name="Syndigo - Financials & DCF" sheetId="6" r:id="rId5"/>
    <sheet name="Syndigo - WACC" sheetId="7" r:id="rId6"/>
    <sheet name="Not to Print >>>" sheetId="823" state="hidden" r:id="rId7"/>
    <sheet name="Additional Information" sheetId="790" state="hidden" r:id="rId8"/>
    <sheet name="Preliminary Valuation Output" sheetId="800" state="hidden" r:id="rId9"/>
  </sheets>
  <definedNames>
    <!-- THOUSANDS of named ranges for formulas -->
    <definedName name="Company_Name">...</definedName>
    <definedName name="IQ_EBITDA_EST_CUR">...</definedName>
    <!-- ... 500+ more defined names ... -->
  </definedNames>
  <calcPr calcId="162913" fullCalcOnLoad="1"/>
  <extLst>...</extLst>
</workbook>
```

**Regenerated output** (782 bytes, 2 lines):
```xml
<workbook xmlns="..." xmlns:r="...">
  <sheets>
    <sheet name="__snloffice" sheetId="1" r:id="rId1"/>
    <sheet name="_CIQHiddenCacheSheet" sheetId="2" r:id="rId2"/>
    <sheet name="Syndigo - Valuation" sheetId="3" r:id="rId3"/>
    <!-- ... all sheets renumbered 1-9, ALL visible (no state attribute) -->
  </sheets>
</workbook>
```

**What's Lost**:
1. **Sheet visibility attributes** - `veryHidden` system sheets now visible → corruption
2. **Original sheet IDs** - 872, 871, 699 renumbered to 1, 2, 3 → breaks internal refs
3. **Defined names** - All named ranges lost → formulas show `#REF!`
4. **Workbook properties** - codeName, theme version, revision tracking lost
5. **Active tab info** - `activeTab="2"` lost → Excel opens wrong sheet
6. **Calculation settings** - `calcPr` lost → formulas may not recalculate

#### Issue 2: Worksheet Metadata Loss

**Original sheet3.xml** (36KB, 881 cells):
```xml
<worksheet xmlns="..." xmlns:r="..." xmlns:x14="..." mc:Ignorable="..." xr:uid="...">
  <sheetPr>
    <pageSetUpPr fitToPage="1"/>
  </sheetPr>
  <dimension ref="B1:U104"/>
  <sheetViews>
    <sheetView showGridLines="0" tabSelected="1" view="pageBreakPreview" zoomScale="85" ...>
      <selection activeCell="B1" sqref="B1"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultColWidth="8.85" defaultRowHeight="12.75" outlineLevelRow="1" .../>
  <cols>
    <col min="1" max="1" width="10.71" style="7" customWidth="1"/>
    <col min="2" max="2" width="3.71" style="7" customWidth="1"/>
    <!-- ... 6 more col definitions -->
  </cols>
  <sheetData>
    <row r="1" spans="2:16" s="7" customFormat="1" ht="12.95" customHeight="1" .../>
    <row r="2" spans="2:16" s="7" customFormat="1" ht="24.95" customHeight="1" ...>
      <!-- cells -->
    </row>
    <!-- ... 102 more rows -->
  </sheetData>
  <mergeCells count="12">...</mergeCells>
  <pageMargins left="0.7" right="0.7" .../>
  <pageSetup paperSize="1" orientation="portrait" .../>
  <headerFooter>...</headerFooter>
  <drawing r:id="rId1"/>
</worksheet>
```

**Regenerated output** (24KB):
```xml
<worksheet xmlns="..." xmlns:r="...">
  <sheetData>
    <row r="2">
      <c t="s" s="221" r="B2"><v>47</v></c>
      <!-- cells only, no row attributes -->
    </row>
    <!-- ... -->
  </sheetData>
  <mergeCells count="12">...</mergeCells>
</worksheet>
```

**What's Lost**:
1. **Column widths** - All 8 custom column definitions → default width
2. **View settings** - Grid lines hidden, page break view, zoom, selection
3. **Row formatting** - Heights, styles, custom format flags
4. **Page setup** - Margins, orientation, paper size
5. **Sheet properties** - Fit to page, outline levels
6. **Dimension** - Used range reference (B1:U104)

### Why Excel Flags Corruption

Excel's corruption detection triggers because:

1. **Hidden sheets are now visible** - Security/integrity violation
2. **Formula references break** - Defined names missing → `#REF!` errors
3. **Internal references inconsistent** - Sheet IDs changed (872→1) but calcChain still references 872
4. **Missing required metadata** - Excel expects certain elements (dimension, view settings)
5. **Inconsistent structure** - Column widths defined but `<cols>` missing

## Architecture & Design

### Core Principle: Selective Preservation

The key insight is to **preserve all unchanged structure** and only regenerate the minimal changed content:

```
┌─────────────────────────────────────────────────────────────┐
│                       workbook.xml                          │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ PRESERVED:                                           │  │
│  │ - fileVersion, workbookPr, bookViews                 │  │
│  │ - definedNames (ALL named ranges)                    │  │
│  │ - calcPr, extLst                                     │  │
│  └──────────────────────────────────────────────────────┘  │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ UPDATED:                                             │  │
│  │ - <sheets> element (names, order, visibility)        │  │
│  └──────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│                    sheet3.xml (modified)                    │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ PRESERVED:                                           │  │
│  │ - sheetPr, dimension, sheetViews                     │  │
│  │ - sheetFormatPr, cols (column widths)                │  │
│  │ - pageMargins, pageSetup, headerFooter               │  │
│  │ - drawing, legacyDrawing references                  │  │
│  └──────────────────────────────────────────────────────┘  │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ REGENERATED:                                         │  │
│  │ - <sheetData> (cells with modified values)           │  │
│  │ - <mergeCells> (if changed)                          │  │
│  └──────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│              sheet1.xml, sheet2.xml, etc.                   │
│                  (unmodified sheets)                        │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ COPIED VERBATIM: byte-for-byte from source ZIP       │  │
│  └──────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
```

### Data Model Extensions

#### 1. OoxmlWorkbook (xl-ooxml/src/com/tjclp/xl/ooxml/Workbook.scala)

```scala
case class OoxmlWorkbook(
  sheets: Seq[SheetRef],
  // Preserve unparsed workbook elements
  fileVersion: Option[Elem] = None,
  workbookPr: Option[Elem] = None,
  alternateContent: Option[Elem] = None,  // mc:AlternateContent
  revisionPtr: Option[Elem] = None,       // xr:revisionPtr
  bookViews: Option[Elem] = None,
  definedNames: Option[Elem] = None,      // CRITICAL: thousands of named ranges
  calcPr: Option[Elem] = None,
  extLst: Option[Elem] = None,
  otherElements: Seq[Elem] = Seq.empty    // Any other top-level elements
) derives CanEqual
```

**Key Change**: Instead of only storing `sheets`, we preserve the entire workbook structure as parsed XML elements.

#### 2. SheetRef (xl-ooxml/src/com/tjclp/xl/ooxml/Workbook.scala)

```scala
case class SheetRef(
  name: SheetName,
  sheetId: Int,           // Preserve original ID (872, not renumbered to 1)
  relationshipId: String,
  state: Option[String] = None  // "hidden" | "veryHidden" | None (visible)
) derives CanEqual
```

**Key Change**: Add `state` field to preserve sheet visibility.

#### 3. OoxmlWorksheet (xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala)

```scala
case class OoxmlWorksheet(
  rows: Seq[OoxmlRow],
  mergedRanges: Set[CellRange] = Set.empty,

  // Preserve worksheet metadata (emitted in OOXML schema order)
  sheetPr: Option[Elem] = None,           // Sheet properties
  dimension: Option[Elem] = None,         // Used range (B1:U104)
  sheetViews: Option[Elem] = None,        // View settings (zoom, gridLines, selection)
  sheetFormatPr: Option[Elem] = None,     // Default row/col sizes
  cols: Option[Elem] = None,              // Column definitions (CRITICAL for widths)

  // Page layout (after sheetData)
  pageMargins: Option[Elem] = None,
  pageSetup: Option[Elem] = None,
  headerFooter: Option[Elem] = None,

  // Drawings and objects (after page layout)
  drawing: Option[Elem] = None,           // Charts reference
  legacyDrawing: Option[Elem] = None,     // VML drawings (comments)
  picture: Option[Elem] = None,
  oleObjects: Option[Elem] = None,
  controls: Option[Elem] = None,

  // Extensions
  extLst: Option[Elem] = None,
  otherElements: Seq[Elem] = Seq.empty
) extends XmlWritable
```

**Key Change**: Store all worksheet elements as parsed XML, not just rows + merges.

#### 4. OoxmlRow (Optional Enhancement)

```scala
case class OoxmlRow(
  rowIndex: Int,
  cells: Seq[OoxmlCell],

  // Preserve row-level attributes
  spans: Option[String] = None,        // "2:16" (optimization hint)
  style: Option[Int] = None,           // s="7" (row-level style)
  height: Option[Double] = None,       // ht="12.95" (custom height)
  customHeight: Boolean = false,       // customHeight="1"
  customFormat: Boolean = false,       // customFormat="1"
  hidden: Boolean = false,             // hidden="1"
  outlineLevel: Option[Int] = None,    // outlineLevel="1"
  collapsed: Boolean = false,          // collapsed="1"
  thickBot: Boolean = false,           // thickBot="1"
  thickTop: Boolean = false,           // thickTop="1"
  dyDescent: Option[Double] = None,    // x14ac:dyDescent="0.2"
  otherAttrs: MetaData = Null          // Catch-all for unknown attributes
)
```

**Note**: Row-level attributes are lower priority. The main fidelity issues are at workbook and worksheet levels.

## Implementation Phases

### Phase 1: Workbook.xml Preservation (2-3 hours)

**Priority**: P0 (Blocking for all other phases)

#### 1.1 Parse Full Workbook Structure

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/Workbook.scala`

**Changes**:

1. Update `OoxmlWorkbook.fromXml()` to parse ALL elements:
   ```scala
   def fromXml(elem: Elem): Either[String, OoxmlWorkbook] =
     for
       // Parse sheets
       sheetsElem <- getChild(elem, "sheets")
       sheetElems = getChildren(sheetsElem, "sheet")
       sheets <- parseSheets(sheetElems)

       // Extract preserved elements (order doesn't matter, we'll re-emit in order)
       fileVersion = (elem \ "fileVersion").headOption.collect { case e: Elem => e }
       workbookPr = (elem \ "workbookPr").headOption.collect { case e: Elem => e }
       alternateContent = (elem \ "AlternateContent").headOption.collect { case e: Elem => e }
       revisionPtr = (elem \ "revisionPtr").headOption.collect { case e: Elem => e }
       bookViews = (elem \ "bookViews").headOption.collect { case e: Elem => e }
       definedNames = (elem \ "definedNames").headOption.collect { case e: Elem => e }
       calcPr = (elem \ "calcPr").headOption.collect { case e: Elem => e }
       extLst = (elem \ "extLst").headOption.collect { case e: Elem => e }

       // Collect any other top-level elements we don't explicitly handle
       otherElements = elem.child.collect {
         case e: Elem if !knownElements.contains(e.label) => e
       }

     yield OoxmlWorkbook(
       sheets, fileVersion, workbookPr, alternateContent, revisionPtr,
       bookViews, definedNames, calcPr, extLst, otherElements.toSeq
     )
   ```

2. Update `parseSheets()` to extract `state` attribute:
   ```scala
   private def parseSheets(elems: Seq[Elem]): Either[String, Seq[SheetRef]] =
     elems.map { e =>
       for
         name <- getAttr(e, "name").flatMap(SheetName.parse)
         sheetId <- getAttr(e, "sheetId").flatMap(_.toIntOption.toRight("Invalid sheetId"))
         rId <- getAttr(e, "r:id")
         state = getAttrOpt(e, "state")  // Extract visibility
       yield SheetRef(name, sheetId, rId, state)
     }.sequence
   ```

#### 1.2 Serialize with Preserved Structure

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/Workbook.scala`

**Changes**:

Update `OoxmlWorkbook.toXml()` to emit full structure in OOXML schema order:

```scala
def toXml: Elem =
  val children = Seq.newBuilder[Node]

  // Emit elements in OOXML Part 1 schema order
  fileVersion.foreach(children += _)
  workbookPr.foreach(children += _)
  alternateContent.foreach(children += _)
  revisionPtr.foreach(children += _)
  bookViews.foreach(children += _)

  // Sheets element (regenerated with current names/order/visibility)
  val sheetElems = sheets.map { ref =>
    val baseAttrs = Seq(
      "name" -> ref.name.value,
      "sheetId" -> ref.sheetId.toString,
      "r:id" -> ref.relationshipId
    )
    val stateAttr = ref.state.map(s => Seq("state" -> s)).getOrElse(Seq.empty)
    elem("sheet", (baseAttrs ++ stateAttr)*: _*)()
  }
  children += elem("sheets")(sheetElems*)

  // Defined names (CRITICAL - preserve all named ranges)
  definedNames.foreach(children += _)

  calcPr.foreach(children += _)
  extLst.foreach(children += _)
  otherElements.foreach(children += _)

  Elem(
    null, "workbook",
    new UnprefixedAttribute("xmlns", nsSpreadsheetML,
      new PrefixedAttribute("xmlns", "r", nsRelationships, Null)),
    TopScope, minimizeEmpty = false,
    children.result()*
  )
```

**Critical**: The `definedNames` element contains thousands of named ranges that formulas depend on. Preserving this unchanged is essential.

#### 1.3 Update parsePreservedStructure

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala`

**Changes**:

1. Add fourth return value to parse workbook.xml:
   ```scala
   private def parsePreservedStructure(
     sourcePath: Path
   ): (Option[ContentTypes], Option[Relationships], Option[Relationships], Option[OoxmlWorkbook]) =
     var contentTypes: Option[ContentTypes] = None
     var rootRels: Option[Relationships] = None
     var workbookRels: Option[Relationships] = None
     var workbook: Option[OoxmlWorkbook] = None  // NEW

     val sourceZip = new ZipInputStream(new FileInputStream(sourcePath.toFile))
     try
       var entry = sourceZip.getNextEntry
       while entry != null do
         val entryName = entry.getName

         // ... existing ContentTypes, Relationships parsing ...

         // NEW: Parse workbook.xml
         else if entryName == "xl/workbook.xml" then
           val content = new String(sourceZip.readAllBytes(), "UTF-8")
           parseXml(content, "xl/workbook.xml") match
             case Right(elem) =>
               OoxmlWorkbook.fromXml(elem).foreach(wb => workbook = Some(wb))
             case Left(_) => () // Silently ignore - will use minimal fallback

         // ... continue loop ...
     finally sourceZip.close()

     (contentTypes, rootRels, workbookRels, workbook)
   ```

#### 1.4 Update hybridWrite to Use Preserved Workbook

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala`

**Changes**:

```scala
private def hybridWrite(
  workbook: Workbook,
  ctx: SourceContext,
  target: OutputTarget,
  config: WriterConfig
): Unit =
  // ... existing code ...

  // Preserve structural parts from source
  val (preservedContentTypes, preservedRootRels, preservedWorkbookRels, preservedWorkbook) =
    parsePreservedStructure(ctx.sourcePath)

  // Use preserved workbook structure if available
  val ooxmlWb = preservedWorkbook match
    case Some(preserved) =>
      // Update sheets in preserved structure (names may have changed)
      preserved.updateSheets(workbook.sheets)
    case None =>
      // Fallback to minimal (for programmatically created workbooks)
      OoxmlWorkbook.fromDomain(workbook)

  // ... rest of hybridWrite ...
```

**NEW METHOD**: Add `updateSheets()` to OoxmlWorkbook:

```scala
// In OoxmlWorkbook case class
def updateSheets(newSheets: Vector[Sheet]): OoxmlWorkbook =
  // Map domain sheets to SheetRefs, preserving original sheetIds and visibility
  val updatedRefs = newSheets.zipWithIndex.map { case (sheet, idx) =>
    // Try to find original SheetRef to preserve sheetId and visibility
    sheets.find(_.name == sheet.name) match
      case Some(original) =>
        // Preserve original ID and visibility
        original.copy(relationshipId = s"rId${idx + 1}")
      case None =>
        // New sheet - generate new ID (use max existing ID + 1)
        val newId = sheets.map(_.sheetId).maxOption.getOrElse(0) + 1
        SheetRef(sheet.name, newId, s"rId${idx + 1}", None)
  }
  copy(sheets = updatedRefs)
```

### Phase 2: Worksheet Metadata Preservation (4-6 hours)

**Priority**: P0 (Core fidelity requirement)

#### 2.1 Parse Worksheet Metadata

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala`

**Changes**:

Update `OoxmlWorksheet.fromXmlWithSST()` to parse and store ALL elements:

```scala
def fromXmlWithSST(elem: Elem, sst: Option[SharedStrings]): Either[String, OoxmlWorksheet] =
  for
    // Parse sheetData (existing code)
    sheetDataElem <- getChild(elem, "sheetData")
    rowElems = getChildren(sheetDataElem, "row")
    rows <- parseRows(rowElems, sst)
    mergedRanges <- parseMergeCells(elem)

    // Extract ALL preserved metadata
    sheetPr = (elem \ "sheetPr").headOption.collect { case e: Elem => e }
    dimension = (elem \ "dimension").headOption.collect { case e: Elem => e }
    sheetViews = (elem \ "sheetViews").headOption.collect { case e: Elem => e }
    sheetFormatPr = (elem \ "sheetFormatPr").headOption.collect { case e: Elem => e }
    cols = (elem \ "cols").headOption.collect { case e: Elem => e }  // CRITICAL for column widths

    pageMargins = (elem \ "pageMargins").headOption.collect { case e: Elem => e }
    pageSetup = (elem \ "pageSetup").headOption.collect { case e: Elem => e }
    headerFooter = (elem \ "headerFooter").headOption.collect { case e: Elem => e }

    drawing = (elem \ "drawing").headOption.collect { case e: Elem => e }
    legacyDrawing = (elem \ "legacyDrawing").headOption.collect { case e: Elem => e }
    picture = (elem \ "picture").headOption.collect { case e: Elem => e }
    oleObjects = (elem \ "oleObjects").headOption.collect { case e: Elem => e }
    controls = (elem \ "controls").headOption.collect { case e: Elem => e }

    extLst = (elem \ "extLst").headOption.collect { case e: Elem => e }

    // Collect other elements
    otherElements = elem.child.collect {
      case e: Elem if !knownWorksheetElements.contains(e.label) => e
    }

  yield OoxmlWorksheet(
    rows, mergedRanges,
    sheetPr, dimension, sheetViews, sheetFormatPr, cols,
    pageMargins, pageSetup, headerFooter,
    drawing, legacyDrawing, picture, oleObjects, controls,
    extLst, otherElements.toSeq
  )
```

#### 2.2 Serialize with Preserved Metadata

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala`

**Changes**:

Update `OoxmlWorksheet.toXml()` to emit in OOXML schema order:

```scala
def toXml: Elem =
  val children = Seq.newBuilder[Node]

  // OOXML Part 1 schema order (critical for Excel compatibility)
  sheetPr.foreach(children += _)
  dimension.foreach(children += _)
  sheetViews.foreach(children += _)
  sheetFormatPr.foreach(children += _)
  cols.foreach(children += _)  // Column definitions BEFORE sheetData

  // sheetData (always regenerated)
  val sheetDataElem = elem("sheetData")(
    rows.sortBy(_.rowIndex).map(_.toXml)*
  )
  children += sheetDataElem

  // mergeCells (regenerated if present)
  if mergedRanges.nonEmpty then
    val mergeCellElems = mergedRanges.toSeq
      .sortBy(r => (r.start.row.index0, r.start.col.index0))
      .map(range => elem("mergeCell", "ref" -> range.toA1)())
    children += elem("mergeCells", "count" -> mergedRanges.size.toString)(mergeCellElems*)

  // Page layout (after sheetData/mergeCells)
  pageMargins.foreach(children += _)
  pageSetup.foreach(children += _)
  headerFooter.foreach(children += _)

  // Drawings and objects
  drawing.foreach(children += _)
  legacyDrawing.foreach(children += _)
  picture.foreach(children += _)
  oleObjects.foreach(children += _)
  controls.foreach(children += _)

  // Extensions
  extLst.foreach(children += _)
  otherElements.foreach(children += _)

  Elem(
    null, "worksheet",
    new UnprefixedAttribute("xmlns", nsSpreadsheetML,
      new PrefixedAttribute("xmlns", "r", nsRelationships, Null)),
    TopScope, minimizeEmpty = false,
    children.result()*
  )
```

**Note**: The order of elements in OOXML is strictly defined by the schema. Excel may fail validation if elements are out of order.

### Phase 3: Row-Level Attributes (Optional - 2-3 hours)

**Priority**: P1 (Polish, not critical for core functionality)

#### 3.1 Parse Row Attributes

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala`

Update `parseRows()` to extract all row attributes:

```scala
private def parseRows(
  elems: Seq[Elem],
  sst: Option[SharedStrings]
): Either[String, Seq[OoxmlRow]] =
  val parsed = elems.map { e =>
    for
      rStr <- getAttr(e, "r")
      rowIdx <- rStr.toIntOption.toRight(s"Invalid row index: $rStr")

      // Extract all row attributes
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

      // Parse cells
      cellElems = getChildren(e, "c")
      cells <- parseCells(cellElems, sst)

    yield OoxmlRow(
      rowIdx, cells,
      spans, style, height, customHeight, customFormat,
      hidden, outlineLevel, collapsed, thickBot, thickTop, dyDescent,
      Null  // otherAttrs - TODO: capture unknown attributes
    )
  }

  // ... error handling ...
```

#### 3.2 Emit Row Attributes

**File**: `xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala`

Update `OoxmlRow.toXml()` to include all preserved attributes:

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

  val attrs = baseAttrs ++ optionalAttrs.result()

  elem("row", attrs*)(
    cells.sortBy(_.ref.col.index0).map(_.toXml)*
  )
```

## Testing Strategy

### Unit Tests

**New Test File**: `xl-ooxml/test/src/com/tjclp/xl/ooxml/XlsxWriterPreservationSpec.scala`

```scala
class XlsxWriterPreservationSpec extends ScalaCheckSuite:

  test("workbook.xml round-trip preserves defined names"):
    // Parse workbook.xml with defined names
    // Serialize back to XML
    // Verify defined names element is identical

  test("workbook.xml preserves sheet visibility (veryHidden, hidden)"):
    // Parse workbook with hidden sheets
    // Update sheet names
    // Verify visibility attributes preserved

  test("workbook.xml preserves original sheet IDs"):
    // Parse workbook with IDs 872, 871, 699
    // Verify IDs not renumbered to 1, 2, 3

  test("worksheet metadata round-trip preserves cols"):
    // Parse sheet3.xml with custom column widths
    // Modify a cell
    // Serialize back
    // Verify <cols> element identical

  test("worksheet metadata preserves view settings"):
    // Parse sheet with gridLines="0", zoom, selection
    // Verify sheetViews element preserved

  test("row attributes preserved (spans, height, style)"):
    // Parse row with custom height, style
    // Verify attributes preserved in toXml
```

### Integration Tests

**Test with Real-World File**: Syndigo Valuation file (343KB, 9 sheets)

```scala
test("Syndigo file surgical modification - no corruption"):
  // 1. Read Syndigo file
  val input = Paths.get("data/Syndigo Valuation_Q3 2025_2025.10.15_VALUES.xlsx")
  val result = XlsxReader.read(input)

  // 2. Modify cell Z100 in "Syndigo - Valuation"
  val modified = for
    wb <- result
    sheet <- wb("Syndigo - Valuation")
    updated <- sheet.put(ref"Z100" -> "Test")
    final <- wb.put(updated)
  yield final

  // 3. Write with surgical modification
  val output = Files.createTempFile("syndigo-test", ".xlsx")
  XlsxWriter.write(modified.toOption.get, output)

  // 4. Re-read and verify
  val reloaded = XlsxReader.read(output).toOption.get

  // 5. Assertions
  assert(reloaded.sheets.size == 9, "Sheet count preserved")

  // Verify hidden sheets remain hidden
  val firstSheetIsHidden = // TODO: Check sheet visibility in XML
  assert(firstSheetIsHidden, "Sheet 0 (__snloffice) still veryHidden")

  // Verify defined names preserved
  // Extract workbook.xml and check for definedNames element

  // Verify column widths preserved
  // Extract sheet3.xml and check for <cols> element

  // MANUAL TEST: Open in Excel - should not show corruption warning
```

### Manual Testing Protocol

1. **Run Demo Script**:
   ```bash
   scala-cli run data/surgical-demo.sc
   ```

2. **Open Output in Excel**:
   - Open `data/syndigo-surgical-output.xlsx` in Excel
   - **Expected**: No corruption warning
   - **Verify**: Hidden sheets remain hidden (not visible in tabs)
   - **Verify**: Column widths match original
   - **Verify**: Formulas work (no `#REF!` errors)

3. **Compare with Original**:
   - Open original file side-by-side
   - Verify page layout identical (zoom, grid lines, margins)
   - Verify defined names exist (Formulas → Name Manager)
   - Verify charts and images intact

4. **Oracle Tool (Optional)**:
   ```bash
   # Install oracle (Go-based OOXML diff tool)
   brew install steipete/tap/oracle

   # Compare input vs output
   oracle diff "data/Syndigo Valuation_Q3 2025_2025.10.15_VALUES.xlsx" \
               "data/syndigo-surgical-output.xlsx"
   ```

## Success Criteria

### Must Have (P0)

- ✅ No Excel corruption warnings on open
- ✅ Hidden sheets (`veryHidden`, `hidden`) remain hidden
- ✅ Defined names preserved (formulas work, no `#REF!`)
- ✅ Column widths preserved (visual layout intact)
- ✅ View settings preserved (zoom, grid lines, selection)
- ✅ Page setup preserved (margins, orientation, fit to page)
- ✅ Charts, images, comments intact (already working)
- ✅ Modified cell values correctly updated

### Should Have (P1)

- ✅ Row heights preserved
- ✅ Original sheet IDs preserved (not renumbered)
- ✅ Active tab preserved (correct sheet opens)
- ✅ Outline levels preserved (grouping)

### Nice to Have (P2)

- ✅ Cell-level preservation (only changed cells regenerated, not entire row)
- ✅ Performance: <100ms for single cell modification
- ✅ Memory: O(1) for unmodified sheets (no parsing)

## Risk Assessment

### High Risk

1. **OOXML Schema Compliance**
   - **Risk**: Element order matters; incorrect order may cause Excel to reject file
   - **Mitigation**: Follow schema order exactly (sheetPr → dimension → sheetViews → ... → extLst)
   - **Test**: Validate with real-world files

2. **Namespace Handling**
   - **Risk**: Elements like `mc:AlternateContent`, `xr:revisionPtr` have special namespaces
   - **Mitigation**: Preserve xmlns declarations from original file
   - **Test**: Round-trip files with multiple namespaces

3. **Defined Names with Sheet References**
   - **Risk**: Named ranges may reference sheets by ID; renumbering breaks references
   - **Mitigation**: Preserve original sheet IDs (don't renumber)
   - **Test**: File with 500+ defined names (Syndigo)

### Medium Risk

1. **Partial Updates to Complex Structures**
   - **Risk**: Updating `<sheets>` while preserving `<definedNames>` may create inconsistencies
   - **Mitigation**: Don't allow sheet deletions/reordering in surgical mode (document limitation)
   - **Test**: Try to delete a sheet with defined names referencing it

2. **Column Width Calculations**
   - **Risk**: Column widths in `<cols>` may conflict with cell widths
   - **Mitigation**: Preserve `<cols>` element unchanged for unmodified sheets
   - **Test**: Sheets with merged cells and custom column widths

### Low Risk

1. **Row Attribute Edge Cases**
   - **Risk**: Some row attributes may interact (e.g., `customHeight` without `ht`)
   - **Mitigation**: Preserve attributes exactly as parsed
   - **Test**: Property-based tests with generated row attributes

## Future Enhancements

### Phase 4: Cell-Level Preservation

Instead of regenerating entire `<row>` elements, only regenerate changed cells:

```xml
<row r="100">
  <!-- Preserve unchanged cells as raw XML -->
  <c r="A100" s="1" t="s"><v>0</v></c>
  <c r="B100" s="1" t="s"><v>1</v></c>

  <!-- Regenerate only Z100 (modified) -->
  <c r="Z100" t="inlineStr"><is><t>🎯 Modified</t></is></c>
</row>
```

**Benefit**: Ultimate fidelity - only changed cells touched
**Effort**: 4-6 hours additional
**Priority**: P2 (optimization, not required)

### Phase 5: Smart Merge for Sheet Structure Changes

Support add/remove/reorder sheets by intelligently updating defined names:

```scala
def updateDefinedNamesForSheetRename(
  definedNames: Elem,
  oldName: String,
  newName: String
): Elem =
  // Parse each defined name formula
  // Update references from old name to new name
  // Re-serialize
```

**Benefit**: Full surgical modification flexibility
**Effort**: 8-12 hours (complex formula parsing)
**Priority**: P3 (future, not MVP)

## Timeline

| Phase | Description | Effort | Status |
|-------|-------------|--------|--------|
| **Phase 1** | Workbook.xml preservation | 2-3 hours | Not Started |
| **Phase 2** | Worksheet metadata preservation | 4-6 hours | Not Started |
| **Phase 3** | Row attributes (optional) | 2-3 hours | Not Started |
| **Testing** | Integration + manual tests | 1-2 hours | Not Started |
| **Total** | | **8-12 hours** | |

**Estimated Completion**: 1-2 working days

## References

- **ECMA-376 Part 1**: Office Open XML File Formats - SpreadsheetML
  - Section 18.2: Workbook Part
  - Section 18.3: Worksheet Part
- **Microsoft Excel OOXML Structure**: https://docs.microsoft.com/en-us/openspecs/office_standards/
- **Syndigo Test File**: `/Users/rcaputo3/git/xl/data/Syndigo Valuation_Q3 2025_2025.10.15_VALUES.xlsx`
- **Oracle Tool**: https://github.com/steipete/oracle (OOXML diff/validation)

## Appendix: Known OOXML Elements

### Workbook Elements (in schema order)

1. `<fileVersion>` - App version info
2. `<fileSharing>` - Password protection
3. `<workbookPr>` - Workbook properties (codeName, defaultThemeVersion)
4. `<mc:AlternateContent>` - File path metadata
5. `<xr:revisionPtr>` - Revision tracking
6. `<workbookProtection>` - Workbook-level protection
7. `<bookViews>` - Window size, active tab, first visible sheet
8. `<sheets>` - Sheet definitions (name, ID, visibility, relationship)
9. `<functionGroups>` - Custom function groups
10. `<externalReferences>` - Links to external workbooks
11. `<definedNames>` - Named ranges (CRITICAL)
12. `<calcPr>` - Calculation settings
13. `<oleSize>` - OLE object size
14. `<customWorkbookViews>` - Custom views
15. `<pivotCaches>` - Pivot table caches
16. `<smartTagPr>` - Smart tag properties
17. `<smartTagTypes>` - Smart tag types
18. `<webPublishing>` - Web publishing settings
19. `<fileRecoveryPr>` - File recovery properties
20. `<webPublishObjects>` - Web publish objects
21. `<extLst>` - Extensions

### Worksheet Elements (in schema order)

1. `<sheetPr>` - Sheet properties
2. `<dimension>` - Used range reference
3. `<sheetViews>` - View settings (zoom, grid, selection)
4. `<sheetFormatPr>` - Default row/column sizes
5. `<cols>` - Column definitions (CRITICAL for widths)
6. `<sheetData>` - Cell data
7. `<sheetCalcPr>` - Sheet-level calc settings
8. `<sheetProtection>` - Sheet protection
9. `<protectedRanges>` - Protected ranges
10. `<scenarios>` - What-if scenarios
11. `<autoFilter>` - AutoFilter settings
12. `<sortState>` - Sort state
13. `<dataConsolidate>` - Data consolidation
14. `<customSheetViews>` - Custom views
15. `<mergeCells>` - Merged cell ranges
16. `<phoneticPr>` - Phonetic properties
17. `<conditionalFormatting>` - Conditional formatting
18. `<dataValidations>` - Data validation rules
19. `<hyperlinks>` - Hyperlinks
20. `<printOptions>` - Print options
21. `<pageMargins>` - Page margins
22. `<pageSetup>` - Page setup (orientation, paper size)
23. `<headerFooter>` - Header/footer
24. `<rowBreaks>` - Page break locations (rows)
25. `<colBreaks>` - Page break locations (cols)
26. `<customProperties>` - Custom properties
27. `<cellWatches>` - Cell watches
28. `<ignoredErrors>` - Ignored errors
29. `<smartTags>` - Smart tags
30. `<drawing>` - Drawing reference (charts)
31. `<legacyDrawing>` - Legacy drawing (VML comments)
32. `<legacyDrawingHF>` - Legacy drawing header/footer
33. `<picture>` - Background picture
34. `<oleObjects>` - OLE objects
35. `<controls>` - ActiveX controls
36. `<webPublishItems>` - Web publish items
37. `<tableParts>` - Table parts
38. `<extLst>` - Extensions

**Note**: The order above is mandated by the OOXML schema. Excel may reject files with elements out of order.
