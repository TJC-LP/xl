# OOXML Research & Quick Reference

Comprehensive research on OOXML (Office Open XML) SpreadsheetML specification with concrete API sketches and implementation guidance.

---

## Quick Reference: Common OOXML Paths

**Required Parts** (minimum workbook):
- `[Content_Types].xml` - Content type declarations
- `/_rels/.rels` - Package relationships
- `/xl/workbook.xml` - Workbook definition
- `/xl/_rels/workbook.xml.rels` - Workbook relationships
- `/xl/worksheets/sheet1.xml` - Worksheet data

**Optional Parts**:
- `/xl/sharedStrings.xml` - Shared strings table (SST)
- `/xl/styles.xml` - Style definitions
- `/xl/theme/theme1.xml` - Theme colors and fonts
- `/xl/drawings/drawing1.xml` - Images and shapes
- `/xl/charts/chart1.xml` - Chart definitions
- `/xl/comments1.xml` - Cell comments
- `/xl/tables/table1.xml` - Excel tables

**Common Relationship Types** (fragments):
```
http://schemas.openxmlformats.org/officeDocument/2006/relationships/
  - worksheet
  - sharedStrings
  - styles
  - theme
  - drawing
  - chart
  - comments
```

---

# Executive summary

**Goal.** Build a pure, functional Scala 3 library for XLSX/XLSM that’s faster, safer, and nicer to use than POI—without sacrificing interoperability.

**Core calls.**

* Use **OPC** (Open Packaging Conventions) as the ground truth for container layout and relationships; **minimize required parts** and lean on **inline strings** for small/streaming writes, switching to **sharedStrings** when dedup wins. ([Microsoft Learn][1])
* Represent the Excel grid with a **persistent, chunked** model, stream **XML in fs2** and parse via **fs2-data-xml**, falling back to **Aalto (StAX)** for maximum throughput hot paths. Zip with **Commons Compress** for Zip64/non-seekable streams. ([fs2-data.gnieh.org][2])
* Keep **formulas** as an AST + DSL first; don’t ship a full evaluator in v1.0. (Excel caches results and you can force recalculation.) If/when needed, consider a separate evaluation module or interop with external engines. ([poi.apache.org][3])
* **Styles**: centralize dedup in `styles.xml` (fonts/fills/borders/numFmts/cellXfs) and cap growth to avoid Excel’s 64k unique style limit. ([Office Open XML][4])
* Prioritize **streaming APIs** (row and cell cursors) for reads/writes; offer an **immutable builder DSL** on top. For very large writes, copy SXSSF’s approach—window rows to disk. ([poi.apache.org][5])

---

# 1) OOXML (SpreadsheetML) deep-dive

## Executive summary

An `.xlsx` is a **ZIP** with parts and **relationships**. The **minimum workbook** contains `[Content_Types].xml`, `/_rels/.rels`, `/xl/workbook.xml`, `/xl/_rels/workbook.xml.rels`, and one `/xl/worksheets/sheet1.xml`. Everything else (styles, sharedStrings, theme, docProps) is optional. Relationships are looked up via sibling `"_rels/*.rels"` files; types are URIs. ([Wikipedia][6])

## Detailed findings

**OPC structure & relationships.** Packages expose a content-types stream (`[Content_Types].xml`) plus `.rels` files at the root and beside parts; relationships carry (Id, Type URI, Target[, TargetMode]). ([Wikipedia][6])

**Minimum parts for a working workbook.** A minimal SpreadsheetML doc: workbook with at least one `<sheet>` pointing (by relationship `r:id`) to a worksheet part. ([Microsoft Learn][7])

**Shared vs inline strings.**

* `t="s"` means the cell value is an **index** into `xl/sharedStrings.xml` (dedup wins for large text repetition).
* `t="inlineStr"` embeds the text inline—**simpler and often faster** for generated content and streaming. ([Microsoft Learn][8])

**Worksheet dimension.** `<dimension ref="A1:C2"/>` is **optional**; apps often compute used range themselves. Don’t rely on it for correctness. ([Microsoft Learn][9])

**Namespaces.** Core SpreadsheetML uses `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` and `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`. ([schemas.liquid-technologies.com][10])

**Content types.** Crucial overrides include:

* workbook: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml`
* worksheet: `…spreadsheetml.worksheet+xml`
* styles: `…spreadsheetml.styles+xml`
* sharedStrings: `…sharedStrings+xml` (or for XLSB/alt, see Microsoft specs) ([Wikipedia][11])

### Minimal “Hello World” package (inline string)

```
/[Content_Types].xml
/_rels/.rels
/xl/workbook.xml
/xl/_rels/workbook.xml.rels
/xl/worksheets/sheet1.xml
```

`[Content_Types].xml`

```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
```

`/_rels/.rels`

```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
</Relationships>
```

`/xl/workbook.xml`

```xml
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
```

`/xl/_rels/workbook.xml.rels`

```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>
```

`/xl/worksheets/sheet1.xml`

```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Hello, world</t></is></c>
    </row>
  </sheetData>
</worksheet>
```

This exact layout opens in Excel; it shows why inline strings are perfect for small generated sheets. ([Microsoft Learn][7])

## Design recommendations

* Provide `XlsxPackager.minimal(workbookName, sheetName)` that emits exactly the parts above; flag to prefer `inlineStr` for small texts and switch to `sharedStrings` automatically above a dedup threshold. ([ericwhite.com][12])
* Treat `<dimension>` as advisory only.

## Open questions

* When should we introduce `docProps/core.xml` and theme parts by default? (They’re optional but common.)

---

# 2) Excel data model & type system

## Executive summary

Cell values come in a handful of XML encodings (`t` attribute): **number (`n`)**, **shared string (`s`)**, **inline string (`inlineStr`)**, **boolean (`b`)**, **error (`e`)**, **string (`str`, formula text result)**, and sometimes **`d` for ISO date** (rare in practice). Dates are **serial doubles** with formatting; Excel is limited to **~15 significant digits** because it uses IEEE-754 doubles. ([Microsoft Learn][13])

## Detailed findings

* `c` (cell): A1 reference in `r`, style index `s`, dtype `t`; value in `<v>` or `<is>` (inlineStr). ([Microsoft Learn][13])
* **Dates/times** are normally numbers plus a date/time number format; **1900/1904 systems** differ by 1462 days; Excel **wrongly treats 1900 as leap year** for historical Lotus 1-2-3 compatibility. Your API must abstract the epoch and bug. ([Microsoft Support][14])
* **Error values** include `#DIV/0!`, `#N/A`, `#NAME?`, `#NULL!`, `#NUM!`, `#REF!`, `#VALUE!`; dynamic arrays add cases like `#SPILL!` and `#CALC!`. ([Microsoft Support][15])
* **Hyperlinks** live under `<hyperlinks><hyperlink ref="A1" r:id="…"/></hyperlinks>` with a relationship of type `…/hyperlink` (Target may be external or internal location). ([Microsoft Learn][16])

## Type-safe ADT (sketch)

```scala
enum XCellValue:
  case Number(value: Double)                // beware precision; expose BigDecimal helpers
  case Text(value: String, rich: Option[RichRuns] = None) // inline or shared
  case Boolean(value: Boolean)
  case Error(value: XError)                 // DIV0, REF, etc.
  case Formula(ast: Formula, cached: Option[XCellValue])
  case Blank

enum DateSystem: case Excel1900, Excel1904    // plus offset helpers

final case class Cell(
  ref: CellRef, t: Option[CellTypeTag], s: Option[StyleId],
  value: XCellValue
)
```

**Numeric precision.** Keep input/output in `Double`, but offer `BigDecimal` parse/format helpers and warn after **15 significant digits**. ([Microsoft Learn][17])

## Design recommendations

* Distinguish **semantic** vs **display** types: e.g., `Number(45054.0)` with format `yyyy-mm-dd`.
* Add `DateCodec` that round-trips with both date systems and compensates for the 1900-02-29 phantom. ([Microsoft Support][14])

## Open questions

* Do we want a “strict dates” mode that refuses `t="d"` and enforces numeric+format?

---

# 3) Formula system architecture

## Executive summary

Start with **AST + parser + DSL**; defer a full evaluator. Excel stores a cached result and you can mark workbooks **“recalc on open”**. Cover structured references, 3-D refs, and dynamic arrays in the AST. ([poi.apache.org][3])

## Detailed findings

* **Language surface**: functions, operators, A1 references, **structured table references**, **3-D references** (`Sheet1:Sheet4!A1`), and **external links**. ([Microsoft Support][18])
* **Dynamic arrays** spill; failures produce `#SPILL!`. Model single->multi results and spill ranges. ([Microsoft Support][19])
* **Shared formulas**: first cell stores master formula; peers reference it via `f` with `t="shared"` and `si` id; array formulas use `t="array"`. ([Microsoft Learn][20])

### AST & DSL (sketch)

```scala
enum Ref: 
  case A1(cell: CellRef)
  case Range(a: CellRef, b: CellRef)
  case Sheet(name: String, ref: Ref)
  case Sheets3D(first: String, last: String, ref: Ref)
  case Table(name: String, column: Option[String], thisRow: Boolean)

enum Expr:
  case Lit(v: XCellValue)
  case Ref(r: Ref)
  case Unary(op: UOp, e: Expr)
  case Bin(l: Expr, op: BOp, r: Expr)
  case Call(name: Fn, args: List[Expr])

object xl:
  def ref(str: String): Ref = ??? // validated interpolator: xl"A1:B9"
  def sum(xs: Expr*): Expr = Expr.Call(Fn.Sum, xs.toList)
```

**Parser strategy.** Start with a combinator or fast hand-rolled Pratt parser for A1 grammar; table/3-D refs and dynamic array suffixes need special handling.

**Evaluator strategy.** Not in v1: rely on cached results; optionally set `forceRecalcOnOpen=true`. If users need evaluation, provide a separate module and/or bridge to POI’s evaluator or **HyperFormula** (JS, external). ([poi.apache.org][3])

## Open questions

* Scope of built-ins for v1: propose top 50 (SUM, AVERAGE, IF, VLOOKUP/XLOOKUP, INDEX/MATCH, FILTER/SORT/UNIQUE/SEQUENCE, TEXT/DATE/TIME, ROUND/INT, MAX/MIN, CONCAT/TEXTJOIN).

---

# 4) Style & formatting system

## Executive summary

All formatting funnels through `styles.xml` (fonts, fills, borders, numFmts, cellXfs, cellStyleXfs, dxfs). You must deduplicate aggressively to avoid Excel’s **~64k** unique formats ceiling. Colors can be **rgb**, **theme** (+tint), or **indexed**. Custom number formats are limited. ([Office Open XML][4])

## Detailed findings

* `cellXfs` define concrete cell formats, each referencing collections (font/fill/border/numFmt indices); `cellStyleXfs` define named style bases; `dxfs` are for conditional formatting. ([Microsoft Learn][21])
* Excel limits: **Unique cell formats/styles ~65k**; many other caps exist (fonts, fills). ([Microsoft Support][22])
* Colors: theme has 12 core colors (dk1/lt1/accent1..6, hyperlink/followed). `tint` shifts luminance. ([c-rex.net][23])
* Custom number formats: 4 sections (pos/neg/zero/text); users can add only a few hundred customs. ([Microsoft Support][24])

## Design recommendations

* Make styles **structural** and hashable; maintain a **style pool** keyed by full normalized shape `{numFmt,font,fill,border,alignment}` to return an `xfId`.
* Provide **common presets** (currency, percent, date/time).
* Add `ConditionalFormatting` builders that compile to `dxfs` + `cfRule`. ([Microsoft Learn][25])

## Open questions

* Theme awareness API for automatic dark/light color derivation?

---

# 5) Performance & streaming

## Executive summary

Use **fs2** pipelines to stream ZIP entries and XML; prefer **fs2-data-xml** for event streams; for maximum throughput, integrate **Aalto (StAX)**. For ZIP, use **Commons Compress** (+Zip64) to stream non-seekable outputs. For huge writes, window rows to disk like POI **SXSSF**. ([fs2-data.gnieh.org][2])

## Findings & trade-offs

* **DOM vs SAX/StAX**: DOM is simple but explodes memory; event streaming is the right default for large sheets. POI’s guidance mirrors this (XSSF event model; SXSSF for streaming writes). ([poi.apache.org][26])
* **Zip considerations**: some parts may exceed 4 GB; Zip64 and “unknown size” streaming favor Commons Compress’ `ZipArchiveOutputStream`. ([commons.apache.org][27])
* **Parsing libs**: `fs2-data-xml` integrates with fs2 pipes; **Aalto** is a known high-performance StAX implementation. ([fs2-data.gnieh.org][28])

## Design recommendations

* **Reader API**: `WorkbookStream.read(in): Stream[F, Row]` with back-pressure; optional projection to case classes.
* **Writer API**: `WorkbookStream.write(out): Pipe[F, Row, Unit]` with internal row windowing and temp file spill.
* Offer a **RowChunk** policy for batching.

## Open questions

* Where to expose parallelism: per-sheet, per-chunk, or per-zip-entry?

---

# 6) Type class design patterns

## Executive summary

Mirror **doobie/circe**: `CellReader[A]`/`CellWriter[A]` and `RowReader[A]`/`RowWriter[A]` plus **automatic derivation** (Scala 3 `derives`, Magnolia, or Shapeless-3). Provide optics via **Monocle**. Effects via **tagless-final** or concrete `IO`. ([typelevel.org][29])

## API sketch

```scala
trait CellReader[A]:
  def read(cell: Cell): Either[CellError, A]

trait CellWriter[A]:
  def write(a: A): CellPatch // value + style hints

object CellCodec:
  given CellReader[Int] = ???
  given CellWriter[Int] = ???
  // derives for case classes via Magnolia or Scala 3 `derives`
```

Provide `ValidatedNel[CellError, A]` variants for whole-row decoding.

---

# 7) API ergonomics & DX

## Executive summary

Lean, immutable core + **nice DSLs**: validated **cell/range interpolators** (`xl"A1:B9"`), **table** helpers, and **builder** patterns for workbook/sheet/conditional formatting. Take cues from **http4s** and **ScalaTest** for expressive DSL ergonomics. ([http4s.org][30])

## Recommendations

* `xl"A1"` & `xl"A1:B10"` interpolators (compile-time validation if possible).
* Fluent builders with immutable snapshots; helpful error messages and quick fixes.
* Documentation style: small runnable examples, “pit of success” recipes.

---

# 8) Real-world usage patterns

## Executive summary

Typical data-engineering tasks: tabular ingest/egress, lightweight reporting, and hand-off sheets. File sizes vary widely; Excel caps at **1,048,576 rows × 16,384 cols** per sheet and **~32,767 chars** per cell—your streaming model should assume big files but respect these limits. For Spark, many use **crealytics/spark-excel** (POI-backed). ([Microsoft Support][22])

## Findings

* Many pipelines just **read/write values**; formula evaluation is often unnecessary (display caches suffice). ([poi.apache.org][3])
* Spark ecosystems rely on **POI** under the hood; demand exists for **faster, streaming** readers. ([GitHub][31])

## Recommendations

* Provide **CSV-like streaming** ergonomics but with Excel niceties (styles, hyperlinks).
* Clear guidance for files exceeding Excel limits (split sheets, chunk writes). ([Microsoft Support][32])

---

# 9) Competing solutions analysis

## Executive summary

**POI** is feature-rich and battle-tested; alternatives optimize subsets: **fastexcel** (fast/low-mem Java), **openpyxl/xlsxwriter** (Python), **ExcelJS/SheetJS** (JS), **excelize** (Go), **calamine/rust_xlsxwriter** (Rust). A Scala-native, purely functional, streaming library is differentiated. ([poi.apache.org][26])

**Key takeaways.**

* **POI**: full feature surface incl. evaluator; heavier memory unless event APIs are used. ([poi.apache.org][26])
* **fastexcel (Java)**: very fast, but narrower feature set; often “write-only” workflows. ([GitHub][33])
* **openpyxl/xlsxwriter**: Python staples; show API breadth and writer-focused design. ([Openpyxl][34])
* **ExcelJS/SheetJS**: wide adoption in Node/web; useful reference for formula & table UX. ([GitHub][35])
* **excelize (Go)** & **calamine/rust_xlsxwriter (Rust)**: strong streaming stories. ([GitHub][36])

---

# 10) OOXML extension points & advanced features

## Executive summary

Advanced pieces (PivotTables, charts, macros/VBA, data validation, conditional formatting, custom XML, Power Query) map to distinct parts with relationships; complexity varies. Prioritize **data validation** and **conditional formatting** first; treat **charts** and **PivotTables** as phase-2; **VBA (vbaProject.bin)** requires special content types and is often “copy-through.” ([Microsoft Learn][37])

## Findings

* **Pivot** uses `pivotCacheDefinition` → `pivotCacheRecords` → `pivotTableDefinition` graph. ([Microsoft Learn][38])
* **Charts** are DrawingML (`c:chartSpace`/`c:chart`) in separate parts linked from drawings. ([Microsoft Learn][39])
* **Macros**: `.xlsm` uses macro-enabled content types; VBA stored in `xl/vbaProject.bin` with `application/vnd.ms-office.vbaProject`. Most libs copy the part wholesale. ([Microsoft Learn][40])
* **Data validation**: `dataValidations > dataValidation` with type/operator/formula1/2. ([c-rex.net][41])
* **Conditional formatting**: `conditionalFormatting > cfRule` referencing `dxfId`. ([Microsoft Learn][42])
* **Custom XML**: arbitrary data stored as custom parts; Power Query adds a **Data Mashup** binary part. ([Microsoft Learn][43])

---

# Synthesis & decisions

1. **Core vs. high-level modules.**
   **Yes**:

* `core-opc-xml` (OPC + XML codecs + streaming),
* `sml-core` (cell/row/worksheet/workbook, styles, hyperlinks),
* `sml-dsl` (ranges, formula AST, builders),
* `sml-interop` (converters, Spark helpers),
* `sml-eval` (optional, future).

2. **Performance trade-offs.**
   Default to streaming reads/writes with persistent structures; offer a small in-mem model for small sheets. For huge writes, **window rows** to disk. Use **inlineStr** by default; auto-promote to sharedStrings if dedup ratio wins. Zip with Commons Compress Zip64. ([poi.apache.org][5])

3. **v1.0 scope.**

* ✅ Read/write core (workbook/worksheets, values, basic styles, hyperlinks)
* ✅ Streaming APIs (fs2)
* ✅ Shared/inline strings
* ✅ Basic number/date formats
* ✅ Formula AST + writer, no evaluator (set recalc)
* ✅ Data validation + conditional formatting (common rules)
* ⏩ v1.x: charts (write), tables (structured refs in writer), themes, named ranges helpers
* ⏩ v2: PivotTables, full evaluator, Power Query interop, macro authoring

4. **Dependencies.**
   Essentials: **cats-effect, fs2, fs2-data-xml, Commons Compress**; optional: **Aalto (StAX)**. For derivation: native Scala 3 `derives` first; fall back to **Magnolia** if needed; optics via **Monocle**. ([fs2-data.gnieh.org][2])

5. **Success metrics.**

* Streaming read/write throughput vs POI SAX/SXSSF on 10–500 MB files
* Peak memory per million cells
* API ergonomics (lines of code for common tasks)
* Style dedup effectiveness (unique style count vs. raw)
* Interop (files open cleanly across Excel/LibreOffice/Numbers)

---

# Appendix A — Common pitfalls & gotchas

* **1900 leap-year bug** & 1900/1904 offsets (always abstract epoch). ([Microsoft Learn][44])
* **Dimension** element is **optional**; don’t rely on it for used range. ([Microsoft Learn][9])
* **Too many styles** (64k cap). Hash-dedup styles centrally. ([Microsoft Learn][45])
* **Precision**: Excel doubles ~15 significant digits. Avoid silent BigDecimal→Double surprises. ([Microsoft Learn][17])
* **Dynamic arrays & `#SPILL!`**: spill ranges need empties; your writer should avoid blocking spill areas when possible. ([Microsoft Support][19])

---

# Appendix B — File tree & relationship map (diagram)

```
/[Content_Types].xml
/_rels/.rels                         --> officeDocument -> xl/workbook.xml
/xl/workbook.xml                    (sheets -> r:id)
xl/_rels/workbook.xml.rels           --> rId1: worksheet -> worksheets/sheet1.xml
/xl/worksheets/sheet1.xml
[optional]
/xl/styles.xml
/xl/sharedStrings.xml
/xl/theme/theme1.xml
/docProps/core.xml
```

(OPC relationships resolve via sibling `_rels/*.rels` files; types are fixed URIs.) ([Wikipedia][6])

---

# Appendix C — Limits you should surface in docs

* Rows × cols: **1,048,576 × 16,384**; chars per cell ~**32,767**; unique formats ~**65k**. ([Microsoft Support][22])

---

## What’s next (implementation roadmap)

1. **scalaxlsx-core**: OPC + streaming ZIP IO + fs2 XML cursors; read minimal workbook, emit rows. ([fs2-data.gnieh.org][2])
2. **Writer**: generate minimal package; inline strings; basic styles; hyperlinks. ([Microsoft Learn][16])
3. **SharedStrings** pool and auto-promotion heuristic. ([Microsoft Learn][8])
4. **Styles** pool + dedup; number formats presets. ([Office Open XML][4])
5. **Formula AST + DSL**; write `f` elems; `forceRecalc` flag. ([poi.apache.org][3])
6. **Validation & conditional formatting** writers. ([c-rex.net][41])
7. **Benchmarks** vs POI SAX/SXSSF on 10–500 MB sheets; publish memory/latency. ([poi.apache.org][5])

If you want, I can spin this into a repo skeleton (Scala 3.7, cats-effect, fs2, fs2-data-xml, Commons Compress) with the minimal reader/writer and test files ready to open in Excel.

[1]: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview?utm_source=chatgpt.com "Open Packaging Conventions Fundamentals"
[2]: https://fs2-data.gnieh.org/documentation/xml/?utm_source=chatgpt.com "Introduction - fs2-data"
[3]: https://poi.apache.org/components/spreadsheet/eval.html?utm_source=chatgpt.com "Formula Evaluation - Apache POI"
[4]: https://officeopenxml.com/SSstyles.php?utm_source=chatgpt.com "Spreadsheets - Overview of Styles"
[5]: https://poi.apache.org/components/spreadsheet/limitations.html?utm_source=chatgpt.com "Apache POI™ - HSSF and XSSF Limitations"
[6]: https://en.wikipedia.org/wiki/Open_Packaging_Conventions?utm_source=chatgpt.com "Open Packaging Conventions"
[7]: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/structure-of-a-spreadsheetml-document?utm_source=chatgpt.com "Structure of a SpreadsheetML document"
[8]: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-shared-string-table?utm_source=chatgpt.com "Working with the shared string table"
[9]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetdimension?view=openxml-3.0.1&utm_source=chatgpt.com "SheetDimension Class (DocumentFormat.OpenXml. ..."
[10]: https://schemas.liquid-technologies.com/officeopenxml/2006/sml-supplementaryworkbooks_xsd.html?utm_source=chatgpt.com "Schema Name - XML Standards Library - Liquid Technologies"
[11]: https://en.wikipedia.org/wiki/Office_Open_XML_file_formats?utm_source=chatgpt.com "Office Open XML file formats"
[12]: https://www.ericwhite.com/blog/advice-when-generating-spreadsheets-use-inline-strings-not-shared-strings/?utm_source=chatgpt.com "Advice: When Generating Spreadsheets, Use Inline Strings ..."
[13]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cell?view=openxml-3.0.1&utm_source=chatgpt.com "Cell Class (DocumentFormat.OpenXml.Spreadsheet)"
[14]: https://support.microsoft.com/en-us/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487?utm_source=chatgpt.com "Date systems in Excel"
[15]: https://support.microsoft.com/en-us/office/hide-error-values-and-error-indicators-in-cells-d171b96e-8fb4-4863-a1ba-b64557474439?utm_source=chatgpt.com "Hide error values and error indicators in cells"
[16]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.hyperlinks?view=openxml-3.0.1&utm_source=chatgpt.com "Hyperlinks Class (DocumentFormat.OpenXml.Spreadsheet)"
[17]: https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/floating-point-arithmetic-inaccurate-result?utm_source=chatgpt.com "Floating-point arithmetic may give inaccurate result in Excel"
[18]: https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e?utm_source=chatgpt.com "Using structured references with Excel tables"
[19]: https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531?utm_source=chatgpt.com "Dynamic array formulas and spilled array behavior"
[20]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformula?view=openxml-3.0.1&utm_source=chatgpt.com "CellFormula Class (DocumentFormat.OpenXml.Spreadsheet)"
[21]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyleformats?view=openxml-3.0.1&utm_source=chatgpt.com "CellStyleFormats Class (DocumentFormat.OpenXml. ..."
[22]: https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3?utm_source=chatgpt.com "Excel specifications and limits"
[23]: https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_clrScheme_topic_ID0ES2FMB.html?utm_source=chatgpt.com "clrScheme (Color Scheme)"
[24]: https://support.microsoft.com/en-us/office/number-format-codes-in-excel-for-mac-5026bbd6-04bc-48cd-bf33-80f18b4eae68?utm_source=chatgpt.com "Number format codes in Excel for Mac"
[25]: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting?utm_source=chatgpt.com "Working with conditional formatting"
[26]: https://poi.apache.org/components/spreadsheet/?utm_source=chatgpt.com "POI-HSSF and POI-XSSF/SXSSF - Java API To Access ..."
[27]: https://commons.apache.org/proper/commons-compress/zip.html?utm_source=chatgpt.com "Commons Compress ZIP package"
[28]: https://fs2-data.gnieh.org/?utm_source=chatgpt.com "fs2-data"
[29]: https://typelevel.org/doobie/docs/12-Custom-Mappings.html?utm_source=chatgpt.com "Custom Mappings · doobie"
[30]: https://http4s.org/v0.20/dsl/?utm_source=chatgpt.com "The http4s DSL"
[31]: https://github.com/nightscape/spark-excel?utm_source=chatgpt.com "A Spark plugin for reading and writing Excel files"
[32]: https://support.microsoft.com/en-us/office/what-to-do-if-a-data-set-is-too-large-for-the-excel-grid-976e6a34-9756-48f4-828c-ca80b3d0e15c?utm_source=chatgpt.com "What to do if a data set is too large for the Excel grid"
[33]: https://github.com/dhatim/fastexcel?utm_source=chatgpt.com "dhatim/fastexcel: Generate and read big Excel files quickly"
[34]: https://openpyxl.readthedocs.io/?utm_source=chatgpt.com "OpenPyXL - Read the Docs"
[35]: https://github.com/exceljs/exceljs?utm_source=chatgpt.com "exceljs/exceljs: Excel Workbook Manager"
[36]: https://github.com/qax-os/excelize?utm_source=chatgpt.com "qax-os/excelize: Go language library for reading and ..."
[37]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.datavalidations?view=openxml-3.0.1&utm_source=chatgpt.com "DataValidations Class (DocumentFormat.OpenXml. ..."
[38]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pivotcachedefinition?view=openxml-3.0.1&utm_source=chatgpt.com "PivotCacheDefinition Class (DocumentFormat.OpenXml. ..."
[39]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.chartspace?view=openxml-3.0.1&utm_source=chatgpt.com "ChartSpace Class (DocumentFormat.OpenXml.Drawing. ..."
[40]: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-offmacro2/fa1f6007-088f-4f54-ba65-83b5a9d635a6?utm_source=chatgpt.com "[MS-OFFMACRO2]: Workbook"
[41]: https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_dataValidations_topic_ID0EX4W4.html?utm_source=chatgpt.com "dataValidations (Data Validations)"
[42]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.conditionalformattingrule?view=openxml-3.0.1&utm_source=chatgpt.com "ConditionalFormattingRule Class"
[43]: https://learn.microsoft.com/en-us/visualstudio/vsto/how-to-add-custom-xml-parts-to-document-level-customizations?view=vs-2022&utm_source=chatgpt.com "Add custom XML parts to document-level customizations"
[44]: https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/wrongly-assumes-1900-is-leap-year?utm_source=chatgpt.com "Excel incorrectly assumes that the year 1900 is a leap year"
[45]: https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/too-many-different-cell-formats-in-excel?utm_source=chatgpt.com "You receive a Too many different cell formats error ..."
