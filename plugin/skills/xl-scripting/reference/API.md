# XL Scripting API Reference

Everything below is in scope after `import com.tjclp.xl.scripting.{*, given}`.

## Key Types

| Type | Description |
|------|-------------|
| `ARef` | Cell reference (opaque, 64-bit packed). `ref"A1"` |
| `CellRange` | Inclusive rectangular range, auto-normalized. `ref"A1:B10"` |
| `RefType` | Cell \| Range \| QualifiedCell \| QualifiedRange — what runtime `ref"$s"` parses to |
| `SheetName` | Validated sheet name (≤31 chars, no `:\/?*[]`). `SheetName("Q1")` → `Either` |
| `Cell` | `(ref, value, styleId, ...)` |
| `CellValue` | enum: `Text`, `Number(BigDecimal)`, `Bool`, `DateTime`, `Formula(expr, cached)`, `RichText`, `Error`, `Empty` |
| `Sheet` | Immutable sheet: `cells: Map[ARef, Cell]`, name, merges, comments, col/row properties |
| `Workbook` | `sheets: Vector[Sheet]` + metadata |
| `Patch` | Composable change set (monoid): `Put`, `SetCellStyle`, `SetRangeStyle`, `Merge`, `MergeBorder`, `Remove`, `Batch`, ... |
| `CellStyle` | font, fill, border, numFmt, alignment |
| `SheetView` | display settings: `(showGridLines, zoomScale)` — see Sheet View & Print Setup |
| `PageSetup` | print settings: scale, orientation, fit, header/footer, margins, print area, repeat rows |
| `NumFmt` | `General`, `Currency`, `Percent`, `Date`, `DateTime`, `Decimal`, custom |
| `Formatted` | `(value: CellValue, numFmt: NumFmt)` — value + display format pair |
| `XLResult[A]` | `Either[XLError, A]` — every fallible operation |
| `XLError` | Structured error enum (`SheetNotFound`, `InvalidCellRef`, `FormulaError`, ...) with `.message` |
| `RecalcResult` | `(workbook, evaluated, errors)` from `wb.recalculate()` |
| `CellEvalError` | `(sheet, ref, error)` with `.render` → `"Sales!B2: ..."` |

## Compile-Time Literals (macros)

Invalid literals **fail compilation**. Runtime-interpolated forms return `Either[XLError, _]`.

| Literal | Result | Example |
|---------|--------|---------|
| `ref"A1"` | `ARef` | `ref"XFD1048576"` |
| `ref"A1:B10"` | `CellRange` | `ref"A:A"` rejected at compile time if malformed |
| `ref"A$i"` / `ref"$s"` | `Either[XLError, RefType]` | runtime validation |
| `fx"=SUM(A1:B10)"` | `CellValue.Formula` | parens/syntax checked |
| `fx"=B$i*2"` | `Either[XLError, CellValue]` | runtime validation |
| `money"$$1,234.56"` | `Formatted(Number, Currency)` | `$$` escapes `$` |
| `percent"45.5%"` | `Formatted(Number(0.455), Percent)` | stored as fraction |
| `date"2025-01-15"` | `Formatted(DateTime, Date)` | ISO format |
| `accounting"($$1,234)"` | `Formatted(Number(-1234), Currency)` | parens = negative |
| `col"A"` / `col(0)` | `Column` | |

## ARef / CellRange / Column / Row

```scala
ref"B3".toA1                 // "B3"
ref"B3".col / ref"B3".row    // Column / Row
ref"B3".shift(1, 2)          // D5 (colOffset, rowOffset) — unchecked bounds
ref"B3".down(2)              // B5   (also up/right/left, default n=1)
ref"B3".col.toLetter         // "B"
ref"B3".col.index0           // 1    (also index1)
ref"A1:B10".cells            // Iterator[ARef], row-major
ref"A1:B10".width / .height / .size
ARef.parse("C3")             // Either[String, ARef]
ARef.from0(colIdx, rowIdx)   // 0-based construction
"C3".asCell                  // XLResult[ARef]   (String helpers; also .asRange, .asSheetName)
```

## Sheet Operations

| Method | Returns | Notes |
|--------|---------|-------|
| `Sheet("Name")` | `Sheet` | literal validated at compile time |
| `sheet.put(ref"A1", value)` | `Sheet` | value: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText, CellValue, Formatted |
| `sheet.put(ref"A1", value, style)` | `Sheet` | put with inline style |
| `sheet.put("A1", value)` | `Sheet` (literal) / `XLResult[Sheet]` (runtime string) | |
| `sheet.put(patch)` | `Sheet` | apply a composed Patch |
| `sheet.style(ref"A1:D1", style)` | `Sheet` | merges into existing style |
| `sheet.cell("A1")` / `sheet.range("A1:B3")` | `Option[Cell]` / `Iterable[Cell]` | safe lookups |
| `sheet.cells` | `Map[ARef, Cell]` | |
| `sheet.readTyped[A](ref)` | `Either[CodecError, Option[A]]` | distinguish mismatch from empty |
| `sheet.readTypedOr[A](ref, default)` | `A` | total |
| `sheet.readTypedOpt[A](ref)` | `Option[A]` | total, flat |
| `sheet.comment(ref, Comment.plainText("note", Some("author")))` | `Sheet` | |
| `sheet.toHtml(ref"A1:B10")` | `String` | inline-CSS HTML table |
| `sheet.usedRange` | `Option[CellRange]` | |
| `sheet.withViewSettings(view)` | `Sheet` | gridlines/zoom — see Sheet View & Print Setup |
| `sheet.withPageSetup(setup)` | `Sheet` | print settings — see Sheet View & Print Setup |

Codec types for `put`/`readTyped*`: String, Int, Long, Double, BigDecimal, Boolean, LocalDate (→ Date format), LocalDateTime (→ DateTime format), RichText.

## Workbook Operations

| Method | Returns | Notes |
|--------|---------|-------|
| `Workbook(sheet1, sheet2)` | `Workbook` | |
| `Workbook("Name")` | `Workbook` (literal) | one named sheet |
| `Workbook.empty` | `Workbook` | contains default "Sheet1" |
| `wb.sheets` | `Vector[Sheet]` | |
| `wb("Sales")` / `wb(SheetName)` / `wb(0)` | `XLResult[Sheet]` | |
| `wb.put(sheet)` | `Workbook` | add-or-replace by name, total |
| `wb.upsert("Name", f: Sheet => Sheet)` | `Workbook` (literal) / `XLResult` (runtime) | update-or-create, total |
| `wb.update("Name", f)` | `XLResult[Workbook]` | fails if sheet absent |
| `wb.remove("Name")` | `XLResult[Workbook]` | can't remove last sheet |
| `wb.rename(old, new)` | `XLResult[Workbook]` | |

## Patch DSL

Patches are pure values; `++` composes (monoid, last-write-wins per cell). Apply with `sheet.put(patch)`.

```scala
(ref"A1" := "Title")            // Put — overloads: CellValue, String, Int, Long, Double, BigDecimal, Boolean, LocalDateTime
(ref"A1:C10" := 0)              // range fill: every cell gets the value (Excel Ctrl+Enter)
ref"A1".styled(style)           // style one cell
ref"A1:C1".styled(style)        // style a range
ref"A1:C1".merge                // merge cells (also .unmerge)
ref"A1:C10".remove              // clear cells
ref"A1".clearStyle
ref"B2:D6".outlined(BorderStyle.Medium)                       // box the range's outer edges
ref"B2:D6".outlined(BorderStyle.Thin, Color.fromRgb(128, 128, 128))  // colored outline
Patch.empty                     // identity — fold seed
```

`outlined` is edge-correct (top row gets top, corners get both, interior untouched; 1×1 gets all
four sides) and desugars to per-cell `Patch.MergeBorder(ref, border)` — an apply-time border
overlay that preserves each cell's font, fill, numFmt, and untouched border sides. Cost is
proportional to the perimeter; prefer bounded ranges over whole columns.

Runtime `RefType` (from `ref"$s"`) supports the same `:=` / `.styled` / `.merge` / `.remove`; `:=` on a parsed range fills it.

## Styling

```scala
CellStyle.default
  .bold.italic.underline
  .size(12.0).fontFamily("Calibri")
  .red / .green / .blue / .black / .white / .yellow / .hex("FF8800") / .rgb(255, 136, 0)
  .bgBlue / .bgGray / .bgGreen / .bgRed / .bgWhite / .bgYellow / .bgHex("EEEEEE") / .bgNone
  .center / .left / .right          // horizontal align
  .top / .middle / .bottom          // vertical align
  .wrap
  .indent(2)                        // alignment indent level (~3 chars each); negatives clamp to 0
  .bordered / .borderedMedium / .borderedThick / .borderNone
  .borderTop(BorderStyle.Thin)      // per-side: also borderBottom/borderLeft/borderRight
  .borderBottom(BorderStyle.Medium, Color.fromRgb(0, 0, 128))  // color overloads on each side
  .currency / .percent / .decimal / .dateFormat / .dateTime   // numFmt shortcuts
  .withNumFmt(NumFmt.Percent)
```

Per-side border builders merge into the existing border — only the named side is replaced, so
they compose: `.borderTop(BorderStyle.Thin).borderBottom(BorderStyle.Medium)`. Indentation lives
in the style (survives round-trips, keeps the stored value clean) — prefer it over leading spaces.

Rich text: `"Bold ".bold + "red".red + " plain"` → `RichText`, put directly into a cell.

Smart detection: `FormattedParsers.detect(s)` / `s.toFormatted` (prelude) — total: `"$1,234.56"` → Currency, `"45.5%"` → Percent (stored 0.455), `"2025-01-15"` → Date, `"123.4"`/`"true"` → Number/Bool (General), anything else → Text (General).

## Sheet View & Print Setup

The four settings types live one import deeper than the prelude:

```scala
import com.tjclp.xl.sheets.{HeaderFooter, PageMargins, PageSetup, SheetView}
```

```scala
// Display: gridlines off (professional templates), zoom 10-400
sheet.withViewSettings(SheetView(showGridLines = false, zoomScale = Some(90)))

// Print/PDF setup — all fields optional with Excel defaults
sheet.withPageSetup(
  PageSetup(
    scale = 100,                                              // 10-400
    orientation = Some("landscape"),                          // or "portrait"
    fitToWidth = Some(1),                                     // pages wide (also fitToHeight; emits pageSetUpPr fitToPage)
    headerFooter = Some(HeaderFooter(
      oddFooter = Some("&CPage &P of &N"),
      firstHeader = Some("&CCONFIDENTIAL"), differentFirst = true // 0.11.1+: even/first variants
    )),                                                       // also evenHeader/evenFooter + differentOddEven
    margins = Some(PageMargins(left = 0.5, right = 0.5)),     // inches; defaults match Excel Normal
    printArea = Some(ref"A1:F40"),                            // _xlnm.Print_Area defined name
    repeatRows = Some((1, 2))                                 // 1-based rows repeated on every page
  )
)
```

Header/footer strings support Excel codes — `&P` page, `&N` total pages, `&D` date, `&T` time,
`&F` file, `&A` sheet, with `&L`/`&C`/`&R` section markers. View settings share one
`<sheetView>` element with freeze panes; print area/repeat rows ride the defined-names pipeline.

## Formula Evaluation

| Method | Returns | Notes |
|--------|---------|-------|
| `wb.evaluateFormula(formula, onSheet[, clock])` | `XLResult[CellValue]` | cross-sheet context automatic; onSheet: String or SheetName |
| `wb.recalculate([clock][, rng])` | `RecalcResult` | total whole-workbook recalc, per-cell errors, cycle isolation; `rng` (0.11.2+) seeds RAND — `Rng.seeded(42L)` for reproducible scripts |
| `wb.withCachedFormulas([clock])` | `Workbook` | = `recalculate(clock).workbook` |
| `sheet.evaluateFormula(formula[, clock][, rng][, workbook])` | `XLResult[CellValue]` | pass `Some(wb)` iff formula references other sheets |
| `sheet.putFormulaInheriting(ref, formula[, workbook])` | `XLResult[Sheet]` | 0.11.2+: puts the formula AND inherits the referenced cells' number format into a General target (Excel's entry behavior) |
| `sheet.evaluateWithDependencyCheck([clock])` | `XLResult[Map[ARef, CellValue]]` | fail-fast on first error/cycle |
| `FormulaParser.parse("=A1+B1")` | `Either[ParseError, TExpr[?]]` | AST access |
| `DependencyGraph.fromSheet(sheet)` | `DependencyGraph` | `.precedents(ref)` / `.dependents(ref)` |
| `Clock.system` / `Clock.fixedDate(LocalDate)` | `Clock` | deterministic TODAY()/NOW() |

`RecalcResult`: `.workbook` (successful formulas cached), `.evaluated: Map[SheetName, Map[ARef, CellValue]]`, `.errors: Vector[CellEvalError]`, `.isClean`, `.toEither`.

### All 107 functions

SUM, SUMIF, SUMIFS, SUMPRODUCT, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, AVERAGE, AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS, MEDIAN, STDEV, STDEVP, VAR, VARP, LARGE, SMALL, RANK, PERCENTILE, QUARTILE, MIN, MAX, IF, IFS, IFERROR, SWITCH, CHOOSE, AND, OR, NOT, ISNUMBER, ISTEXT, ISBLANK, ISERR, ISERROR, CONCATENATE, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, FIND, SUBSTITUTE, TEXT, VALUE, TODAY, NOW, DATE, YEAR, MONTH, DAY, EOMONTH, EDATE, DATEDIF, NETWORKDAYS, WORKDAY, YEARFRAC, ABS, ROUND, ROUNDUP, ROUNDDOWN, INT, MOD, POWER, SQRT, LOG, LN, EXP, FLOOR, CEILING, TRUNC, SIGN, PMT, FV, PV, RATE, NPER, NPV, IRR, XNPV, XIRR, VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH, OFFSET, INDIRECT, PI, ROW, COLUMN, ROWS, COLUMNS, ADDRESS, TRANSPOSE, SEQUENCE, SORT, UNIQUE, FILTER, RAND, RANDBETWEEN — plus LET (lexical bindings; a parser-level special form, not in the registry listing)

## IO

### Sync facade (`Excel`) — for scripts

| Method | Notes |
|--------|-------|
| `Excel.read(path: String): Workbook` | throws `XLException`/IO errors at this edge |
| `Excel.write(wb, path: String): Unit` | also accepts `XLResult[Workbook]` |
| `Excel.modify(path)(f: Workbook => Workbook): Unit` | atomic in-place replacement |

### Streaming (`ExcelIO`) — 100k+ rows, O(1) memory

```scala
import cats.effect.IO
import cats.effect.unsafe.implicits.global
import java.nio.file.Paths

val excel = ExcelIO.instance[IO]
excel.readStream(path)                        // Stream[IO, RowData] — first sheet
excel.readSheetStream(path, "Sales")          // by name
excel.writeStream(path, "Sheet1")             // fs2.Pipe[IO, RowData, Unit]
excel.writeStreamsSeq(path, Seq("S1" -> rows1, "S2" -> rows2))  // multi-sheet, sequential

case class RowData(rowIndex: Int /* 1-based */, cells: Map[Int /* 0-based col */, CellValue])
```

Run effects at the script edge with `.unsafeRunSync()`.

## Display

```scala
given Sheet = sheet
println(excel"A1 = ${ref"A1"}")     // formats through NumFmt ($1,234.56, 45.5%, ...)
sheet.displayCell(ref"A1")          // String
```

## Errors

```scala
XLResult[A]                  // = Either[XLError, A]
err.message                  // human-readable
result.unsafe                // throws XLException(err) — the one sanctioned unwrap (prelude)
result.getOrElse(fallback)
```
