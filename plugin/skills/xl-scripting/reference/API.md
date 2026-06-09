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
| `Patch` | Composable change set (monoid): `Put`, `SetCellStyle`, `SetRangeStyle`, `Merge`, `Remove`, `Batch`, ... |
| `CellStyle` | font, fill, border, numFmt, alignment |
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
Patch.empty                     // identity — fold seed
```

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
  .bordered / .borderedMedium / .borderedThick / .borderNone
  .currency / .percent / .decimal / .dateFormat / .dateTime   // numFmt shortcuts
  .withNumFmt(NumFmt.Percent)
```

Rich text: `"Bold ".bold + "red".red + " plain"` → `RichText`, put directly into a cell.

Smart detection: `FormattedParsers.detect(s)` / `s.toFormatted` (prelude) — total: `"$1,234.56"` → Currency, `"45.5%"` → Percent (stored 0.455), `"2025-01-15"` → Date, `"123.4"`/`"true"` → Number/Bool (General), anything else → Text (General).

## Formula Evaluation

| Method | Returns | Notes |
|--------|---------|-------|
| `wb.evaluateFormula(formula, onSheet[, clock])` | `XLResult[CellValue]` | cross-sheet context automatic; onSheet: String or SheetName |
| `wb.recalculate([clock])` | `RecalcResult` | total whole-workbook recalc, per-cell errors, cycle isolation |
| `wb.withCachedFormulas([clock])` | `Workbook` | = `recalculate(clock).workbook` |
| `sheet.evaluateFormula(formula[, clock][, workbook])` | `XLResult[CellValue]` | pass `Some(wb)` iff formula references other sheets |
| `sheet.evaluateWithDependencyCheck([clock])` | `XLResult[Map[ARef, CellValue]]` | fail-fast on first error/cycle |
| `FormulaParser.parse("=A1+B1")` | `Either[ParseError, TExpr[?]]` | AST access |
| `DependencyGraph.fromSheet(sheet)` | `DependencyGraph` | `.precedents(ref)` / `.dependents(ref)` |
| `Clock.system` / `Clock.fixedDate(LocalDate)` | `Clock` | deterministic TODAY()/NOW() |

`RecalcResult`: `.workbook` (successful formulas cached), `.evaluated: Map[SheetName, Map[ARef, CellValue]]`, `.errors: Vector[CellEvalError]`, `.isClean`, `.toEither`.

### All 104 functions

SUM, SUMIF, SUMIFS, SUMPRODUCT, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, AVERAGE, AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS, MEDIAN, STDEV, STDEVP, VAR, VARP, LARGE, SMALL, RANK, PERCENTILE, QUARTILE, MIN, MAX, IF, IFS, IFERROR, SWITCH, CHOOSE, AND, OR, NOT, ISNUMBER, ISTEXT, ISBLANK, ISERR, ISERROR, CONCATENATE, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, FIND, SUBSTITUTE, TEXT, VALUE, TODAY, NOW, DATE, YEAR, MONTH, DAY, EOMONTH, EDATE, DATEDIF, NETWORKDAYS, WORKDAY, YEARFRAC, ABS, ROUND, ROUNDUP, ROUNDDOWN, INT, MOD, POWER, SQRT, LOG, LN, EXP, FLOOR, CEILING, TRUNC, SIGN, PMT, FV, PV, RATE, NPER, NPV, IRR, XNPV, XIRR, VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH, OFFSET, PI, ROW, COLUMN, ROWS, COLUMNS, ADDRESS, TRANSPOSE, SEQUENCE, SORT, UNIQUE, FILTER

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
