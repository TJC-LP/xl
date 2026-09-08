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
| `Patch` | Composable change set (monoid): `Put`, `SetCellStyle`, `SetRangeStyle`, `Merge`, `MergeBorder`, `SetComment`, `SetConditionalFormat`, `Remove`, `Batch`, ... |
| `CellStyle` | font, fill, border, numFmt, alignment (incl. `textRotation`) |
| `SheetView` | display settings: `(showGridLines, zoomScale, tabSelected)` — see Sheet View & Print Setup |
| `PageSetup` | print settings: scale, orientation, fit, header/footer, margins, print area, repeat rows |
| `DataValidation` | `.Rules(ranges, kind, allowBlank, showDropdown)` \| `.Preserved` — list dropdowns via `DataValidation.list`/`listOf` |
| `CalcPr` | workbook `<calcPr>`: `(iterativeCalculation, maxIterations, maxChange, calcMode, fullCalcOnLoad, calcId)` — iterative-calc + calculation-mode authoring; `CalcMode.Manual/Auto/AutoNoTable` (0.14.0) |
| `IterativeCalc` | `(maxIter, maxChange)` — opt-in bounded circular recalc; `fromCalcPr` bridges a file's `CalcPr` |
| `NumFmt` | `General`, `Currency`, `Percent`, `Date`, `DateTime`, `Decimal`, custom |
| `Formatted` | `(value: CellValue, numFmt: NumFmt)` — value + display format pair |
| `XLResult[A]` | `Either[XLError, A]` — every fallible operation |
| `XLError` | Structured error enum (`SheetNotFound`, `InvalidCellRef`, `FormulaError`, ...) with `.message`; 0.21.1 adds `InvalidArgument(op, reason)` (`INVALID_ARGUMENT` — every workbook-free `Edit.validate` refusal, #617) and `NameNotFound(name, available)` (`NAME_NOT_FOUND`, with `did you mean` candidates like `SheetNotFound`, #626) |
| `RecalcResult` | `(workbook, evaluated, errors)` from `wb.recalculate()` |
| `RowCodec[A]` | one record ↔ one row of cells: `fields` (header names, column order), `read(cells)`, `write(a)`; `final case class Order(...) derives RowCodec` (0.21.0) |
| `RowCodecError` | `Field(row, column, field, cause)` \| `Missing(row, column, field)` \| `HeaderNotFound(header, headerRow, available)` \| `Width(expected, actual)` — `.message`, `.toXLError` (0.21.0) |
| `RowsPlaced` | `(sheet, headerRange, dataRange)` + `range` (header ∪ data) and `count` — what `putRows`/`putRowsWithHeader`/`putTable` return (0.21.0) |
| `CellEvalError` | `(sheet, ref, error)` with `.render` → `"Sales!B2: ..."` |
| `TExpr[A]` | typed formula AST from `FormulaParser.parse`. **Breaking in 0.21.0** ([#612](https://github.com/TJC-LP/xl/issues/612)): `RangeRef`/`SheetRange`/`ExternalRange` and `RangeLocation.Local`/`CrossSheet`/`External` gained a trailing `form: RangeForm = RangeForm.Cells` field (a positional match on any of the six needs one more `_`), and `TExpr.ErrorLit(err)` / `RangeLocation.Error(err)` are new arms an exhaustive match must add — error literals (`#REF!`, `#N/A`, `#DIV/0!`, …) parse to them and print back verbatim. Constructor calls are source-compatible; a downstream artifact compiled against ≤0.20.0 must be rebuilt |
| `RangeForm` | `Cells` (`A1:B2`) \| `Columns` (`A:A`, `$A:C`) \| `Rows` (`1:1`, `$3:$10`) \| `Cell` (`C1` in a range slot) — the whole-column / whole-row marker on range nodes, so `A:A` prints as `A:A` (not `$A1:$A1048576`) and a drag moves it only along its axis; exported beside `TExpr` (0.21.0). **Breaking in 0.21.1** ([#631](https://github.com/TJC-LP/xl/issues/631)): `Cell` is a fourth case — `SUMIF(A1:A10, ">0", C1)` parses, prints back as `C1` and drags as the cell it is; an exhaustive `match` over `RangeForm` needs the arm |
| `CellError` | the Excel error values a cell can hold or a formula can name: `Div0`, `NA`, `Name`, `Null`, `Num`, `Ref`, `Value`. **Breaking in 0.21.1** ([#630](https://github.com/TJC-LP/xl/issues/630)): seven modern cases — `Spill`, `Calc`, `Field`, `Connect`, `Blocked`, `Unknown`, `GettingData` — parse (case-insensitively, `CellError.parse`), print, evaluate and round-trip, so an Excel 365 book with a `#SPILL!` cell opens; an exhaustive `match` over `CellError` needs the arms. `error.errorTypeNumber` is `ERROR.TYPE`'s number |
| `Edit` | enum, 49 cases — the batch/CLI operation vocabulary as values (`Put`, `DragFormula`, `Fill`, `Copy`, `Sort`, `Clear`, `Style`, `Merge`, `InsertRows`, `AddSheet`, `RenameSheet`, …); `Edit.FillDir` / `Edit.SortDir` / `Edit.SortMode` / `Edit.SortKeySpec` nest in the companion (0.21.0) — see Edit algebra |
| `Loc` / `Area` | `(sheet: Option[SheetName], ref: ARef)` / `(sheet, range: CellRange)` — an edit's target; `None` = the scope's default sheet; `Loc.parse("'Q1 Data'!B7")` / `Area.parse("A1:B2")`; `Area.cell(loc)` (0.21.0) |
| `RowSpan` / `ColSpan` | inclusive, normalized `10:20` / `E:H`; `RowSpan.parse` / `ColSpan.parse` refuse the other axis; `.rows` / `.columns`, `.size`, `.render` (0.21.0) |
| `EditScope` | `(defaultSheet: Option[SheetName])` — `EditScope.none`, `EditScope.of(sheet)`, `scope.after(edit)`; the prelude's name for `ops.Scope` (0.21.0) |
| `FormatHint` | `Inferred(fmt)` (a codec's hint: fills a General format only) \| `Explicit(fmt)` (replaces the number format) (0.21.0) |
| `StyleOverlay` | partial style, every field `Option`; `++` right-biased with `StyleOverlay.empty` as identity; `.of(style)` / `.ofBorder(border)` / `.applyTo(style)`; paired with `StyleMode.Merge` \| `Replace` (0.21.0) |
| `ClearWhat` | `(contents, styles, comments)` — `ClearWhat.contents` / `.styles` / `.comments` / `.all` (0.21.0) |
| `Applied` / `Planned` | `Edit.applyAll`'s result `(workbook, planned, scope)` + `touchedBySheet`, `structural`; one `Planned(index, edit, sheet, touched)` per edit (0.21.0) |
| `EditSchema` / `EditSpec` | the algebra's own metadata: one `EditSpec(name, aliases, batchOp, cliVerb, fields, …)` per `Edit` case (49) — `EditSchema.all`, `find("putf")`, `nameOf(edit)`; the shape of the batch schema, but `xl batch --schema` prints xl-cli's `OpRegistry` (32 ops), a subset (0.21.0) |
| `FormulaSupport` | the parser seam the formula-aware edits need; the prelude's `given` is xl-evaluator's, `FormulaSupport.textOnly` refuses drags/structural/rename with `UnsupportedCapability` (0.21.0) |
| `DisplayWrapper` | `(formatted: String)` with `toString = formatted` — what `sheet.displayCell` returns and what `excel""`/`s""` interpolation renders |

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
Column.parse("D")            // Either[String, Column] — runtime column letter (0.13.0; "D1" tolerated)
RefType.parse("Sales!C2:E9").map(_.col)  // Right(C) — runtime ref's (starting) column (0.13.0)
"C3".asCell                  // XLResult[ARef]   (String helpers; also .asRange, .asSheetName)
```

## Sheet Operations

| Method | Returns | Notes |
|--------|---------|-------|
| `Sheet("Name")` | `Sheet` (literal) / `XLResult[Sheet]` (runtime string) | literal validated at compile time; the specialization is invisible at the call site — prefer `Sheet.named` for runtime names ([#420](https://github.com/TJC-LP/xl/issues/420)) |
| `Sheet.named(name)` | `XLResult[Sheet]` | THE dynamic-name factory (0.17.0): same validation as `Sheet(name)`, result type spelled in the signature — `Sheet.named(nm).map(_.put(ref"A1", 1))` |
| `sheet.put(ref"A1", value)` | `Sheet` | value: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText, CellValue, Formatted |
| `sheet.put(ref"A1", value, style)` | `Sheet` | put with inline style |
| `sheet.put("A1", value)` | `Sheet` (literal) / `XLResult[Sheet]` (runtime string) | |
| `sheet.put(patch)` | `Sheet` | apply a composed Patch |
| `sheet.style(ref"A1:D1", style)` | `Sheet` | merges into existing style |
| `sheet.cell("A1")` / `sheet.range("A1:B3")` | `Option[Cell]` / `Iterable[Cell]` | safe lookups |
| `sheet.cells` | `Map[ARef, Cell]` | |
| `sheet.readTyped[A](ref)` | `Either[CodecError, Option[A]]` | distinguish mismatch from empty; since 0.20.0 (GH-477) a formula cell decodes as its cached value and only an uncached formula is a `TypeMismatch`; on ≤0.19.3 every formula cell is a `TypeMismatch` |
| `sheet.readTypedOr[A](ref, default)` | `A` | total; 0.20.0: cached formula → its value, uncached → default (≤0.19.3: any formula → default) |
| `sheet.readTypedOpt[A](ref)` | `Option[A]` | total, flat; 0.20.0: cached formula → `Some`, uncached → `None` (≤0.19.3: any formula → `None`) |
| `sheet.readTypedStrict[A](ref)` | `Either[CodecError, Option[A]]` | 0.20.0: like `readTyped`, but ANY formula cell is `Left(TypeMismatch(expected, formula))`, cached or not (GH-477) |
| `sheet.readRows[A](range)` | `Either[RowCodecError, Vector[A]]` | one record per row of `range`, positional (`fields(i)` ↔ column i); range width must equal the record's, else `Width`; a blank row is `Missing` unless every field is `Option` (0.21.0, needs `RowCodec[A]`) |
| `sheet.readRowsByHeader[A](headerRow)` | `Either[RowCodecError, Vector[A]]` | fields matched to header text via `columnOf` (exact, then case/space/`_`/`-`-insensitive), any column order, extra columns ignored; reads the contiguous block under the header, stops at the first blank row (0.21.0) |
| `sheet.columnHeaders(row)` | `Vector[(Column, String)]` | header texts in `row`, left to right, verbatim (numbers/rich text as text; blanks skipped) (0.21.0) |
| `sheet.columnOf(header, headerRow)` | `Option[Column]` | the column headed `header`: exact match, else case/whitespace/`_`/`-`-insensitive, leftmost wins (0.21.0) |
| `sheet.putRows(at, records)` | `XLResult[RowsPlaced]` | one row per record from `at`, no header; `None` fields stay empty; codec formats register like `put`; only the records' cells are written (clear a longer old block first); `OutOfBounds` past XFD/1048576 (0.21.0) |
| `sheet.putRowsWithHeader(at, records)` | `XLResult[RowsPlaced]` | field names as a header row at `at`, records below (0.21.0) |
| `sheet.putTable(at, records, name)` | `XLResult[RowsPlaced]` | `putRowsWithHeader` + an Excel table named/columned after the record; `name` = display name, letters/digits/`_`, unique on the sheet; no records → header + one blank data row (0.21.0) |
| `sheet.comment(ref, Comment.plainText("note", Some("author")))` | `Sheet` | |
| `sheet.toHtml(ref"A1:B10")` | `String` | inline-CSS HTML table |
| `sheet.usedRange` | `Option[CellRange]` | |
| `sheet.withViewSettings(view)` | `Sheet` | gridlines/zoom/tabSelected — see Sheet View & Print Setup |
| `sheet.withPageSetup(setup)` | `Sheet` | print settings — see Sheet View & Print Setup |
| `sheet.freezeAt(ref"A3")` | `Sheet` | freeze rows above + cols left; `freezeAt(anchor, scrolledTo)` sets the pane scroll target (0.13.0) |
| `sheet.withTabColor(color)` | `Sheet` | sheet tab color (rgb or theme; 0.13.0). `unfreeze` removes freeze panes |
| `sheet.withDataValidation(range, DataValidation.list("\"Yes,No\""))` | `Sheet` | list dropdown (0.13.0); also `DataValidation.listOf("Yes", "No")` and a `Vector[CellRange]` overload |
| `sheet.collapseRows(first: Row, last: Row)` / `collapseRows(rows: RowSpan)` | `Sheet` | 0.21.0 ([#465](https://github.com/TJC-LP/xl/issues/465)): hide the member rows AND mark the row after the span `collapsed` (Excel's "+" button) — the same fold as `groupRows(span, 1, collapsed = true)` / `group-rows --collapsed`; members without an outline level become a level-1 group, an existing level is kept; total (either order, no summary past the last row). Full-row spans are runtime strings: `sheet.collapseRows(orExit(RowSpan.parse("5:8")))` |
| `sheet.collapseRows(range: CellRange)` | `XLResult[Sheet]` | 0.21.0: full-row ranges only (`"5:8".asRange`); a column range or cell range is `InvalidReference`, never projected onto the row axis |
| `sheet.collapseCols(first: Column, last: Column)` / `collapseCols(cols: ColSpan)` / `collapseCols(range: CellRange): XLResult[Sheet]` | `Sheet` | 0.21.0: the column forms — `ColSpan.parse("E:H")` for a letter span (the `ref` macro takes `A1` / `A1:B2` shapes only; whole-row/column spans are runtime strings); the `CellRange` form takes full columns only |
| `sheet.expandRows(…)` / `sheet.expandCols(…)` | `Sheet` / `XLResult[Sheet]` | 0.21.0: the inverses — members unhidden, marker cleared, outline level kept; rows/columns without properties stay untouched (same overloads and axis rules) |
| `sheet.edit(edits: Edit*)` | `XLResult[Sheet]` | 0.21.0: the Edit algebra on one sheet (a one-sheet workbook under `EditScope.of(name)`), fail-fast and all-or-nothing; a target naming another sheet is `SheetNotFound`; a `RenameSheet` of this sheet is followed — see Edit algebra |

Codec types for `put`/`readTyped*`: String, Int, Long, Double, BigDecimal, Boolean, LocalDate (→ Date format), LocalDateTime (→ DateTime format), RichText.

### Records (0.21.0)

```scala
final case class Order(id: Int, customer: String, qty: Int, price: BigDecimal, note: Option[String])
  derives RowCodec                                        // fields = header names = column order

val placed = sheet.putRowsWithHeader(ref"A1", orders).unsafe // RowsPlaced(sheet, Some(A1:E1), Some(A2:E…))
placed.sheet.readRowsByHeader[Order](Row.from1(1))          // Either[RowCodecError, Vector[Order]]
placed.sheet.readRows[Order](ref"A2:E4")                    // positional, exact width
sheet.putTable(ref"A1", orders, "Orders")                   // + Excel table "Orders"
```

Field types: the nine codec types and `Option` of each (an empty cell); any other type is a compile error until you add a `given CellCodec[T]`. `Option[T]` is `None` only for a cell that is absent or `CellValue.Empty`: a cell holding the empty string — SheetJS and some exporters write `<v></v>` text cells for "blank" — decodes as `Some("")` for `Option[String]` (and is a `TypeMismatch` for `Option[Int]`), because Excel itself distinguishes `""` from blank (`ISBLANK` is FALSE, `COUNTA` counts it). Normalise with `.filter(_.nonEmpty)` or clear such cells before reading (#617). Reads see cached formula values (GH-477). Errors name the cell: `RowCodecError.Field(row, column, "qty", TypeMismatch("Int", Text("three")))` → `.message` `C2 (qty): expected Int, got Text(three)`. Full rules: `docs/reference/scripting.md` → "Records".

Typed reads see through a formula's cached value (since 0.20.0; GH-477): `Formula(expr, Some(v), kind)` decodes exactly as a plain cell holding `v` would, whatever the `FormulaKind`; `Formula(expr, None, kind)` (authored, not yet recalculated) has nothing to read and is a `TypeMismatch` whose `actual` is the formula. Do not unwrap `CellValue.Formula(_, Some(v), _)` by hand — `cell.effectiveValue` (0.20.0) is the same rule when you do need a `CellValue`. `readTypedStrict` (0.20.0) is the escape hatch that rejects every formula cell. On ≤0.19.3 every formula cell is a `TypeMismatch` whatever its cache, and `effectiveValue`/`readTypedStrict` do not exist.

## Workbook Operations

| Method | Returns | Notes |
|--------|---------|-------|
| `Workbook(sheet1, sheet2)` | `Workbook` | |
| `Workbook("Name")` | `Workbook` (literal) | one named sheet |
| `Workbook.empty` | `Workbook` | contains default "Sheet1" |
| `wb.sheets` | `Vector[Sheet]` | |
| `wb("Sales")` / `wb(SheetName)` / `wb(0)` | `XLResult[Sheet]` | 0.21.1 ([#615](https://github.com/TJC-LP/xl/issues/615)): a miss is `SheetNotFound(name, available)` carrying every sheet name, so `orExit(wb("Sumary"))` prints `did you mean: Summary`; `update`/`remove`/`rename`/`setSheetState` carry the same |
| `wb.put(sheet)` | `Workbook` | add-or-replace by name, total |
| `wb.upsert("Name", f: Sheet => Sheet)` | `Workbook` (literal) / `XLResult` (runtime) | update-or-create, total |
| `wb.update("Name", f)` | `XLResult[Workbook]` | fails if sheet absent |
| `wb.remove("Name")` | `XLResult[Workbook]` | can't remove last sheet |
| `wb.rename(old, new)` | `XLResult[Workbook]` | tab-only: does NOT rewrite `Sheet1!A1` in formulas — use `SheetRenamer.rename` |
| `SheetRenamer.rename(wb, from, to)` | `XLResult[Workbook]` | 0.20.0 ([#559](https://github.com/TJC-LP/xl/issues/559)): `wb.rename` plus the reference rewrite in every cell formula, defined name, CF and DV formula, caches preserved; `Left(FormulaError)` naming the first text that mentions `from` but cannot be parsed, workbook untouched |
| `SheetRenamer.renameLocated(wb, from, to)` | `Either[SheetRenamer.Refusal, Workbook]` | 0.21.0 ([#608](https://github.com/TJC-LP/xl/issues/608)): `rename` with WHERE it refused — `Refusal(site: Option[Site], error: XLError)`; `Site.Cell(sheet, ref)` \| `ConditionalFormat(sheet)` \| `DataValidation(sheet)` \| `Name(name)`, `site.describe` spells `Summary!I23`; `site` is `None` for `Workbook.rename`'s own `SheetNotFound`/`DuplicateSheet` |
| `wb.withCalcPr(CalcPr(iterativeCalculation = true, maxIterations = Some(100), maxChange = Some(BigDecimal("0.001"))))` | `Workbook` | author `<calcPr>` iterative calc (0.13.0); `wb.metadata.calcPr` reads it back |
| `wb.edit(edits: Edit*)` | `XLResult[Workbook]` | 0.21.0: apply `Edit`s in order under no default sheet — fail-fast, all-or-nothing, `Left(EditFailed(index, op, cause))` names the 1-based failing edit; a `None` target resolves only on a single-sheet book (else `SheetRequired`) — see Edit algebra |
| `wb.editIn(sheet: SheetName)(edits: Edit*)` | `XLResult[Workbook]` | 0.21.0: `edit` with `sheet` as the default for every unqualified target (the CLI's `-s`) |

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
ref"A1".comment(Comment.plainText("provenance", Some("me")))          // Patch.SetComment (0.13.0)
ref"A1:C10".conditionalFormat(CfRule.cellIs(CfOperator.GreaterThan, "100", dxf))  // Patch.SetConditionalFormat (0.13.0)
ref"B2:D6".outlined(BorderStyle.Medium)                       // box the range's outer edges
ref"B2:D6".outlined(BorderStyle.Thin, Color.fromRgb(128, 128, 128))  // colored outline
Patch.empty                     // identity — fold seed
```

`outlined` is edge-correct (top row gets top, corners get both, interior untouched; 1×1 gets all
four sides) and desugars to per-cell `Patch.MergeBorder(ref, border)` — an apply-time border
overlay that preserves each cell's font, fill, numFmt, and untouched border sides. Cost is
proportional to the perimeter; prefer bounded ranges over whole columns.

Runtime `RefType` (from `ref"$s"`) supports the same `:=` / `.styled` / `.merge` / `.remove`; `:=` on a parsed range fills it.

## Edit algebra (0.21.0)

`Edit` is the operation vocabulary behind `xl batch` and every mutating CLI verb, as values — one case per op (49) — applied by one interpreter (`Edit.applyAll`) behind `wb.edit` / `wb.editIn` / `sheet.edit`. At 0.21.0 the CLI keeps its own batch path (lowering it onto `Edit.applyAll` is [#583](https://github.com/TJC-LP/xl/issues/583)) but calls the same library kernels (`Sheet.fill`/`copyRange`/`sort`/`clearRange`/`groupRows`/`autoFit`, `StructuralEditor`, `SheetRenamer`). Targets carry their sheet: `Loc(sheet: Option[SheetName], ref)`, `Area(sheet, range)`, or a `sheet: Option[SheetName]` field beside a `RowSpan`/`ColSpan`; `None` is the scope's default (THE sheet rule: qualifier > default > the only sheet of a single-sheet book > `SheetRequired`).

| Call | Returns | Notes |
|------|---------|-------|
| `wb.edit(edits*)` | `XLResult[Workbook]` | no default sheet — qualify with `Some(sheet)` or edit a single-sheet book |
| `wb.editIn(sheet)(edits*)` | `XLResult[Workbook]` | `sheet` is the default for every `None` target (the CLI's `-s`) |
| `sheet.edit(edits*)` | `XLResult[Sheet]` | one-sheet scope; another sheet's target is `SheetNotFound`; a `RenameSheet` of this sheet is followed |
| `Edit.applyAll(wb, edits: Vector[Edit], scope: EditScope)` | `XLResult[Applied]` | `Applied(workbook, planned, scope)`: the edited book, one `Planned(index, edit, sheet, touched)` per edit, the scope after the sequence; `applied.touchedBySheet`, `applied.structural` |
| `Edit.plan(wb, edits, scope)` | `XLResult[Vector[Planned]]` | the same fold, workbook discarded — a *semantic* dry-run (fails exactly where `applyAll` fails, costs the same) |
| `Edit.validate(edit)` | `XLResult[Unit]` | the *static*, workbook-free check of one edit: value counts match the area, spans/levels/widths in range, formula syntax, `MoveSheet` names exactly one destination; 0.21.1 ([#617](https://github.com/TJC-LP/xl/issues/617)): a refusal is `XLError.InvalidArgument(op, reason)` — code `INVALID_ARGUMENT`, message `<op>: <reason>` — never `XLError.Other`, and `EditFailed` supplies `opIndex` |
| `Edit.lower(edit, sheet)` / `Patch.toEdits(patch, sheet)` | `Option[Patch]` / `Option[Vector[Edit]]` | the bridge to `Patch`; `None` when formula support, another sheet or the workbook is needed |
| `Edit.put(loc, a)` / `Edit.put(ref, a)` | `Edit` | THE way to build a `Put`: lifts the codec's format as `FormatHint.Inferred` so a `LocalDate`/`BigDecimal` keeps its format, exactly like `sheet.put` |
| `EditScope.none` / `EditScope.of(sheet)` / `scope.after(edit)` | `EditScope` | the scope type (`ops.Scope`, renamed on export); a `RenameSheet` of the default sheet retargets what follows |
| `EditSchema.all` / `EditSchema.find(name)` / `EditSchema.nameOf(edit)` | `Vector[EditSpec]` / `Option[EditSpec]` / `String` | the algebra's metadata, one row per case (49; each names its `batchOp`/`cliVerb`); `nameOf(Edit.DragFormula(…))` is `"drag-formula"`. Not the document `xl batch --schema` prints (xl-cli's `OpRegistry`, 32 ops) |

**Semantics.** Edits apply in order, fail-fast and all-or-nothing: the first failure is `Left(XLError.EditFailed(index, op, cause))` — `index` 1-based (`err.opIndex`), `op` the kebab name, `cause` the real error (`err.root` is the innermost cause; `err.code`/`hint`/`candidates` are the cause's) — and the input workbook is never partially written. `err.message` is `op 2 (drag-formula): …` — the same 1-based position `xl batch` reports as `location.opIndex` on a `BATCH_OP_FAILED`; `orExit` prints it as the CLI's stderr diagnostic. Formula-aware cases (`DragFormula`, `Fill`, `Copy`, the structural four, `RenameSheet`) use the prelude's `given FormulaSupport` (xl-evaluator's); `FormulaSupport.textOnly` refuses them with `UnsupportedCapability`.

| Group | Cases |
|-------|-------|
| Cell content | `Put(at: Loc, value: CellValue, format: Option[FormatHint])`, `PutValues(at: Area, values: Vector[CellValue], format)` (row-major), `PutFormula(at: Loc, formula: String, format)`, `PutFormulas(at: Area, formulas: Vector[String], format)`, `DragFormula(at: Area, formula, anchor: ARef, format)`, `Fill(source: Area, target: CellRange, direction: Edit.FillDir)` (`Down`/`Right`), `Copy(source: Area, target: Loc, valuesOnly: Boolean)` (a single-cell target; batch `copy` also takes a range), `Sort(at: Area, keys: Vector[Edit.SortKeySpec], hasHeader)` (`SortKeySpec.ascending(col)`/`.descending(col)`), `Clear(at: Area, what: ClearWhat)` |
| Style & layout | `Style(at: Area, overlay: StyleOverlay, mode: StyleMode)`, `Merge(at)`, `Unmerge(at)`, `ColWidth(sheet, cols: ColSpan, width)`, `RowHeight(sheet, rows: RowSpan, height)`, `HideCols`/`ShowCols(sheet, cols)`, `HideRows`/`ShowRows(sheet, rows)`, `AutoFit(sheet, cols: Option[ColSpan])` (0.21.1, [#613](https://github.com/TJC-LP/xl/issues/613): a column is fitted to what it DISPLAYS — a formula cell by its cached value, an uncached formula contributing nothing — so fit after `recalculate`, or let `Excel.writeChecked` supply the caches and accept the values' widths) |
| Outline | `GroupRows(sheet, rows, level, collapsed)`, `GroupCols(sheet, cols, level, collapsed)`, `UngroupRows(sheet, rows)`, `UngroupCols(sheet, cols)` |
| Annotations & objects | `SetComment(at: Loc, comment)`, `RemoveComment(at)`, `Hyperlink(at, target: Option[String])`, `AddConditionalFormat(sheet, ranges, rules)`, `AddChart(sheet, chart, anchor)`, `AddImage(sheet, image, anchor)` |
| Sheet view & print | `Freeze(at: Loc)`, `Unfreeze(sheet)`, `SetSheetView(sheet, gridlines, zoom, tabSelected)`, `SetTabColor(sheet, color)`, `SetAutoFilter(sheet, range)`, `SetPageSetup(sheet, orientation, scale, fitToWidth, fitToHeight, fitToPage)`, `SetHeaderFooter(sheet, oddHeader, oddFooter, evenHeader, evenFooter, firstHeader, firstFooter, differentOddEven, differentFirst)` |
| Structure | `InsertRows(sheet, at: Row, count)`, `DeleteRows(sheet, at: Row, count)`, `InsertCols(sheet, at: Column, count)`, `DeleteCols(sheet, at: Column, count)` — references rewritten on every sheet, `#REF!` on loss |
| Workbook | `AddSheet(name, after, before)`, `RemoveSheet(name)`, `RenameSheet(from, to)` (rewrites references), `MoveSheet(name, toIndex, after, before)` (exactly one; `toIndex` = the FINAL 0-based position — `xl move-sheet --to N` still uses the pre-removal, clamped index, [#583](https://github.com/TJC-LP/xl/issues/583)), `CopySheet(source, target)`, `HideSheet(name, veryHidden)`, `ShowSheet(name)`, `DefineName(name, refersTo, scope)`, `RemoveName(name, scope)` |

```scala
val Data = SheetName.unsafe("Data")
val edited: XLResult[Workbook] = wb.editIn(Data)(
  Edit.put(ref"D1", "Total"),
  Edit.DragFormula(Area(None, ref"D2:D9"), "=B2*C2", ref"D2", None),
  Edit.PutFormula(Loc(None, ref"D10"), "=SUM(D2:D9)", Some(FormatHint.Explicit(NumFmt.Currency))),
  Edit.Style(Area(None, ref"A1:D1"), StyleOverlay(bold = Some(true)), StyleMode.Merge),
  Edit.DeleteRows(None, Row.from1(5), 2),                     // rows 5-6; formulas on every sheet rewritten
  Edit.AddSheet(SheetName.unsafe("Summary"), after = Some(Data), before = None)
)
Excel.writeChecked(orExit(edited), "out.xlsx")                // Left → the CLI's diagnostic on stderr, exit 1
```

Full prose, two runnable scripts and the design rationale: `docs/reference/scripting.md` → "The Edit algebra".

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
  .rotated(90)                      // text rotation (Excel-UI degrees: -90..90, 255 = stacked; 0.13.0)
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
// Pictures & charts (0.12.0+) — drawing layer with hybrid preservation
val image = ImageData.detect(bytes).unsafe                     // XLResult: format from magic bytes
sheet.addImage(image, ref"B2").unsafe                          // natural size (XLResult; PNG/JPEG sniffed)
sheet.addImage(image, ref"B2:D8")                              // stretched over a range — total
val chart = Chart.bar(series, title = Some("Revenue")).unsafe  // XLResult: validated; also Chart.line/pie
sheet.addChart(chart, ref"F2:K16")                             // anchored over the range — total
val branded = firstSeries.copy(fill = Some(Color.Rgb(0xFF307FE2))) // explicit series color (0.15.0, packed
                                                               // ARGB); fill = None cycles theme accents (LO-visible)
sheet.pictures; sheet.charts                                   // typed views (unmodeled drawings preserved)

// Conditional formatting (0.12.1+) — typed rules with dxf differential formats
sheet.conditionalFormat(ref"A1:A10", CfRule.cellIs(CfOperator.GreaterThan, "100", Dxf(fill = Some(Fill.Solid(Color.Rgb(0xFFFFC7CE))))))
sheet.conditionalFormat(
  ref"B1:B10",
  CfRule.colorScale3(
    CfPoint(Cfvo.Min, Color.Rgb(0xFFF8696B)),
    CfPoint(Cfvo.Percentile(50), Color.Rgb(0xFFFFEB84)),
    CfPoint(Cfvo.Max, Color.Rgb(0xFF63BE7B))
  )
)
// also CfRule.expression / dataBar / top10 / text ops; unmodeled rule families preserved byte-faithfully

sheet.withViewSettings(SheetView(showGridLines = false, zoomScale = Some(90), tabSelected = Some(true)))
// tabSelected (0.13.0): tab SELECTION/grouping, not the active sheet (that is Workbook.activeSheetIndex)

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
| `wb.recalculate(IterativeCalc(maxIter, maxChange))` | `RecalcResult` | 0.13.0: opt-in Jacobi fixpoint for declared cycles (also `(clock, iterative)` / `(clock, rng, iterative)`); `IterativeCalc.fromCalcPr(cp)` bridges a file's `<calcPr>` |
| `wb.withCachedFormulas([clock])` | `Workbook` | = `recalculate(clock).workbook` |
| `sheet.evaluateFormula(formula[, clock][, rng][, workbook])` | `XLResult[CellValue]` | pass `Some(wb)` iff formula references other sheets |
| `sheet.putFormulaInheriting(ref, formula[, workbook])` | `XLResult[Sheet]` | 0.11.2+: puts the formula AND inherits the referenced cells' number format into a General target (Excel's entry behavior) |
| `sheet.evaluateWithDependencyCheck([clock])` | `XLResult[Map[ARef, CellValue]]` | fail-fast on first error/cycle |
| `FormulaParser.parse("=A1+B1")` | `Either[ParseError, TExpr[?]]` | AST access; 0.21.0: range nodes carry a `RangeForm` and error literals parse to `TExpr.ErrorLit` — a **Breaking** pattern-arity change for code matching range nodes, see `TExpr` in Key Types |
| `DependencyGraph.fromSheet(sheet)` | `DependencyGraph` | `.precedents(ref)` / `.dependents(ref)` |
| `Clock.system` / `Clock.fixedDate(LocalDate)` | `Clock` | deterministic TODAY()/NOW() |

`RecalcResult`: `.workbook` (successful formulas cached), `.evaluated: Map[SheetName, Map[ARef, CellValue]]`, `.errors: Vector[CellEvalError]`, `.isClean`, `.toEither`.

**Formula syntax notes (0.13.0)**: percent postfix `=A1*10%` / `=10%` / `=(1+5%)^2` parses, evaluates (`10%` → exact `0.1`), broadcasts over ranges, and prints byte-identically (never rewritten to `/100`). Leading unary plus is preserved through print (`=+A1`, `=++A1`, `=2^+2` round-trip; evaluation is pure identity). Defined names resolve — `=IF(case=2,…)`, `=entry_mult*ltm_ebitda`, `=SUM(rev_range)` evaluate against workbook- and sheet-scoped names (sheet-scoped shadows global) and contribute dependency edges; unresolvable names are clean per-cell errors.

### All 112 functions

SUM, SUMIF, SUMIFS, SUMPRODUCT, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, AVERAGE, AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS, MEDIAN, STDEV, STDEVP, VAR, VARP, LARGE, SMALL, RANK, PERCENTILE, QUARTILE, MIN, MAX, IF, IFS, IFERROR, SWITCH, CHOOSE, AND, OR, NOT, ISNUMBER, ISTEXT, ISBLANK, ISERR, ISERROR, N, CONCATENATE, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, FIND, SEARCH, SUBSTITUTE, TEXT, VALUE, TODAY, NOW, DATE, YEAR, MONTH, DAY, EOMONTH, EDATE, DATEDIF, NETWORKDAYS, WORKDAY, YEARFRAC, ABS, ROUND, ROUNDUP, ROUNDDOWN, MROUND, INT, MOD, POWER, SQRT, LOG, LN, EXP, FLOOR, CEILING, TRUNC, SIGN, PMT, FV, PV, RATE, NPER, NPV, IRR, XNPV, XIRR, VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH, OFFSET, INDIRECT, HYPERLINK, PI, ROW, COLUMN, ROWS, COLUMNS, ADDRESS, CELL, TRANSPOSE, SEQUENCE, SORT, UNIQUE, FILTER, RAND, RANDBETWEEN — plus LET (lexical bindings; a parser-level special form, not in the registry listing)

## IO

### Sync facade (`Excel`) — for scripts

| Method | Notes |
|--------|-------|
| `Excel.read(path: String): Workbook` | throws `XLException`/IO errors at this edge |
| `Excel.readSheet(path, name: String): XLResult[Sheet]` | 0.21.1: one sheet by exact name — `Excel.read` plus the lookup, so the whole workbook is loaded (stream one sheet of a large file with `ExcelIO.readSheetStream`); a missing name is `Left(SheetNotFound(name, available))` — the message lists the available sheets, `.candidates` the nearest ("did you mean") — and a missing/corrupt/over-limit file `Left` of the reader's `IOError`/`ParseError`/`SecurityError`; unwrap with `orExit(...)` or `.unsafe`. 0.21.0 returned `Sheet` and threw `XLException` |
| `Excel.readMetadata(path): LightMetadata` | 0.21.0: `sheets` (`SheetInfo`: name, sheetId, state, dimension), `definedNames`, `date1904` — no cells loaded, instant on any file size, the inflated parts held to `Excel.read`'s ZIP-bomb limits; `XLException` on an unreadable or over-limit package |
| `Excel.write(wb, path: String): Unit` | also accepts `XLResult[Workbook]`; writes formulas with whatever cache they carry — never for a freshly built model |
| `Excel.writeChecked(wb, path: String): RecalcResult` | 0.21.0: compute ONLY the uncached formulas (`recalculateUncached`), write, return the result — every existing cache is written byte for byte; overloads take `RecalcOptions` or accept `XLResult[Workbook]` |
| `Excel.writeRecalculated(wb, path: String): RecalcResult` | 0.13.0: recalc EVERY formula + write (even on partial failure) + return the result; overloads add `Clock`, `Clock`+`Rng`, `RecalcOptions` (0.21.0), or accept `XLResult[Workbook]` |
| `Excel.modify(path)(f: Workbook => Workbook): Unit` | atomic in-place replacement |
| `Excel.modifyR(path)(f: Workbook => XLResult[Workbook]): Unit` | 0.21.0: the same replacement for a fallible transform; `Left` throws `XLException` BEFORE any write (file byte-identical, no scratch file) |

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
sheet.displayCell(ref"A1")          // DisplayWrapper — the NumFmt-formatted text, NOT a String
sheet.displayCell(ref"A1").toString // String ("$1,234.56"); .formatted is the same field
s"A1 = ${sheet.displayCell(ref"A1")}"            // interpolation renders it via toString
sheet.displayCell(ref"A1").toString.padTo(12, ' ') // String methods need the unwrap first (.padTo on the wrapper does not compile)
sheet.displayFormula(ref"A1")       // String: the formula text ("=SUM(A1:A3)"), or the formatted value for a non-formula cell
```

## Errors

```scala
XLResult[A]                  // = Either[XLError, A]
err.message                  // human-readable
err.code; err.hint; err.candidates   // 0.20.0: stable SCREAMING_SNAKE code, next step, "did you mean"
result.unsafe                // throws XLException(err) — the one sanctioned unwrap (prelude)
result.getOrElse(fallback)
err.opIndex; err.root        // 0.21.0: Some(1-based index) of the failing edit in an EditFailed(index, op, cause), and the innermost cause (code/hint/candidates are the cause's)
orExit(result)               // 0.21.0: the value, or print exitMessage(err) to stderr and exit 1
exitMessage(err)             // 0.21.0: err.renderDiagnostic — "Error: …\n  code: …" then "  did you mean: …" / "  hint: …" when present; the CLI's exact stderr diagnostic
```
