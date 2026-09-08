---
name: xl-scripting
description: "Write type-safe Scala scripts against the xl library (com.tjclp::xl, via scala-cli) for complex Excel/.xlsx work: bulk or conditional transformations, multi-file pipelines, typed data extraction into Scala types, formula-heavy model building, and streaming 100k+ row files. Prefer over the xl-cli skill when a task needs loops, intermediate computation, or many dependent edits that would round-trip the file repeatedly; use xl-cli for quick reads, single edits, and visual exports."
---

# XL Scripting - Type-Safe Excel in Scala

xl is a purely functional Excel library for Scala 3: compile-time validated cell references, total APIs, and structured errors (`Either`, never exceptions). A script is a `.sc` file run with `scala-cli` — no project setup, no JDK install (scala-cli provisions one).

## When to Use This Skill (vs xl-cli)

| Task | Use |
|------|-----|
| Quick look at a sheet, single cell edit, search | `xl-cli` |
| Export to PNG/PDF/HTML, visual verification | `xl-cli` |
| Bulk generation (100s of cells from data) | **xl-scripting** |
| Loops, conditionals, or computation over cell values | **xl-scripting** |
| Multi-file pipelines (merge N workbooks, batch convert) | **xl-scripting** |
| Typed extraction into Scala values for further logic | **xl-scripting** |
| Build a formula model + recalculate + inspect failures | **xl-scripting** |
| Streaming filters/aggregates over 100k+ rows | **xl-scripting** |

The stateless CLI re-reads and re-writes the file on every invocation; a script holds the workbook in memory across the whole transformation. The two compose well: **generate with a script, verify visually with the CLI** (`xl -f out.xlsx -s Sheet1 view A1:F20 --format png`).

## Setup

Check: `which scala-cli || echo "not installed"`

**macOS:** `brew install Virtuslab/scala-cli/scala-cli`
**Linux:** `curl -sSLf https://scala-cli.virtuslab.org/get | sh`
**Windows:** `winget install virtuslab.scalacli`

No JDK prerequisite — scala-cli auto-provisions a JVM via coursier. The first run downloads dependencies (~30-60s); subsequent runs are cached.

## Script Skeleton & Version

The canonical header for every script (this is the single source of truth — recipes in `reference/RECIPES.md` use the identical header):

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.0

import com.tjclp.xl.scripting.{*, given}

val sheet = Sheet("Demo").put(ref"A1", "Hello").put(ref"B1", 42)
Excel.write(Workbook(sheet), "/tmp/demo.xlsx")
println(s"wrote ${sheet.cells.size} cells")
```

- **One import.** `com.tjclp.xl.scripting.{*, given}` bundles the core API, DSL operators, compile-time literals, formula evaluation, sync `Excel` IO, streaming `ExcelIO`, smart value detection, and the `.unsafe` boundary. Never combine it with `import com.tjclp.xl.{*, given}` in the same file (ambiguous forwarders).
- `.sc` files run **top-level statements** — no `@main`, no object wrapper. Run with `scala-cli run script.sc`.
- For multi-script pipelines, put the two `//> using` directives in a shared `project.scala` and start each script with `//> using file project.scala`.
- To use a release newer than this skill documents: `curl -s https://api.github.com/repos/TJC-LP/xl/releases/latest | grep '"tag_name"' | cut -d'"' -f4` and bump the dep line. Maven artifacts are immutable, so the pinned version above always resolves.

## Quick Reference

```scala
// Create / read / write (sync facade; throws only at this IO edge)
val wb  = Excel.read("in.xlsx")                    // Workbook
Excel.write(wb, "out.xlsx")                        // also accepts XLResult[Workbook]; NEVER for a freshly built model — see writeChecked
Excel.modify("file.xlsx")(_.upsert("Log", identity)) // atomic in-place read→transform→write
Excel.modifyR("file.xlsx")(_.update("Log", f))     // 0.21.0: XLResult-returning transform; a Left throws BEFORE any write
Excel.readSheet("in.xlsx", "Summary")              // 0.21.0: Sheet — loads the WHOLE workbook, then picks one; a typo throws SheetNotFound(name, available) with "did you mean" candidates
Excel.readMetadata("in.xlsx")                      // 0.21.0: LightMetadata (sheet names/dimensions/defined names), no cells loaded, ZIP-bomb guarded

// Sheets in a workbook
wb.sheets                                          // Vector[Sheet]
wb("Sales")                                        // XLResult[Sheet] (literal name compile-checked)
wb.upsert("Summary", _.put(ref"A1", "Total"))      // total: update-or-create, returns Workbook
wb.update("Sales", f)                              // XLResult[Workbook] (errors if absent)

// Sheet creation: Sheet("lit") is compile-checked and returns Sheet; for RUNTIME names
Sheet.named(dynamicName)                           // XLResult[Sheet] — the dynamic-name factory (0.18.0)
Workbook.named("Data", "Summary")                  // 0.20.0: XLResult[Workbook]; DuplicateSheet on a repeat

// Cell writes (literal refs are compile-time validated and infallible)
sheet.put(ref"A1", "Title")                        // Sheet
sheet.put("A1", "Title")                           // string LITERAL: compile-checked, returns Sheet
sheet.putAt(cellStr, "Title")                      // 0.20.0: RUNTIME string → XLResult[Sheet] (also styleAt/mergeAt/commentAt)
sheet.put(ref"B1", 42)                             // Int/Long/Double/BigDecimal/Boolean/LocalDate(Time)/RichText
sheet.put(ref"C1", "$1,234.56".toFormatted)        // smart detection → Currency format
sheet.style(ref"A1:D1", CellStyle.default.bold)    // Sheet
sheet.collapseRows(Row.from1(5), Row.from1(8))     // 0.21.0: hide 5:8 + mark row 9 collapsed (level-1 group if ungrouped); also takes a RowSpan
sheet.collapseCols(orExit(ColSpan.parse("E:H")))   // 0.21.0: column form; spans are runtime strings (ref"" takes A1/A1:B2 only) and carry their axis
sheet.collapseRows(orExit("5:8".asRange))          // 0.21.0: CellRange form is XLResult[Sheet]: full rows only — "E:H" here is InvalidReference, not 1M hidden rows

// Patch DSL (compose pure values, apply once)
val patch = (ref"A1" := "Report") ++ ref"A1:C1".merge ++ ref"A1".styled(CellStyle.default.bold)
sheet.put(patch)                                   // Sheet
ref"E2:E9" := 0                                    // range fill: every cell (Ctrl+Enter semantics)
ref"A2".down(3).right(1)                           // total navigation → B5 (unchecked at the grid edge)
ref"A2".tryDown(3)                                 // 0.20.0: bounded → Some(A5); None past the edge; clampShift pins
ref"A1:D10".rows                                   // 0.20.0: lazy one-row slices; row(i)/column(i) are Option

// Edit algebra (0.21.0): the batch/CLI operation vocabulary as values; one interpreter (Edit.applyAll) over the kernels the CLI verbs call
wb.editIn(sales)(edits*)                           // XLResult[Workbook]; `sales` is the sheet for every target with sheet = None (the CLI's -s)
wb.edit(Edit.put(Loc(Some(sales), ref"A1"), 1))    // no default sheet: qualify with Some(...), or a single-sheet book — else SheetRequired
sheet.edit(Edit.Merge(Area(None, ref"A1:C1")))     // XLResult[Sheet]; fail-fast, all-or-nothing
Edit.DragFormula(Area(None, ref"D2:D9"), "=B2*C2", ref"D2", None) // fill-down; also Fill/Copy/Sort/Clear/InsertRows/DeleteRows/AddSheet/RenameSheet/...
Edit.plan(wb, edits, EditScope.of(sales))          // XLResult[Vector[Planned]]: semantic dry-run; Edit.validate(e) is the static check
err.opIndex                                        // Some(n): 1-based position of the failing edit (EditFailed); err.code/hint are the cause's

// Typed reads (since 0.20.0 a formula cell reads as its cached value — GH-477)
sheet.readTyped[BigDecimal](ref"C1")               // Either[CodecError, Option[BigDecimal]]
sheet.readTypedOr[Int](ref"B1", 0)                 // total, with default
sheet.readTypedOpt[LocalDate](ref"D1")             // flat Option
sheet.readTypedStrict[BigDecimal](ref"C1")         // 0.20.0: any formula cell → Left(TypeMismatch), cached or not

// Formulas
sheet.put(ref"D2", fx"=B2*C2")                     // compile-time validated literal
wb.evaluateFormula("=SUM(Sales!A1:A9)", "Summary") // XLResult[CellValue], cross-sheet aware
val r = wb.recalculate()                           // RecalcResult: total, per-cell errors
r.isClean; r.errors.map(_.render); r.workbook      // inspect, then write r.workbook
Excel.writeChecked(wb, "out.xlsx")                 // 0.21.0: cache ONLY the uncached formulas + write + RecalcResult — THE write for a built model
Excel.writeRecalculated(wb, "out.xlsx")            // 0.13.0: recompute EVERY formula + write + RecalcResult; both take a RecalcOptions (0.21.0)

// Errors: XLResult[A] = Either[XLError, A]; unwrap ONCE at the edge
wb.update("Sales", f).unsafe                       // throws structured XLException if Left
orExit(wb.update("Sales", f))                      // 0.21.0: or print "Error: …" + indented "code:/did you mean:/hint:" (the CLI's exact stderr diagnostic), exit 1
```

## Essential Patterns

### Read → modify → write

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.0
import com.tjclp.xl.scripting.{*, given}

val wb = Excel.read("input.xlsx")
val updated = wb
  .upsert("Audit", _.put(ref"A1", "reviewed"))         // total: creates sheet if missing
  .update("Data", _.put(ref"B2", 99))                  // XLResult: Data must exist
  .unsafe                                              // one unwrap, at the edge
Excel.write(updated, "output.xlsx")
```

`Excel.modify("file.xlsx")(f)` does the same in place with atomic file replacement. Since 0.21.0 `Excel.modifyR("file.xlsx")(f)` takes an `XLResult`-returning transform — `_.update("Data", …)` needs no `.unsafe` inside the lambda, and a `Left` throws *before* anything is written, leaving the file byte-identical. `Excel.readSheet(path, name)` (0.21.0) is `Excel.read` plus the lookup — the whole workbook is loaded, then one sheet is selected — throwing an `XLException(SheetNotFound(name, available))` on a typo whose message lists every available sheet and whose `candidates` name the nearest (`orExit` prints them as `did you mean:`); `Excel.readMetadata(path)` (0.21.0) lists sheets, dimensions and defined names without loading a cell — decide what to read (or stream) before reading it.

### Compile-time literals vs runtime refs

Literal refs/formulas are validated **at compile time** — a typo fails the build, not the workbook:

```scala
val a = ref"A1"               // ARef
val rng = ref"A1:B10"         // CellRange
val f = fx"=SUM(A1:B10)"      // CellValue.Formula (parens/syntax checked)
val m = money"$$1,234.56"     // Formatted(Number, Currency) — note $$ escapes $ in interpolators
```

Runtime interpolation returns `Either` because validation must happen at runtime:

```scala
val row = 5
val cellE = ref"A$row"                    // Either[XLError, RefType]
val formE = fx"=B$row*C$row"              // Either[XLError, CellValue]
// sequence with for-comprehensions, or unwrap explicitly:
val patch = (ref"A$row").map(_ := "x").getOrElse(Patch.empty)
val cell2 = fx"=B$row*2".unsafe           // explicit boundary
```

**The same split applies to every string-taking form** — `Sheet(name)`, `Workbook(name, …)`,
`sheet.put("A1", v)`, `sheet.style("A1:D1", st)`, `sheet.merge("A1:C1")`, `sheet.comment("A1", c)` —
with no `$` at the call site to warn you. They are `transparent inline`: a string literal validates
at compile time and returns `Sheet`/`Workbook`; the very same call with a `val` returns `XLResult[…]`,
so a chained `.put(...)` type-errors. Two rules:

1. **Literals are total** — use the literal forms (and `ref"…"`) whenever the address is known when
   you write the script.
2. **Computed strings use the explicit runtime twins**, which spell `XLResult` in their signatures:
   `Sheet.named` (0.18.0) and, since 0.20.0, `Workbook.named`, `sheet.putAt`, `sheet.styleAt`,
   `sheet.mergeAt`, `sheet.commentAt`. Same validation and the same value/style path as the
   **literal** forms (`putAt` reaches the very code the literal `put` expands to). Do not pass a
   computed string to the transparent `put`/`style`/`merge`/`comment` — it compiles, but the return
   type flips. **Corner forms only**: the twins take `A1` and `A1:B2` — no `A:A`/`1:1`, no `$`
   anchors, and `mergeAt` needs two corners (the *dynamic* transparent `merge`/`style` accept those
   via `CellRange.parse`); for such spellings parse first: `s.asRange.map(sheet.merge)`.

```scala
val lit = Sheet("Acquisitions").put("A1", 1)      // : Sheet (both literals, compile-time validated)
val nm: String = config.sheetName
val cell: String = s"B${row + 1}"
val dyn = Sheet(nm)                               // : XLResult[Sheet] — return type changed silently
val dyn2 = lit.put(cell, 42)                      // : XLResult[Sheet] — same trap on put

// 0.20.0: spelled-out twins — flatMap the chain, unwrap once at the edge
val ok: XLResult[Sheet] =
  for
    s <- Sheet.named(nm)                            // Left(InvalidSheetName) on a bad name
    a <- s.putAt(cell, 42)                          // Left(InvalidCellRef) on a bad ref; a range is rejected
    b <- a.putAt("C2", BigDecimal("2.5"), currency) // styled twin: same codec-format merge as the literal
    c <- b.styleAt("A1:C1", header)                 // cell → that cell; range → every cell in it
    d <- c.mergeAt("A1:C1")                         // Left(InvalidRange) for a single cell, A:A, $-anchors, garbage
  yield d
val wb: XLResult[Workbook] = Workbook.named("Data", "Summary") // Left(DuplicateSheet) on a repeat

// Corner forms only: for "D:D", "1:1", "$A$1:C3" (or a one-cell merge) parse first, then use the typed overloads
val col: String = "D"
s"$col:$col".asRange.map(sheet.merge)            // XLResult[Sheet]; String.asRange is CellRange.parse-backed
s"$col:$col".asRange.map(r => sheet.style(r, header))
cell.asCell.map(r => sheet.put(r, 42))           // String.asCell: A1 cells (no $ anchors — use asRange)
```

A sheet-qualified string (`"Sales!A1"`) is refused by every twin with `InvalidReference` — qualify
at the workbook instead: `wb.update(sheetName, _.putAt("A1", v))`. On ≤0.19.x (twins absent),
ascribe the union explicitly (`val s: XLResult[Sheet] = sheet.put(cell, 42)`) or parse once with
`s.asCell`/`s.asRange` (or `RefType.parse`) and use the typed `ARef`/`CellRange` overloads.

**Prefer total navigation over interpolated refs in loops** — no Either at all:

```scala
val base = ref"A2"
val r3 = base.down(2)         // A4
val c2 = base.right(1)        // B2
val shifted = base.shift(1, 2) // B4
```

`shift`/`down`/`up`/`left`/`right` are unchecked at the grid edge (`ref"A1".up()` mints "A0").
Since 0.20.0 the **bounded** forms return `Option` or clamp, so a loop stops cleanly instead:

```scala
// 0.20.0
ref"A1".tryDown(1)                // Some(A2)
ref"A1".tryShift(-1, 0)           // None — would be column -1
ref"XFD1".tryRight(1)             // None — past the last column (XFD1048576 is the corner)
ref"C3".clampShift(-10, 5)        // A8  — column pinned to A, row shifted
val cellPatch = ref"A1".tryDown(2).fold(Patch.empty)(_ := "third row") // no Either in the loop body

val table = ref"A1:D10"           // range slices instead of interpolated corners
table.rows.size                   // 10 one-row-high CellRanges (lazy), top to bottom
table.row(0)                      // Some(A1:D1); row(10) is None (0-based within the range)
table.columns.map(_.toA1).toList  // List("A1:A10", "B1:B10", "C1:C10", "D1:D10"); column(i) likewise
```

### Bulk generation: fold data into a Patch

Patches are pure values forming a monoid (`++`). Build the whole change set, apply once:

```scala
val products = List(("Widget", 150, 19.99), ("Gadget", 75, 29.99))

val rows = products.zipWithIndex.foldLeft(Patch.empty) { case (acc, ((name, units, price), i)) =>
  val r = ref"A2".down(i)
  acc ++ (r := name) ++ (r.right(1) := units) ++ (r.right(2) := price) ++
    (r.right(3) := fx"=B${i + 2}*C${i + 2}".unsafe)
}
val sheet = Sheet("Sales").put(rows)
```

`range := value` fills every cell in the range (Excel Ctrl+Enter): `ref"E2:E100" := 0`.

### Batch-shaped edits: the `Edit` algebra (0.21.0)

A `Patch` is sheet-local and formula-blind. `Edit` is the operation vocabulary behind `xl batch` and every mutating verb — one case per op, 49 in all — applied by one interpreter, `Edit.applyAll`, behind `wb.edit` / `wb.editIn(sheet)` / `sheet.edit`. At 0.21.0 the CLI still runs its own batch path (`OpRegistry`/`BatchParser`; lowering it onto `Edit.applyAll` is [#583](https://github.com/TJC-LP/xl/issues/583)), but both call the same library kernels (`Sheet.fill`/`copyRange`/`sort`/`clearRange`/`groupRows`/`autoFit`, `StructuralEditor`, `SheetRenamer`), so a script and a verb make the same change. Reach for it when a script needs what a `Patch` cannot say: formula dragging (`DragFormula`, `Fill`, `Copy`), structural edits (`InsertRows`/`DeleteRows`/`InsertCols`/`DeleteCols` — references rewritten on every sheet, `#REF!` on loss), sheet management (`AddSheet`, `RenameSheet` with reference rewriting, `MoveSheet`, `CopySheet`, `HideSheet`/`ShowSheet`), outline groups, comments/hyperlinks, CF/charts/images, view and print setup, defined names — beside the plain `Put`/`PutValues`/`PutFormula`/`PutFormulas`, `Style`, `Merge`/`Unmerge`, `Clear`, `Sort`, widths/heights, hide/show and `AutoFit`.

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.0
import com.tjclp.xl.scripting.{*, given}

val Data = SheetName.unsafe("Data")
val book = Workbook(
  Sheet(Data)
    .put(ref"A1", "Item").put(ref"B1", "Qty").put(ref"C1", "Price")
    .put(ref"A2", "Widget").put(ref"B2", 3).put(ref"C2", BigDecimal("19.99"))
    .put(ref"A3", "Gadget").put(ref"B3", 5).put(ref"C3", BigDecimal("29.99"))
)

val edited = orExit(
  book.editIn(Data)(                                                     // Data is the sheet for every `None` target
    Edit.put(ref"D1", "Total"),                                          // codec format inference kept
    Edit.DragFormula(Area(None, ref"D2:D3"), "=B2*C2", ref"D2", None),   // D3: =B3*C3
    Edit.PutFormula(Loc(None, ref"D4"), "=SUM(D2:D3)", Some(FormatHint.Explicit(NumFmt.Currency))),
    Edit.Style(Area(None, ref"A1:D1"), StyleOverlay(bold = Some(true)), StyleMode.Merge),
    Edit.InsertRows(None, Row.from1(1), 1),                              // every formula shifts: D5 = SUM(D3:D4)
    Edit.AddSheet(SheetName.unsafe("Summary"), after = Some(Data), before = None),
    Edit.PutFormula(Loc(Some(SheetName.unsafe("Summary")), ref"B2"), "=Data!D5", None) // qualified target
  )
) // Left → "Error: op N (<op>): …" + code/hint on stderr (the CLI's diagnostic), exit 1
val result = Excel.writeChecked(edited, "/tmp/edited.xlsx")              // caches the new formulas, writes
println(s"clean: ${result.isClean}")
```

- **Targets carry their sheet**: `Loc(sheet: Option[SheetName], ref)`, `Area(sheet, range)`, and `sheet: Option[SheetName]` on the row/column and sheet-level cases. `None` is *the scope's default* — THE sheet rule (qualifier > default > the only sheet of a single-sheet book > `SheetRequired`). `wb.edit` has no default: give one with `wb.editIn(sheet)(…)`, qualify with `Some(sheet)`, or use `sheet.edit` (a one-sheet scope; a target naming another sheet is `SheetNotFound`). `Loc.parse("'Q1 Data'!B7")`, `Area.parse("A1:B2")`, `RowSpan.parse("10:20")`, `ColSpan.parse("E:H")` read the CLI spellings; `Area.cell(loc)` is a 1x1 area.
- **All-or-nothing**: edits apply in order and the first failure is `Left(EditFailed(index, op, cause))` — `index` 1-based (`err.opIndex`), `op` the kebab name (`drag-formula`), and `err.code`/`hint`/`candidates` are the *cause's* (`err.root` is the innermost cause) — with the input workbook untouched. `err.message` reads `op 2 (drag-formula): …` — the same 1-based position `xl batch` reports as `location.opIndex` on a `BATCH_OP_FAILED`; `orExit` prints it as the CLI's stderr diagnostic.
- **Formats**: `Edit.put(ref, a)` / `Edit.put(loc, a)` lift the codec's format as `FormatHint.Inferred` (fills a General format only, as `sheet.put` does); `FormatHint.Explicit(fmt)` replaces the number format and keeps font/fill/border. `StyleOverlay` is a partial style (every field `Option`, so "un-bold" is sayable; `++` right-biased; `StyleOverlay.of(style)`); `StyleMode.Merge` overlays each cell's style, `Replace` starts from `CellStyle.default`.
- **Scope and the fold**: the scope type is **`EditScope`** on the prelude surface (the source name `ops.Scope` collides with JMH's/ZIO's) — `EditScope.none`, `EditScope.of(sheet)`; a `RenameSheet` of the default sheet retargets the edits after it. `Edit.applyAll(wb, edits, scope)` → `Applied(workbook, planned, scope)` (`touchedBySheet`, `structural`); `Edit.plan(wb, edits, scope)` keeps only the `Planned(index, edit, sheet, touched)` rows — a *semantic* dry-run that fails exactly where `applyAll` would; `Edit.validate(edit)` is the static, workbook-free check (counts, spans, levels, formula syntax); `Edit.lower(edit, sheet)` / `Patch.toEdits(patch, sheet)` bridge to `Patch`. `EditSchema.all` / `find("putf")` / `nameOf(edit)` is the algebra's own metadata — one `EditSpec` per case (49 rows, each naming its `batchOp`/`cliVerb`), the shape of the batch schema but not the document `xl batch --schema` prints (xl-cli's `OpRegistry`, 32 ops, a subset that has already drifted in detail: `copy.target`).
- **Edit vs verb, today**: `Edit.MoveSheet.toIndex` is the sheet's FINAL 0-based position, while `xl move-sheet --to N` still reads `N` against the pre-removal order and clamps ([#583](https://github.com/TJC-LP/xl/issues/583)); `Edit.Copy.target` is a single cell (`Loc`) where batch `copy` also takes a range. Port CLI recipes with those two in mind.
- The prelude's `given FormulaSupport` (xl-evaluator's) powers the formula-aware cases; `FormulaSupport.textOnly` (`wb.edit(e)(using FormulaSupport.textOnly)`) stores formula text as written and **refuses** drags, structural edits and renames with `UnsupportedCapability` rather than write a silent `#REF!`.

### Typed extraction

Records first (0.21.0): a case class that `derives RowCodec` reads and writes whole rows — field
order is column order, field names are the header row, `Option[T]` fields are empty cells.

```scala
final case class Product(name: String, units: Int, price: BigDecimal, note: Option[String])
  derives RowCodec

val products = sheet.readRowsByHeader[Product](Row.from1(1)) // Either[RowCodecError, Vector[Product]]
val placed = Sheet("Out").putRowsWithHeader(ref"A1", products.unsafe).unsafe // header + rows
placed.dataRange                                             // Some(A2:D…) — style/filter/total from here
Sheet("Out").putTable(ref"A1", products.unsafe, "Products")  // + an Excel table over header + rows
```

`readRows[Product](range)` is the positional twin (range width must equal the record's); errors
are `RowCodecError.Field(row, column, field, cause)` / `Missing` / `HeaderNotFound` / `Width`
with `.message` and `.toXLError`. Per-cell reads remain for ad-hoc shapes:

```scala
final case class Product(name: String, units: Int, price: BigDecimal)

val products = (2 to 10).toList.flatMap { row =>
  val r = ref"A1".down(row - 1)
  for
    name <- sheet.readTypedOpt[String](r)
    units <- sheet.readTypedOpt[Int](r.right(1))
    price <- sheet.readTypedOpt[BigDecimal](r.right(2))
  yield Product(name, units, price)
}
```

9 codec types: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText. Use `readTyped` (full `Either[CodecError, Option[A]]`) when you must distinguish a type mismatch from an empty cell; `readTypedOr(ref, default)` when you just need a value.

Formula cells read through their cached value (since 0.20.0; GH-477): after `recalculate()`, `writeRecalculated`, or `Excel.read` of a book Excel saved, `readTypedOpt[BigDecimal](ref"B1")` on `Formula("A1*3", Some(Number(6)), _)` is `Some(6)` — never unwrap `CellValue.Formula(_, Some(v), _)` by hand. A formula authored with `fx"…"` and not yet recalculated has no cache, so `readTyped` is `Left(TypeMismatch)`, `readTypedOpt` is `None`, and `readTypedOr` is the default: recalculate first. `readTypedStrict` (0.20.0) rejects every formula cell, cached or not, when the distinction itself is what you are checking. On ≤0.19.3 every formula cell is `Left(TypeMismatch)` / `None` / the default whatever its cache: match `CellValue.Formula(_, Some(v), _)` yourself there.

### Styling

```scala
val header = CellStyle.default.bold.size(12.0).center.bgBlue.white
val pct = CellStyle.default.withNumFmt(NumFmt.Percent)

sheet
  .style(ref"A1:D1", header)
  .put(ref"E2", 0.345, pct)                     // put with inline style
  .put(ref"A3", "OK ".green.bold + "ship it")   // rich text runs
```

Smart detection for raw strings preserves formats: `"45.5%".toFormatted` stores `0.455` with Percent format; `"$1,234.56".toFormatted` → Currency; `"2025-01-15".toFormatted` → Date. (Also available everywhere as `FormattedParsers.detect`.) `.underline` is single; for double/accounting variants use `CellStyle.default.withUnderline(Underline.Double)` (0.20.0; `= withFont(font.withUnderline(u))`).

### Formulas & recalculation

```scala
val model = Sheet("Model")
  .put(ref"A1", 1000)
  .put(ref"A2", fx"=A1*1.08")
  .put(ref"A3", fx"=A2*1.08")

// 0.21.0: compute the formulas that have no cached value (all of them here), write, report.
val result = Excel.writeChecked(Workbook(model), "model.xlsx")   // total: never throws, never partial-silently
if !result.isClean then
  result.errors.foreach(e => println(s"⚠ ${e.render}"))          // e.g. "Model!A7: Circular reference"
```

**A freshly built model is written with `writeChecked` or `writeRecalculated`, never `Excel.write`.** `Excel.writeChecked(wb, path)` (0.21.0) computes only the formulas that have no cached value — in dependency order, reading every input's cache as it is — writes the workbook, and returns the `RecalcResult`; every cache the book already carried (Excel's, another engine's) is written byte for byte, so an edited model keeps what it had and gains what it lacked. `Excel.writeRecalculated(wb, path)` (0.13.0) recomputes *every* formula instead — reach for it when the caches themselves are suspect. Both write even when some formulas fail (errors are data; failed cells stay uncached and Excel computes them on open), and both take a `RecalcOptions` (0.21.0) for a fixed clock, seeded rng, iterative mode or parallelism. The pure step is `wb.recalculate()`: it evaluates every formula across all sheets in dependency order, resolves cross-sheet references automatically, isolates reference cycles (the rest of the workbook still computes), and reports failures per cell in `result.errors`; `result.toEither` gives `Left(errors)` for pipelines that must abort before anything lands on disk — then `Excel.write(result.workbook, path)` is the one legitimate plain write of a model, because the workbook is already cached. **Excel error values are results, not failures** (0.14.0): `=1/0` evaluates to `#DIV/0!` — catchable with `IFERROR`, cached into the written file exactly like Excel would, and listed via `result.excelErrors` rather than `result.errors` (which carries only host failures: parse errors, missing sheets, cycles). For one-off questions: `wb.evaluateFormula("=SUM(Data!A:A)", "Summary")`.

**Defined names resolve** (0.13.0): `fx"=IF(case=2,rev,cost)"`, `fx"=entry_mult*ltm_ebitda"`, `fx"=SUM(rev_range)"` evaluate against workbook- and sheet-scoped defined names (sheet-scoped shadows global), contribute dependency edges so recalc orders name-gated families correctly, and round-trip byte-faithfully; an unresolvable name is a clean per-cell error.

**Circular models are opt-in** (0.13.0): professional schedules (interest on average debt) ship circular by design. `wb.recalculate(IterativeCalc(maxIter = 100, maxChange = BigDecimal("0.001")))` Jacobi-fixpoints declared cycles instead of erroring; plain `recalculate()` still isolates cycles as errors. Honor a file's own `<calcPr>` with `wb.metadata.calcPr.filter(_.iterativeCalculation).map(IterativeCalc.fromCalcPr).fold(wb.recalculate())(wb.recalculate)`, and author it on scratch builds with `wb.withCalcPr(CalcPr(iterativeCalculation = true, maxIterations = Some(100), maxChange = Some(BigDecimal("0.001"))))`.

**One options record** (since 0.20.0): every recalculation knob lives on `RecalcOptions`, and `RecalcOptions()` reproduces `recalculate()` exactly, so there is one thing to learn:

```scala
// since 0.20.0 — fragment, not a runnable script
val opts = RecalcOptions(iterative = IterativeMode.FromCalcPr, parallelism = 4)  // clock, rng, seedTables too
wb.recalculate(opts).summary                            // "Recalculated 12 formulas" — the line `xl recalc` prints
wb.recalculateAfterEdit(sheet, Set(ref"A1"), opts)      // only the edit's dependents; every other cache byte-identical
wb.recalculateUncached(opts)                            // only cache-less formulas; cached cells never touched
```

**Renaming a sheet rewrites its references** (since 0.20.0, [#559](https://github.com/TJC-LP/xl/issues/559)): `SheetRenamer.rename(wb, SheetName.unsafe("Sheet1"), SheetName.unsafe("Q1 Data"))` renames the tab AND rewrites `Sheet1!A1` → `'Q1 Data'!A1` in cell formulas on every sheet, defined names, CF and DV formulas, caches preserved; it refuses (`Left`) when a dependent that mentions the sheet cannot be parsed. `Workbook.rename` alone stays tab-only. `SheetRenamer.references(wb, sheet)` lists the cells a rename would touch; `FormulaOps.renameSheet/shift/mentionsSheet` are the string-level pieces.

112 functions supported (SUM/SUMIFS/VLOOKUP/XLOOKUP/INDEX/MATCH/INDIRECT/MROUND/RAND/NPV/IRR/SEARCH/N/HYPERLINK/CELL/... plus LET) — full list in `reference/API.md`.

### Data tables (what-if sensitivity)

Native Excel `TABLE()` two-variable data tables (0.18.0) — the house sensitivity engine. `interior` is the RESULT GRID: for `D5:F6` the corner formula sits at C4 (one-up-one-left), the row-input axis rides D4:F4 (above) and the column-input axis rides C5:C6 (left). Only the corner cell carries the record — exactly Excel's own bytes.

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.0
import com.tjclp.xl.scripting.{*, given}

val model = Sheet("Sensitivity")
  .put(ref"B1", 0.08)                        // row input (perturbed by D4:F4)
  .put(ref"B2", 10.0)                        // column input (perturbed by C5:C6)
  .put(ref"C4", fx"=B1*B2*100")              // corner formula the table re-evaluates
  .put(ref"D4", 0.06).put(ref"E4", 0.08).put(ref"F4", 0.10)
  .put(ref"C5", 9.0).put(ref"C6", 11.0)

val authored = model.dataTable(ref"D5:F6", rowInput = ref"B1", colInput = ref"B2").unsafe

// Seed the interior caches: under calcMode="autoNoTable" (the house dialect) Excel does NOT
// recompute data tables on open — even with fullCalcOnLoad — so an unseeded grid opens BLANK.
// Plain calcMode="auto" books self-heal on open; calc levers stay on Workbook.withCalcPr.
val seeded = Workbook(authored).seedDataTables().unsafe

// writeRecalculated, NOT write: seedDataTables caches the record cell and the interior, never the
// CORNER formula (C4) — Excel.write would ship an uncached corner that previews as blank.
Excel.writeRecalculated(seeded, "sensitivity.xlsx")
```

1-D shapes: `sheet.dataTableRow(interior, rowInput)` (axis above, source formulas in the column left — one per interior row) and `sheet.dataTableCol(interior, colInput)` (axis left, source formulas in the row above — one per interior column; multi-result-column tables are legal). All three take optional row-major `seeds` to ship byte-exact caches without evaluating; the corner absorbs a pre-existing plain scalar as its cache, so `fillBy`-then-`dataTable` composes. Authoring refuses to tear existing tables, overwrite real formulas, or accept inputs inside the table block — each with a structured `XLError`.

### Error handling: Either everywhere, unsafe once

Everything fallible returns `XLResult[A]` (= `Either[XLError, A]`). Compose with for-comprehensions; unwrap **once** at the script edge:

```scala
val result: XLResult[Workbook] =
  for
    wb <- Right(Excel.read("in.xlsx"))
    s <- wb("Sales")                      // sheet lookup may fail
    upd <- wb.update("Sales", _.put(ref"A1", s.cells.size))
  yield upd

result match
  case Right(wb) => Excel.write(wb, "out.xlsx")
  case Left(err) => println(s"failed: ${err.message}"); sys.exit(1)
```

Or lean on totality so there is nothing to unwrap: literal refs, `upsert`, range fill, `readTypedOr`, `recalculate` are all total. `.unsafe` throws a structured `XLException` (wraps the `XLError`) — fine for scripts where fail-fast is correct. `orExit(result)` (0.21.0) is the script-shaped form of the `match` above: the value on `Right`, or the error on stderr and exit status 1 as `Error: <message>` followed by indented `code: <CODE>`, `did you mean: …` (when the error has candidates) and `hint: …` (when it has one) — `XLError.renderDiagnostic`, the very renderer behind the CLI's `Diagnostics`, so a failing script prints the same bytes as a failing `xl` call (`exitMessage(err)` is that text):

```scala
val wb = orExit(Workbook.named("Data", "Summary"))   // DuplicateSheet on a repeat → printed, exit 1
val sales = orExit(wb("Sales"))                      // SheetNotFound → printed with its hint, exit 1
```

## Workflows

### Merge many workbooks into one

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.0
import com.tjclp.xl.scripting.{*, given}
import java.nio.file.{Files, Paths}
import scala.jdk.CollectionConverters.*

val inputs = Files.list(Paths.get("reports")).iterator.asScala
  .filter(_.toString.endsWith(".xlsx")).toList.sortBy(_.toString)

val merged = inputs.foldLeft(Workbook.empty) { (acc, path) =>
  val wb = Excel.read(path.toString)
  wb.sheets.foldLeft(acc) { (a, sheet) =>
    a.put(sheet.copy(name = SheetName.unsafe(s"${path.getFileName.toString.stripSuffix(".xlsx")}-${sheet.name.value}".take(31))))
  }
}
Excel.write(merged.remove("Sheet1").getOrElse(merged), "merged.xlsx")
println(s"merged ${inputs.size} files, ${merged.sheets.size} sheets")
```

### Data → styled report

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.0
import com.tjclp.xl.scripting.{*, given}

val data = List(("North", 125000.50), ("South", 98000.25), ("West", 143500.00))
val header = CellStyle.default.bold.size(12.0).center
val currency = CellStyle.default.withNumFmt(NumFmt.Currency)

val body = data.zipWithIndex.foldLeft(Patch.empty) { case (acc, ((region, sales), i)) =>
  val r = ref"A4".down(i)
  acc ++ (r := region) ++ (r.right(1) := BigDecimal(sales)) ++ r.right(1).styled(currency)
}

val report = Sheet("Q2")
  .put(
    (ref"A1" := "Q2 Regional Sales") ++ ref"A1:B1".merge ++ ref"A1".styled(header) ++
      (ref"A3" := "Region") ++ (ref"B3" := "Sales") ++ body ++
      (ref"A8" := "Total") ++ (ref"B8" := fx"=SUM(B4:B6)") ++ ref"B8".styled(currency)
  )

val result = Excel.writeChecked(Workbook(report), "/tmp/q2-report.xlsx") // caches B8, writes, reports
println(if result.isClean then "✓ report written" else result.errors.map(_.render).mkString("\n"))
```

### Streaming a 500k-row file (constant memory)

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.0
import com.tjclp.xl.scripting.{*, given}
import cats.effect.IO
import cats.effect.unsafe.implicits.global
import java.nio.file.Paths

val excel = ExcelIO.instance[IO]
val total = excel.readStream(Paths.get("huge.xlsx"))      // fs2.Stream[IO, RowData], O(1) memory
  .map(_.cells.get(2))                                    // column C (0-based)
  .collect { case Some(CellValue.Number(n)) => n }
  .compile.fold(BigDecimal(0))(_ + _)
  .unsafeRunSync()
println(s"column C total: $total")
```

Switch to streaming above ~100k rows; `Excel.read` loads the whole workbook. Streaming writes: `Stream.emits(rows).through(excel.writeStream(path, "Sheet1")).compile.drain` with `RowData(rowIndex, Map(colIdx -> CellValue))`.

## Gotchas

- **`{*, given}` is required** on the prelude import — plain `*` misses the given instances (codecs, conversions, display).
- **Never combine** `com.tjclp.xl.scripting.{*, given}` with `com.tjclp.xl.{*, given}` in one file.
- **`$$` escapes `$`** inside `money""` and other interpolated literals: `money"$$1,234.56"`.
- **Compose patches with `++`**, not Cats `|+|` (the latter needs type ascription on enum cases).
- **`fx` with runtime interpolation returns `Either`** — there is deliberately no `:=` overload that swallows a `Left`; unwrap with `.unsafe` or sequence it.
- **`wb.update` fails on a missing sheet; `wb.upsert` creates it.** Pick by intent.
- **`wb.edit` has no default sheet** (0.21.0): an `Edit` target with `sheet = None` resolves only on a single-sheet book — otherwise `SheetRequired`, reported as `EditFailed` at that edit's 1-based index. Give the default with `wb.editIn(sheet)(…)`, qualify with `Some(sheet)`, or use `sheet.edit`. The scope type is spelled **`EditScope`** on the prelude surface (`EditScope.none` / `EditScope.of(sheet)`), not `Scope`; `Edit.FillDir` / `Edit.SortDir` / `Edit.SortMode` nest in the companion (xl-cli's `FillDirection`/`SortDirection`/`SortMode` are not on this surface).
- **`wb.rename` does NOT rewrite formulas** ([#559](https://github.com/TJC-LP/xl/issues/559)): it changes the tab and leaves `Sheet1!A1` in every dependent — the file lints clean and Excel shows `#REF!`. Since 0.20.0 use `SheetRenamer.rename(wb, from, to)` (what `xl rename-sheet` and batch `rename-sheet` now do): formulas on every sheet, defined names, CF and DV follow the rename with caches preserved. On ≤0.19.3 rewrite dependents yourself (`FormulaParser.parse` → walk → `FormulaPrinter.printFileForm`) or rename before authoring cross-sheet formulas.
- **Range fill cost = range size**: `ref"A:A" := 0` really creates 1,048,576 cells (that's what a fill means) — size fill ranges to your data.
- **`shift`/`down`/`up`/`left`/`right` are unchecked at the edges**: `ref"A1".up()` produces an invalid "A0" ref that corrupts output if written. Since 0.20.0 ([#465](https://github.com/TJC-LP/xl/issues/465)) use the bounded forms in loops — `ref.tryDown(n)`/`tryRight(n)`/`tryShift(dc, dr)` return `None` past the grid (column A..XFD, row 1..1048576) and `ref.clampShift(dc, dr)` pins each axis to the nearest edge; `range.rows`/`columns` and `range.row(i)`/`column(i)` (`Option`, 0-based) slice a range instead of interpolating its corners. On ≤0.19.x keep loop bounds inside your data extent.
- **First run is slow** (dependency download); afterwards scala-cli caches everything.
- **`.sc` files**: top-level statements, no `@main`. A `.scala` file needs `@main def run(): Unit`.
- **Never `Excel.write` a freshly built model — it does NOT recalculate.** `fx"…"` cells are written with no cached values, so Excel-before-recalc, openpyxl `data_only`, pandas, and previewers all show blanks. Write models with **`Excel.writeChecked(wb, path)`** (0.21.0, [#589](https://github.com/TJC-LP/xl/issues/589)): it computes only the formulas that have no cache (every existing cache — Excel's or another engine's — is written byte for byte), writes, and returns the `RecalcResult`; or with **`Excel.writeRecalculated(wb, path)`** (0.13.0, [#360](https://github.com/TJC-LP/xl/issues/360)) when every formula must be recomputed. Both write even when some formulas fail — errors are data; inspect `result.errors` / `result.isClean` — and both take a `RecalcOptions` (0.21.0). For fail-hard pipelines that must abort *before* anything lands on disk, keep the explicit `val result = wb.recalculate(); …; Excel.write(result.workbook, path)` pattern — the one plain `write` of a model that is right, because the workbook is already cached. On ≤0.20.0 `writeChecked` is unavailable: use `writeRecalculated`; on ≤0.12.x recalculate then write `result.workbook` (a single pass suffices on 0.12.5+).
- **Percent postfix works since 0.13.0** ([#355](https://github.com/TJC-LP/xl/issues/355)): `fx"=A1*10%"`, `fx"=10%"`, `fx"=(1+5%)^2"` parse, evaluate (`10%` → exact `0.1`), broadcast over ranges, and print back byte-identically (never rewritten to `/100`). On ≤0.12.x the parser rejects `%` — write `/100` there. External-workbook refs (`[2]Book!A1`) parse and pin their Excel-written caches **since 0.12.6** ([#353](https://github.com/TJC-LP/xl/issues/353)): `recalculate()` preserves those cells verbatim and dependents compute from the caches (uncached external cells yield a per-cell error); on ≤0.12.5 they fail to parse entirely — compute from cached values there.
- **Runtime column handles for `setColumnProperties`** ([#361](https://github.com/TJC-LP/xl/issues/361), since 0.13.0): fold over letters computed at runtime with `Column.parse("D")` (`Either[String, Column]`; trailing row digits tolerated, so `"D1"` works) — e.g. `Column.parse(letter).map(c => sheet.setColumnProperties(c, ColumnProperties(width = Some(w))))`. A runtime `RefType` also exposes `.col` (`RefType.parse(s).map(_.col)`). On ≤0.12.x only the compile-time `ref"D1".col` existed — set widths with literal refs per column there.
- **A runtime string flips the return type of every transparent string form** — `Sheet(name)`, `Workbook(name, …)`, `sheet.put("A1", v)`, `sheet.style("A1:D1", st)`, `sheet.merge("A1:C1")`, `sheet.comment("A1", c)` ([#420](https://github.com/TJC-LP/xl/issues/420), [#465](https://github.com/TJC-LP/xl/issues/465)): they are `transparent inline` — a string literal validates at compile time and returns `Sheet`/`Workbook`, while the very same call with a `val` returns `XLResult[…]`, so a chained `.put(...)` type-errors with nothing at the call site to warn you. Do not rely on those forms for computed strings. Use the twins that spell `XLResult` in their signatures: **`Sheet.named(name)`** (0.18.0) and, since 0.20.0, **`Workbook.named(…)`** (`DuplicateSheet` on repeats), **`sheet.putAt(ref, v)`** / **`putAt(ref, v, style)`**, **`sheet.styleAt(ref, st)`** (cell or range), **`sheet.mergeAt(range)`**, **`sheet.commentAt(ref, c)`** — the validation of the **literal** forms and the same value/style path; a range where a cell is required is `InvalidCellRef`, a sheet-qualified ref is `InvalidReference` (qualify at the workbook: `wb.update(name, _.putAt(...))`). **The twins take corner forms only** (`A1`, `A1:B2`): no full-column/row `A:A`/`1:1`, no `$` anchors, and `mergeAt` needs two corners — spellings the *dynamic* transparent `merge`/`style` accept today via `CellRange.parse`, so do not rewrite `sheet.merge(s"$c:$c")` as `mergeAt`; parse first instead: `s.asRange.map(sheet.merge)` / `s.asRange.map(r => sheet.style(r, st))` / `s.asCell.map(r => sheet.put(r, v))` (`String.asRange` is `CellRange.parse`-backed, `asCell` is `ARef.parse`-backed). On ≤0.19.x, make the union explicit with an ascription: `val s: XLResult[Sheet] = sheet.put(cell, 42)`.

## Reference

- `reference/API.md` — types, extension methods, the `Edit` algebra, style builders, all 108 formula functions, streaming API
- `reference/RECIPES.md` — 9 complete, runnable scripts (bulk transform, typed extraction, model build, merge, streaming, diff, CSV ingest, recalculated write + runtime column widths, deliverable finish)
- Repo examples: `examples/*.sc` in https://github.com/TJC-LP/xl (start with `scripting_tour.sc`)
- The `xl-cli` skill for CLI operations (visual exports, quick inspection)
