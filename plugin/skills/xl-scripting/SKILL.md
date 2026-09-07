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
//> using dep com.tjclp::xl:0.20.0

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
Excel.write(wb, "out.xlsx")                        // also accepts XLResult[Workbook]
Excel.modify("file.xlsx")(_.upsert("Log", identity)) // atomic in-place read→transform→write

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

// Patch DSL (compose pure values, apply once)
val patch = (ref"A1" := "Report") ++ ref"A1:C1".merge ++ ref"A1".styled(CellStyle.default.bold)
sheet.put(patch)                                   // Sheet
ref"E2:E9" := 0                                    // range fill: every cell (Ctrl+Enter semantics)
ref"A2".down(3).right(1)                           // total navigation → B5 (unchecked at the grid edge)
ref"A2".tryDown(3)                                 // 0.20.0: bounded → Some(A5); None past the edge; clampShift pins
ref"A1:D10".rows                                   // 0.20.0: lazy one-row slices; row(i)/column(i) are Option

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
Excel.writeRecalculated(wb, "out.xlsx")            // 0.13.0: recalc + write + RecalcResult in one call

// Errors: XLResult[A] = Either[XLError, A]; unwrap ONCE at the edge
wb.update("Sales", f).unsafe                       // throws structured XLException if Left
```

## Essential Patterns

### Read → modify → write

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.20.0
import com.tjclp.xl.scripting.{*, given}

val wb = Excel.read("input.xlsx")
val updated = wb
  .upsert("Audit", _.put(ref"A1", "reviewed"))         // total: creates sheet if missing
  .update("Data", _.put(ref"B2", 99))                  // XLResult: Data must exist
  .unsafe                                              // one unwrap, at the edge
Excel.write(updated, "output.xlsx")
```

`Excel.modify("file.xlsx")(f)` does the same in place with atomic file replacement.

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

val result = Workbook(model).recalculate()       // total: never throws, never partial-silently
if !result.isClean then
  result.errors.foreach(e => println(s"⚠ ${e.render}"))   // e.g. "Model!A7: Circular reference"
Excel.write(result.workbook, "model.xlsx")       // computed values cached for Excel/viewers
```

`recalculate` evaluates every formula across all sheets in dependency order, resolves cross-sheet references automatically, isolates reference cycles (the rest of the workbook still computes), and reports failures per cell in `result.errors`. `result.toEither` gives `Left(errors)` for fail-hard pipelines. **Excel error values are results, not failures** (0.14.0): `=1/0` evaluates to `#DIV/0!` — catchable with `IFERROR`, cached into the written file exactly like Excel would, and listed via `result.excelErrors` rather than `result.errors` (which now carries only host failures: parse errors, missing sheets, cycles). For one-off questions: `wb.evaluateFormula("=SUM(Data!A:A)", "Summary")`. When the very next step is a write, `Excel.writeRecalculated(wb, path)` (0.13.0) fuses recalculate + write and returns the same `RecalcResult` — see the gotcha below.

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
//> using dep com.tjclp::xl:0.20.0
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

Or lean on totality so there is nothing to unwrap: literal refs, `upsert`, range fill, `readTypedOr`, `recalculate` are all total. `.unsafe` throws a structured `XLException` (wraps the `XLError`) — fine for scripts where fail-fast is correct.

## Workflows

### Merge many workbooks into one

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.20.0
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
//> using dep com.tjclp::xl:0.20.0
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

val result = Workbook(report).recalculate()
Excel.write(result.workbook, "/tmp/q2-report.xlsx")
println(if result.isClean then "✓ report written" else result.errors.map(_.render).mkString("\n"))
```

### Streaming a 500k-row file (constant memory)

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.20.0
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
- **`wb.rename` does NOT rewrite formulas** ([#559](https://github.com/TJC-LP/xl/issues/559)): it changes the tab and leaves `Sheet1!A1` in every dependent — the file lints clean and Excel shows `#REF!`. Since 0.20.0 use `SheetRenamer.rename(wb, from, to)` (what `xl rename-sheet` and batch `rename-sheet` now do): formulas on every sheet, defined names, CF and DV follow the rename with caches preserved. On ≤0.19.3 rewrite dependents yourself (`FormulaParser.parse` → walk → `FormulaPrinter.printFileForm`) or rename before authoring cross-sheet formulas.
- **Range fill cost = range size**: `ref"A:A" := 0` really creates 1,048,576 cells (that's what a fill means) — size fill ranges to your data.
- **`shift`/`down`/`up`/`left`/`right` are unchecked at the edges**: `ref"A1".up()` produces an invalid "A0" ref that corrupts output if written. Since 0.20.0 ([#465](https://github.com/TJC-LP/xl/issues/465)) use the bounded forms in loops — `ref.tryDown(n)`/`tryRight(n)`/`tryShift(dc, dr)` return `None` past the grid (column A..XFD, row 1..1048576) and `ref.clampShift(dc, dr)` pins each axis to the nearest edge; `range.rows`/`columns` and `range.row(i)`/`column(i)` (`Option`, 0-based) slice a range instead of interpolating its corners. On ≤0.19.x keep loop bounds inside your data extent.
- **First run is slow** (dependency download); afterwards scala-cli caches everything.
- **`.sc` files**: top-level statements, no `@main`. A `.scala` file needs `@main def run(): Unit`.
- **`Excel.write` does NOT recalculate** — freshly built `fx"…"` cells are written with no cached values, so Excel-before-recalc, openpyxl `data_only`, pandas, and previewers all show blanks. Since 0.13.0 the one-call fix is **`Excel.writeRecalculated(wb, path)`** ([#360](https://github.com/TJC-LP/xl/issues/360)): it recalculates, writes the cached workbook (even when some formulas fail — errors are data), and returns the `RecalcResult` (inspect `result.errors` / `result.isClean`). For fail-hard pipelines that must abort *before* anything lands on disk, keep the explicit `val result = wb.recalculate(); …; Excel.write(result.workbook, path)` pattern. On ≤0.12.x, `writeRecalculated` is unavailable — recalculate then write `result.workbook` (a single pass suffices on 0.12.5+).
- **Percent postfix works since 0.13.0** ([#355](https://github.com/TJC-LP/xl/issues/355)): `fx"=A1*10%"`, `fx"=10%"`, `fx"=(1+5%)^2"` parse, evaluate (`10%` → exact `0.1`), broadcast over ranges, and print back byte-identically (never rewritten to `/100`). On ≤0.12.x the parser rejects `%` — write `/100` there. External-workbook refs (`[2]Book!A1`) parse and pin their Excel-written caches **since 0.12.6** ([#353](https://github.com/TJC-LP/xl/issues/353)): `recalculate()` preserves those cells verbatim and dependents compute from the caches (uncached external cells yield a per-cell error); on ≤0.12.5 they fail to parse entirely — compute from cached values there.
- **Runtime column handles for `setColumnProperties`** ([#361](https://github.com/TJC-LP/xl/issues/361), since 0.13.0): fold over letters computed at runtime with `Column.parse("D")` (`Either[String, Column]`; trailing row digits tolerated, so `"D1"` works) — e.g. `Column.parse(letter).map(c => sheet.setColumnProperties(c, ColumnProperties(width = Some(w))))`. A runtime `RefType` also exposes `.col` (`RefType.parse(s).map(_.col)`). On ≤0.12.x only the compile-time `ref"D1".col` existed — set widths with literal refs per column there.
- **A runtime string flips the return type of every transparent string form** — `Sheet(name)`, `Workbook(name, …)`, `sheet.put("A1", v)`, `sheet.style("A1:D1", st)`, `sheet.merge("A1:C1")`, `sheet.comment("A1", c)` ([#420](https://github.com/TJC-LP/xl/issues/420), [#465](https://github.com/TJC-LP/xl/issues/465)): they are `transparent inline` — a string literal validates at compile time and returns `Sheet`/`Workbook`, while the very same call with a `val` returns `XLResult[…]`, so a chained `.put(...)` type-errors with nothing at the call site to warn you. Do not rely on those forms for computed strings. Use the twins that spell `XLResult` in their signatures: **`Sheet.named(name)`** (0.18.0) and, since 0.20.0, **`Workbook.named(…)`** (`DuplicateSheet` on repeats), **`sheet.putAt(ref, v)`** / **`putAt(ref, v, style)`**, **`sheet.styleAt(ref, st)`** (cell or range), **`sheet.mergeAt(range)`**, **`sheet.commentAt(ref, c)`** — the validation of the **literal** forms and the same value/style path; a range where a cell is required is `InvalidCellRef`, a sheet-qualified ref is `InvalidReference` (qualify at the workbook: `wb.update(name, _.putAt(...))`). **The twins take corner forms only** (`A1`, `A1:B2`): no full-column/row `A:A`/`1:1`, no `$` anchors, and `mergeAt` needs two corners — spellings the *dynamic* transparent `merge`/`style` accept today via `CellRange.parse`, so do not rewrite `sheet.merge(s"$c:$c")` as `mergeAt`; parse first instead: `s.asRange.map(sheet.merge)` / `s.asRange.map(r => sheet.style(r, st))` / `s.asCell.map(r => sheet.put(r, v))` (`String.asRange` is `CellRange.parse`-backed, `asCell` is `ARef.parse`-backed). On ≤0.19.x, make the union explicit with an ascription: `val s: XLResult[Sheet] = sheet.put(cell, 42)`.

## Reference

- `reference/API.md` — types, extension methods, style builders, all 108 formula functions, streaming API
- `reference/RECIPES.md` — 9 complete, runnable scripts (bulk transform, typed extraction, model build, merge, streaming, diff, CSV ingest, recalculated write + runtime column widths, deliverable finish)
- Repo examples: `examples/*.sc` in https://github.com/TJC-LP/xl (start with `scripting_tour.sc`)
- The `xl-cli` skill for CLI operations (visual exports, quick inspection)
