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
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.11.1

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

// Cell writes (literal refs are compile-time validated and infallible)
sheet.put(ref"A1", "Title")                        // Sheet
sheet.put("A1", "Title")                           // literal string: Sheet; runtime string: XLResult[Sheet]
sheet.put(ref"B1", 42)                             // Int/Long/Double/BigDecimal/Boolean/LocalDate(Time)/RichText
sheet.put(ref"C1", "$1,234.56".toFormatted)        // smart detection → Currency format
sheet.style(ref"A1:D1", CellStyle.default.bold)    // Sheet

// Patch DSL (compose pure values, apply once)
val patch = (ref"A1" := "Report") ++ ref"A1:C1".merge ++ ref"A1".styled(CellStyle.default.bold)
sheet.put(patch)                                   // Sheet
ref"E2:E9" := 0                                    // range fill: every cell (Ctrl+Enter semantics)
ref"A2".down(3).right(1)                           // total navigation → B5

// Typed reads
sheet.readTyped[BigDecimal](ref"C1")               // Either[CodecError, Option[BigDecimal]]
sheet.readTypedOr[Int](ref"B1", 0)                 // total, with default
sheet.readTypedOpt[LocalDate](ref"D1")             // flat Option

// Formulas
sheet.put(ref"D2", fx"=B2*C2")                     // compile-time validated literal
wb.evaluateFormula("=SUM(Sales!A1:A9)", "Summary") // XLResult[CellValue], cross-sheet aware
val r = wb.recalculate()                           // RecalcResult: total, per-cell errors
r.isClean; r.errors.map(_.render); r.workbook      // inspect, then write r.workbook

// Errors: XLResult[A] = Either[XLError, A]; unwrap ONCE at the edge
wb.update("Sales", f).unsafe                       // throws structured XLException if Left
```

## Essential Patterns

### Read → modify → write

```scala
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.11.1
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

**Prefer total navigation over interpolated refs in loops** — no Either at all:

```scala
val base = ref"A2"
val r3 = base.down(2)         // A4
val c2 = base.right(1)        // B2
val shifted = base.shift(1, 2) // B4
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

### Styling

```scala
val header = CellStyle.default.bold.size(12.0).center.bgBlue.white
val pct = CellStyle.default.withNumFmt(NumFmt.Percent)

sheet
  .style(ref"A1:D1", header)
  .put(ref"E2", 0.345, pct)                     // put with inline style
  .put(ref"A3", "OK ".green.bold + "ship it")   // rich text runs
```

Smart detection for raw strings preserves formats: `"45.5%".toFormatted` stores `0.455` with Percent format; `"$1,234.56".toFormatted` → Currency; `"2025-01-15".toFormatted` → Date. (Also available everywhere as `FormattedParsers.detect`.)

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

`recalculate` evaluates every formula across all sheets in dependency order, resolves cross-sheet references automatically, isolates reference cycles (the rest of the workbook still computes), and reports failures per cell in `result.errors`. `result.toEither` gives `Left(errors)` for fail-hard pipelines. For one-off questions: `wb.evaluateFormula("=SUM(Data!A:A)", "Summary")`.

104 functions supported (SUM/SUMIFS/VLOOKUP/XLOOKUP/INDEX/MATCH/NPV/IRR/...) — full list in `reference/API.md`.

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
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.11.1
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
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.11.1
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
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.11.1
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
- **Range fill cost = range size**: `ref"A:A" := 0` really creates 1,048,576 cells (that's what a fill means) — size fill ranges to your data.
- **Navigation is unchecked at the edges**: `ref"A1".up()` produces an invalid "A0" ref that corrupts output if written; keep loop bounds inside your data extent.
- **First run is slow** (dependency download); afterwards scala-cli caches everything.
- **`.sc` files**: top-level statements, no `@main`. A `.scala` file needs `@main def run(): Unit`.

## Reference

- `reference/API.md` — types, extension methods, style builders, all 104 formula functions, streaming API
- `reference/RECIPES.md` — 7 complete, runnable scripts (bulk transform, typed extraction, model build, merge, streaming, diff, CSV ingest)
- Repo examples: `examples/*.sc` in https://github.com/TJC-LP/xl (start with `scripting_tour.sc`)
- The `xl-cli` skill for CLI operations (visual exports, quick inspection)
