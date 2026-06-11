# Scripting Guide — `com.tjclp.xl.scripting`

The scripting prelude is the fastest way to use XL from a script, a REPL, or any JVM project that
just wants to get an `.xlsx` in and out without ceremony. One import gives you the core API, the
patch DSL, compile-time literals, formula evaluation, sync IO, streaming IO, smart value
detection, and the `.unsafe` boundary.

This guide is for Maven/scala-cli users of the published library. If you are driving XL through
Claude Code, the same material ships as the `xl-scripting` skill
([`plugin/skills/xl-scripting/SKILL.md`](../../plugin/skills/xl-scripting/SKILL.md)).

**Snippet convention** (same as the skill): every fenced block that starts with `//> using` is a
complete, standalone script — these are compile-verified in CI by
`scripts/verify-skill-snippets.sh`. Blocks without the directive header are fragments.

---

## One import

```scala
import com.tjclp.xl.scripting.{*, given}
```

- `{*, given}` is required — plain `*` misses the given instances (codecs, conversions, display).
- The prelude and the pure library import are **mutually exclusive**: never combine
  `com.tjclp.xl.scripting.{*, given}` with `com.tjclp.xl.{*, given}` in one file — the
  overlapping forwarders become ambiguous. `import com.tjclp.xl.{*, given}` remains the 100%
  pure alternative (no `.unsafe`, no sync `Excel` in scope) for library/production code.
- `java.time` types are not re-exported; `import java.time.LocalDate` yourself when needed.

The canonical script header (byte-identical across the skill, recipes, and this guide — a
release bump is a mechanical substitution):

```scala
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.11.3
import com.tjclp.xl.scripting.{*, given}

val sheet = Sheet("Demo").put(ref"A1", "Hello").put(ref"B1", 42)
Excel.write(Workbook(sheet), "/tmp/demo.xlsx")
println(s"wrote ${sheet.cells.size} cells")
```

Run with `scala-cli run script.sc`. `.sc` files take top-level statements — no `@main`, no
object wrapper.

## Read → modify → write

The sync `Excel` facade (`read`/`write`/`modify`) is the IO edge for scripts. Everything between
read and write is pure values.

```scala
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.11.3
import com.tjclp.xl.scripting.{*, given}

val wb = Excel.read("input.xlsx")
val updated = wb
  .upsert("Audit", _.put(ref"A1", "reviewed")) // total: creates the sheet if missing
  .update("Data", _.put(ref"B2", 99))          // XLResult: "Data" must exist
  .unsafe                                      // ONE unwrap, at the edge
Excel.write(updated, "output.xlsx")
```

- `Workbook.upsert(name, f)` is total update-or-create; `Workbook.update(name, f)` returns
  `XLResult[Workbook]` and fails if the sheet is absent. Pick by intent.
- `Excel.modify("file.xlsx")(f)` does read → transform → write in place with atomic file
  replacement (no ZIP corruption on a crashed write).
- `Excel.write` also accepts an `XLResult[Workbook]` directly.

## Compile-time vs runtime refs (and the `fx` rule)

Literal refs and formulas are validated **at compile time** — a typo fails the build, not the
workbook:

```scala
val a = ref"A1"              // ARef
val rng = ref"A1:B10"        // CellRange
val f = fx"=SUM(A1:B10)"     // CellValue.Formula (syntax/parens checked at compile time)
val m = money"$$1,234.56"    // Formatted(Number, Currency)
```

`$` is the interpolation character inside every interpolated literal, so Excel's absolute
anchors need `$$`: write `fx"=SUM($$A$$1:B10)"` to get `=SUM($A$1:B10)`. Same for
`money"$$1,234.56"`.

With runtime interpolation, validation moves to runtime and the macros return `Either`:

```scala
val row = 5
val cellE = ref"A$row"       // Either[XLError, RefType]
val formE = fx"=B$row*C$row" // Either[XLError, CellValue]
val cell2 = fx"=B$row*2".unsafe // explicit boundary when fail-fast is fine
```

Prefer **total navigation** over interpolated refs in loops — no `Either` at all:

```scala
val base = ref"A2"
base.down(2)     // A4   (default step is 1: base.down() == A3)
base.right(1)    // B2
base.up(1)       // A1
base.left(1)     // out of bounds! see below
base.shift(1, 2) // B4   (colOffset, rowOffset)
```

Navigation is total but **unchecked at the sheet edges**: `ref"A1".up()` produces the
non-existent "A0", which corrupts output if written. Keep loop bounds inside your data extent.

## Range fill and the patch DSL

Patches are pure values forming a monoid — build the whole change set with `++`, apply once with
`sheet.put(patch)`:

```scala
val patch = (ref"A1" := "Report") ++ ref"A1:C1".merge ++ ref"A1".styled(CellStyle.default.bold)
val sheet2 = sheet.put(patch)
```

`range := value` fills **every** cell in the range with the value — Excel Ctrl+Enter semantics:

```scala
ref"E2:E100" := 0          // 99 Puts, one per cell
ref"A1" := "one cell"      // a 1x1 fill is a single Put
```

Fill cost is proportional to range size by design — `ref"A:A" := 0` really creates 1,048,576
cells. Size fill ranges to your data.

## Formulas: build, recalculate, inspect

`wb.recalculate()` is a **total** whole-workbook recalculation: every formula on every sheet
evaluates in dependency order, cross-sheet references resolve automatically, and failures never
throw — they are collected per cell.

```scala
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.11.3
import com.tjclp.xl.scripting.{*, given}

val title = CellStyle.default.bold.size(14.0).center
val label = CellStyle.default.bold.indent(1)
val currencyStyle = CellStyle.default.currency
val totalRow = CellStyle.default.currency.bold.borderTop(BorderStyle.Thin)

val model = Sheet("Model").put(
  (ref"B1" := "FY2026 Plan") ++ ref"B1:C1".merge ++ ref"B1".styled(title) ++
    (ref"B3" := "Revenue") ++ (ref"C3" := 1200000) ++
    (ref"B4" := "Costs") ++ (ref"C4" := fx"=C3*0.62") ++
    (ref"B5" := "Profit") ++ (ref"C5" := fx"=C3-C4") ++
    ref"B3:B5".styled(label) ++ ref"C3:C5".styled(currencyStyle) ++
    ref"C5".styled(totalRow) ++ ref"B3:C5".outlined(BorderStyle.Medium)
)

Workbook(model).recalculate().toEither match
  case Right(wb) =>
    Excel.write(wb, "/tmp/plan.xlsx")
    given Sheet = wb.sheets.headOption.getOrElse(sys.exit(1))
    println(excel"Profit: ${ref"C5"}") // displays through NumFmt: $456,000.00
  case Left(errors) =>
    errors.foreach(e => println(s"✗ ${e.render}"))
    sys.exit(1)
```

Since 0.11.2, formulas may use `LET` (lexical bindings), `INDIRECT` (dynamic references —
evaluated in a deferred last partition), and `RAND`/`RANDBETWEEN`. Randomness is an explicit
capability: pass `Rng.seeded(42L)` to the rng-taking overloads (`wb.recalculate(clock, rng)`,
`sheet.evaluateFormula(f, clock, rng)`) for reproducible runs; the default is `Rng.system`.
For Excel-style format inheritance on formula entry, use the opt-in
`sheet.putFormulaInheriting(ref, formula)`.

`recalculate(clock: Clock = Clock.system)` returns a `RecalcResult`:

| Member | Meaning |
|--------|---------|
| `workbook` | The workbook with every successful formula cached (`Formula(expr, Some(value))`) |
| `evaluated` | `Map[SheetName, Map[ARef, CellValue]]` — computed values for inspection |
| `errors` | `Vector[CellEvalError]` — per-cell failures (parse/eval errors, cycle participants, cells blocked by a cycle) |
| `isClean` | `true` when `errors.isEmpty` |
| `toEither` | `Right(workbook)` when clean, `Left(errors)` otherwise — for fail-hard pipelines |

Reference cycles are **isolated**: the participants and their downstream dependents are reported
(e.g. `Model!A7: Formula error in '=B7': Circular reference` via `CellEvalError.render`) while
the acyclic remainder still evaluates and caches.

For one-off questions, `wb.evaluateFormula("=SUM(Data!A1:A9)", "Summary")` returns
`XLResult[CellValue]` with cross-sheet context wired automatically (107 functions supported —
see the [skill API reference](../../plugin/skills/xl-scripting/reference/API.md) for the full
list).

## Typed extraction

```scala
sheet.readTyped[BigDecimal](ref"C2")  // Either[CodecError, Option[BigDecimal]]
sheet.readTypedOr[Int](ref"B2", 0)    // total, with default
sheet.readTypedOpt[String](ref"A2")   // flat Option — mismatch and empty both None
```

Nine codec types: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime,
RichText. Use `readTyped` when you must distinguish a type mismatch from an empty cell;
`readTypedOr`/`readTypedOpt` when you just need a value.

## Smart value detection

`FormattedParsers.detect` (available everywhere) turns a raw string into a value + number
format; the prelude adds `String.toFormatted` sugar:

```scala
sheet.put(ref"C1", "$1,234.56".toFormatted) // Number(1234.56) + Currency format
"45.5%".toFormatted                          // Number(0.455) + Percent
"2026-01-15".toFormatted                     // DateTime + Date format
"plain text".toFormatted                     // Text (detection is total — never fails)
```

## Styling quick hits

```scala
CellStyle.default.bold.italic.underline.size(12.0).fontFamily("Arial")
CellStyle.default.center.middle.wrap.indent(2)      // alignment (+ Align indent)
CellStyle.default.red.bgGray                        // font / background color
CellStyle.default.currency                          // named formats: .percent .decimal .dateFormat .dateTime
CellStyle.default.withNumFmt(NumFmt.Custom("0.0x")) // any Excel format code
CellStyle.default.bordered                          // thin border, all sides
CellStyle.default.borderTop(BorderStyle.Thin)       // per-side: borderBottom/borderLeft/borderRight, color overloads
ref"B3:F9".outlined(BorderStyle.Medium)             // outline the range edges only (banker box)
```

`range.outlined` is edge-correct (corners get both sides, interior cells untouched) and merges
into existing borders at apply time, preserving each cell's font/fill/format.

## Print and view setup

`SheetView` (gridlines, zoom) and `PageSetup` (orientation, fit, margins, header/footer, print
area, repeat rows) live in `com.tjclp.xl.sheets` and are not part of the prelude export — import
them explicitly:

```scala
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.11.3
import com.tjclp.xl.scripting.{*, given}
import com.tjclp.xl.sheets.{HeaderFooter, PageMargins, PageSetup, SheetView}

val report = Sheet("Report")
  .put(ref"A1", "Quarterly Report")
  .withViewSettings(SheetView(showGridLines = false, zoomScale = Some(90)))
  .withPageSetup(
    PageSetup(
      orientation = Some("landscape"),
      fitToWidth = Some(1),
      // 0.11.1+: HeaderFooter also takes evenHeader/evenFooter/firstHeader/firstFooter
      // with differentOddEven/differentFirst; fitToWidth/Height emit the fitToPage flag
      headerFooter = Some(HeaderFooter(oddFooter = Some("&LACME Corp&RPage &P of &N"))),
      margins = Some(PageMargins(left = 0.5, right = 0.5)),
      printArea = Some(ref"A1:H40"),     // _xlnm.Print_Area defined name
      repeatRows = Some((1, 2))          // rows 1-2 repeat on every printed page
    )
  )

Excel.write(Workbook(report), "/tmp/report.xlsx")
println("wrote print-ready report")
```

Header/footer strings use Excel's codes: `&P` page number, `&N` total pages, `&D` date, `&F`
file name, `&A` sheet name, with `&L`/`&C`/`&R` section markers.

## The `.unsafe` boundary

Everything fallible returns `XLResult[A]` (= `Either[XLError, A]`). The prelude sanctions
exactly one unwrap style — `.unsafe`, which throws a structured `XLException` wrapping the
`XLError`:

```scala
val wb2 = wb.update("Sales", _.put(ref"A1", "x")).unsafe // fail-fast script style
```

Use it **once, at the edge** — compose with `for`-comprehensions in between, or lean on the
total APIs (literal refs, `upsert`, range fill, `readTypedOr`, `recalculate`) so there is
nothing to unwrap.

## `Excel` vs `ExcelIO`

| | `Excel` (sync facade) | `ExcelIO` (cats-effect) |
|---|---|---|
| Style | `Excel.read("in.xlsx")` returns `Workbook`, throws at the IO edge | `ExcelIO.instance[IO].read(path)` returns `IO[Workbook]` |
| For | Scripts, REPL, quick tools | Production services, streaming, resource safety |
| Streaming | — | `readStream`/`writeStream`: `fs2.Stream[F, RowData]`, O(1) memory |

Both are in scope from the prelude. Switch to streaming above ~100k rows — `Excel.read` loads
the whole workbook:

```scala
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.11.3
import com.tjclp.xl.scripting.{*, given}
import cats.effect.IO
import cats.effect.unsafe.implicits.global
import java.nio.file.Paths

val excel = ExcelIO.instance[IO]
val total = excel
  .readStream(Paths.get("huge.xlsx")) // fs2.Stream[IO, RowData], O(1) memory
  .map(_.cells.get(2))                // column C (0-based)
  .collect { case Some(CellValue.Number(n)) => n }
  .compile
  .fold(BigDecimal(0))(_ + _)
  .unsafeRunSync()
println(s"column C total: $total")
```

Streaming writes: `Stream.emits(rows).through(excel.writeStream(path, "Sheet1"))` with
`RowData(rowIndex, Map(colIdx -> CellValue))` (1-based rows, 0-based columns).

## Going further

- [`examples/scripting_tour.sc`](../../examples/scripting_tour.sc) — the canonical runnable tour
  of everything above; [`examples/README.md`](../../examples/README.md) catalogs all example
  scripts.
- [`plugin/skills/xl-scripting/reference/RECIPES.md`](../../plugin/skills/xl-scripting/reference/RECIPES.md)
  — seven complete scripts: bulk transform, typed extraction + validation, model build, workbook
  merge, streaming filter, cell-level diff, CSV ingest.
- [QUICK-START.md](../QUICK-START.md) — the pure-library path (`com.tjclp.xl.{*, given}`).
