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
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.1
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
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.1
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
  replacement (no ZIP corruption on a crashed write). `Excel.modifyR("file.xlsx")(f)` (since
  0.21.0) is the same for an `XLResult`-returning transform — `_.update("Data", …)` needs no
  `.unsafe` inside the lambda, and a `Left` throws *before* anything is written (the file stays
  byte-identical, no scratch file is left behind).
- `Excel.readSheet(path, name)` (since 0.21.0) is `Excel.read` plus the lookup — the whole
  workbook is loaded, then one sheet is selected (stream one sheet of a large file with
  `ExcelIO.readSheetStream`). Since 0.21.1 it returns `XLResult[Sheet]` like the rest of the sync
  surface, so `orExit(Excel.readSheet(path, name))` is the script shape: a missing name is
  `Left(SheetNotFound(name, available))`, whose message lists every available sheet and whose
  `candidates` name the nearest, so `orExit` prints a `did you mean:` line exactly as `xl -s` does;
  a missing, corrupt or over-limit file is `Left` of the reader's own `IOError`/`ParseError`/
  `SecurityError`. (0.21.0 threw an `XLException` instead.) `wb(name)`, `wb.update`, `wb.remove`,
  `wb.rename` and `wb.setSheetState` carry the same candidates (0.21.1).
  `Excel.readMetadata(path)` (since 0.21.0) returns `LightMetadata` — sheet names, visibility,
  dimensions, defined names, the date system — without loading a cell, under the same ZIP-bomb
  limits as `Excel.read`.
- `Excel.write` also accepts an `XLResult[Workbook]` directly.

> **Formulas are written with whatever cache they carry — never `Excel.write` a freshly built
> model.** A `fx"…"` cell has no cached value, so a plain `Excel.write` produces a file whose
> formulas show up **blank** in every cached-value consumer (openpyxl `data_only`, pandas,
> previewers, Excel before its first recalc). Write a model with `Excel.writeChecked` (since
> 0.21.0) — it computes only the formulas that have no cache, keeps every existing cache byte for
> byte, writes, and returns the `RecalcResult` — or with `Excel.writeRecalculated` (since 0.13.0)
> when every formula must be recomputed. Either way failures are visible instead of silent:
>
> ```scala
> val result = Excel.writeChecked(updated, "output.xlsx")       // uncached cells → computed → written
> if !result.isClean then result.errors.foreach(e => println(e.render))
> ```
>
> The file is written even when some formulas fail — errors are data conditions; failed cells
> stay uncached and Excel recalculates them on open. Overloads take an explicit `Clock`
> (deterministic `TODAY`/`NOW`), a `Clock` + `Rng` (reproducible `RAND`/`RANDBETWEEN`), or an
> `XLResult[Workbook]` directly. For full control (e.g. fail-hard pipelines), drop to
> `wb.recalculate()` and write `result.workbook` yourself — see
> [Formulas: build, recalculate, inspect](#formulas-build-recalculate-inspect).

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

The **same split applies to every string-taking form** — `Sheet(name)`, `Workbook(name, …)`,
`sheet.put("A1", v)`, `sheet.style("A1:D1", st)`, `sheet.merge("A1:C1")`,
`sheet.comment("A1", c)` — and it is easier to trip over because there is no `$` at the call site
to warn you. These are `transparent inline`: a string **literal** validates at compile time and
returns `Sheet`/`Workbook`; the very same call with a `val` returns `XLResult[…]`, so a chained
`.put(...)` suddenly type-errors:

```scala
val lit = Sheet("Acquisitions").put("A1", 1)    // : Sheet — both strings are literals
val nm: String = config.sheetName
val cell: String = s"B${row + 1}"
val dyn = Sheet(nm)                             // : XLResult[Sheet]  — the return type changed!
val dyn2 = lit.put(cell, 42)                    // : XLResult[Sheet]  — and again
```

Two rules keep this from ever surprising you:

1. **Literals are total.** A string literal or a `ref"…"`/`fx"…"` literal is checked at compile
   time and the call returns the plain value. Reach for the literal forms whenever the address is
   known when you write the script.
2. **Computed strings use the explicit runtime twins**, which spell `XLResult` in their
   signatures so the `.map`/`.flatMap`/`.unsafe` step reads as intended instead of ambushing the
   chain: **`Sheet.named`** (since 0.18.0), and — since 0.20.0 — **`Workbook.named`**,
   **`sheet.putAt`**, **`sheet.styleAt`**, **`sheet.mergeAt`**, **`sheet.commentAt`**. Validation
   follows the **literal** forms (Excel's sheet-name rules; the same `RefType` parser — corner
   forms only, see below), and the value path is the same code the literal `put` expands to, so
   inferred number formats and style handling do not change.

```scala
// since 0.20.0 (fragment)
val region: String = Seq("North", "East").mkString(" ")
val cell: String = s"B${row + 1}"

val sheet: XLResult[Sheet] =
  for
    s <- Sheet.named(region)                          // InvalidSheetName on a bad name
    a <- s.putAt(cell, total)                         // InvalidCellRef on a bad ref
    b <- a.putAt("C2", BigDecimal("2.50"), currency)  // styled put, same codec merge as the literal
    c <- b.styleAt("A1:C1", header)                   // a cell styles one cell, a range every cell
    d <- c.mergeAt("A1:C1")                           // InvalidRange for a single cell, A:A, $-anchors or garbage
    e <- d.commentAt(cell, Comment.plainText("computed"))
  yield e

val wb: XLResult[Workbook] = Workbook.named("Data", "Summary") // DuplicateSheet on a repeat
```

The twins share one parsing contract: a range where a cell is required is
`Left(InvalidCellRef(ref, "expected a single cell"))`; a **sheet-qualified** ref such as
`"Sales!A1"` is `Left(InvalidReference(…))` — qualify at the workbook instead
(`wb.update(sheetName, _.putAt("A1", v))`); unparseable input is `Left(InvalidCellRef(…))` or
`Left(InvalidRange(…))` naming the offending string.

**Corner forms only.** The twins accept exactly what a *literal* would: `A1` cells and `A1:B2`
two-corner ranges. They do **not** accept full-column/row spellings (`A:A`, `1:1`), `$` anchors
(`$A$1:C3`), or a single cell for `mergeAt` — all of which the *dynamic* branch of the transparent
`merge`/`style` (backed by `CellRange.parse`) happens to accept today. So `sheet.merge(s"$c:$c")`
must not be rewritten as `mergeAt(s"$c:$c")` (that is `Left(InvalidRange)`); keep the existing
parse-then-typed escape hatch for those spellings — it also spells the `XLResult`:

```scala
val s: String = s"$c:$c"                 // "D:D", "1:1", "$A$1:C3" and plain "A1" all parse
s.asRange.map(sheet.merge)               // XLResult[Sheet] — String.asRange is CellRange.parse-backed
s.asRange.map(r => sheet.style(r, header))
cell.asCell.map(r => sheet.put(r, total)) // String.asCell: A1 cells (ARef.parse; no $ anchors — use asRange)
```

`Sheet.named` alone (0.18.0+):

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.1
import com.tjclp.xl.scripting.{*, given}

val region = Seq("North", "East").mkString(" ")                     // runtime name
val sheet = Sheet.named(region).map(_.put(ref"A1", "ready")).unsafe // XLResult, spelled out
Excel.write(Workbook(sheet), "/tmp/named.xlsx")
```

On ≤0.19.x (no `putAt`/`styleAt`/`mergeAt`/`commentAt`/`Workbook.named`), make the union explicit
at the call site with an ascription — `val s: XLResult[Sheet] = sheet.put(cell, 42)` — or parse
once with `RefType.parse(cell)` and use the typed `ARef`/`CellRange` overloads.

Prefer **total navigation** over interpolated refs in loops — no `Either` at all:

```scala
val base = ref"A2"
base.down(2)     // A4   (default step is 1: base.down() == A3)
base.right(1)    // B2
base.up(1)       // A1
base.left(1)     // out of bounds! see below
base.shift(1, 2) // B4   (colOffset, rowOffset)
```

`shift`/`down`/`up`/`left`/`right` are total but **unchecked at the sheet edges**: `ref"A1".up()`
produces the non-existent "A0", which corrupts output if written. Since 0.20.0 the **bounded
navigation** forms make the edge explicit — `None` past it, or a clamp onto it — so a loop can
stop cleanly instead of minting an invalid ref:

```scala
// since 0.20.0
ref"A1".tryDown(1)            // Some(A2)
ref"A1".tryShift(-1, 0)       // None — would be column -1
ref"XFD1".tryRight(1)         // None — past the last column
ref"A1048576".tryDown(1)      // None — past the last row
ref"C3".clampShift(-10, 5)    // A8  — column pinned to A, row shifted
ref"A1".clampShift(-3, -3)    // A1  — already at the corner

// Bounded steps compose with the patch DSL without an Either in the loop body:
val patch = ref"A1".tryDown(2).fold(Patch.empty)(_ := "third row")
```

`tryShift(dc, dr)` agrees with `shift(dc, dr)` whenever it is `Some`, and
`tryShift(dc, dr).flatMap(_.tryShift(-dc, -dr))` is `Some(ref)` in bounds. A range can also be
walked by **slices** instead of by interpolated corners (since 0.20.0):

```scala
// since 0.20.0
val table = ref"A1:D10"
table.rows.size                 // 10 — one-row-high CellRanges, top to bottom (lazy)
table.columns.map(_.toA1).toList // List("A1:A10", "B1:B10", "C1:C10", "D1:D10")
table.row(0)                    // Some(A1:D1)  — 0-based within the range
table.row(10)                   // None         — outside 0 until height
table.column(3)                 // Some(D1:D10)
```

For **runtime column handles** (since 0.13.0) — column-oriented builders that fold over letters
computed at runtime — use `Column.parse` instead of special-casing macro literals; a runtime
`RefType` also exposes `.col` (the cell's column, or the range's starting column):

```scala
Vector("C" -> 14.0, "D" -> 22.0).foldLeft(sheet) { case (s, (letter, w)) =>
  val col = Column.parse(letter).getOrElse(sys.error(s"bad column: $letter"))
  s.setColumnProperties(col, ColumnProperties(width = Some(w)))
}
Column.parse("D1")                      // Right(D) — trailing row digits tolerated
RefType.parse("Sales!C2:E9").map(_.col) // Right(C) — starting column of the range
```

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

## The Edit algebra: `wb.edit` / `sheet.edit` (since 0.21.0)

A `Patch` is sheet-local and formula-blind. `Edit` (`com.tjclp.xl.ops.Edit`, on the prelude
surface) is the **workbook-level** operation vocabulary — one case per batch op and mutating CLI
verb, 49 in all — with one interpreter, `Edit.applyAll`, behind `wb.edit(...)` /
`sheet.edit(...)`. It says what a `Patch` cannot (formula dragging, structural edits, sheet
management).

What a script shares with the CLI at 0.21.0 is the **kernels**, not the interpreter: `xl batch`
and the mutating verbs still run their own path (xl-cli's `OpRegistry` → `BatchParser` → the
command handlers) and nothing in xl-cli calls `Edit.applyAll` yet, but both sides call the same
library methods — `Sheet.fill` / `copyRange` / `sort` / `clearRange` / `groupRows` / `autoFit`,
`StructuralEditor` for the structural four, `SheetRenamer` for renames — so `wb.edit` and the verb
make the same change to the same cells. Lowering the CLI onto `Edit.applyAll` is [#583](https://github.com/TJC-LP/xl/issues/583) (W2.2);
until then two divergences are documented: `MoveSheet.toIndex` (below) and `Copy.target`, a single
cell (`Loc`) where batch `copy` also accepts a range target.

**Targets carry their own sheet.** `Loc(sheet: Option[SheetName], ref: ARef)` is one cell,
`Area(sheet, range: CellRange)` a range, and the row/column edits take `sheet: Option[SheetName]`
directly with a `RowSpan` / `ColSpan`. `None` means *the scope's default sheet* — THE sheet rule,
the same one the CLI applies: a qualifier wins, then the scope's default, then the only sheet of a
single-sheet book, else `SheetRequired`. `Loc.parse("'Q1 Data'!B7")`, `Area.parse("A1:B2")`,
`ColSpan.parse("E:H")` and `RowSpan.parse("10:20")` read the CLI spellings; `Area.cell(loc)` is a
1x1 area.

| Group | Cases |
|-------|-------|
| Cell content | `Put(at: Loc, value: CellValue, format: Option[FormatHint])` — build it with **`Edit.put(loc, a)`** / **`Edit.put(ref, a)`** so a `LocalDate`/`BigDecimal` keeps its codec format; `PutValues(at: Area, values: Vector[CellValue], format)` row-major; `PutFormula(at: Loc, formula: String, format)`; `PutFormulas(at: Area, formulas, format)` one per cell, as written; **`DragFormula(at: Area, formula, anchor: ARef, format)`** shifts relative refs from `anchor` like fill-down; **`Fill(source: Area, target: CellRange, direction: Edit.FillDir)`** (`Down` / `Right`); `Copy(source: Area, target: Loc, valuesOnly)` (either side may name another sheet); `Sort(at: Area, keys: Vector[Edit.SortKeySpec], hasHeader)` (`SortKeySpec.ascending(col)` / `.descending(col)`); `Clear(at: Area, what: ClearWhat)` (`ClearWhat.contents` / `.styles` / `.comments` / `.all`) |
| Style & layout | `Style(at: Area, overlay: StyleOverlay, mode: StyleMode.Merge \| Replace)`; `Merge(at)`, `Unmerge(at)`; `ColWidth(sheet, cols: ColSpan, width)`, `RowHeight(sheet, rows: RowSpan, height)`; `HideCols`/`ShowCols(sheet, cols)`, `HideRows`/`ShowRows(sheet, rows)`; `AutoFit(sheet, cols: Option[ColSpan])` (`None` = every used column; a column is fitted to what it DISPLAYS — a formula cell by its cached value, an uncached formula contributing nothing, so fit after `Excel.writeChecked` / `recalculate` or accept the values' widths (#613)) |
| Outline | `GroupRows(sheet, rows, level, collapsed)`, `GroupCols(sheet, cols, level, collapsed)` (level 1-7); `UngroupRows(sheet, rows)`, `UngroupCols(sheet, cols)` |
| Annotations & objects | `SetComment(at: Loc, comment: Comment)`, `RemoveComment(at)`; `Hyperlink(at, target: Option[String])` (`None` clears); `AddConditionalFormat(sheet, ranges: Vector[CellRange], rules: Vector[CfRule])`; `AddChart(sheet, chart, anchor: DrawingAnchor)`; `AddImage(sheet, image: ImageData, anchor)` |
| Sheet view & print | `Freeze(at: Loc)`, `Unfreeze(sheet)`; `SetSheetView(sheet, gridlines, zoom, tabSelected)`; `SetTabColor(sheet, color: Option[Color])`; `SetAutoFilter(sheet, range: Option[CellRange])`; `SetPageSetup(sheet, orientation, scale, fitToWidth, fitToHeight, fitToPage)`; `SetHeaderFooter(sheet, oddHeader, oddFooter, evenHeader, evenFooter, firstHeader, firstFooter, differentOddEven, differentFirst)` |
| Structure | `InsertRows(sheet, at: Row, count)`, `DeleteRows(sheet, at: Row, count)`, `InsertCols(sheet, at: Column, count)`, `DeleteCols(sheet, at: Column, count)` — references on every sheet are rewritten through the evaluator (`#REF!` on loss) |
| Workbook | `AddSheet(name, after, before)` (at most one of the two); `RemoveSheet(name)`; `RenameSheet(from, to)` (rewrites every reference, and retargets the scope when it renames the default sheet); `MoveSheet(name, toIndex, after, before)` (exactly one; `toIndex` is the sheet's FINAL 0-based position, `0` to `sheetCount - 1`, refused outside — **not** what `xl move-sheet --to N` does today: the verb reads `N` against the order BEFORE the sheet is removed and clamps, so on `[A, B, C]` the verb's `A --to 2` gives `[B, A, C]` where this edit gives `[B, C, A]`; [#583](https://github.com/TJC-LP/xl/issues/583) reconciles them); `CopySheet(source, target)`; `HideSheet(name, veryHidden)`, `ShowSheet(name)`; `DefineName(name, refersTo, scope)`, `RemoveName(name, scope)` |

Two value types ride along. **`FormatHint`** says where a number format came from: `Inferred(fmt)`
(a codec's hint — applied only when the cell's format is General, exactly as `Sheet.put` does)
and `Explicit(fmt)` (the caller asked — replaces the number format, keeps font/fill/border).
**`StyleOverlay`** is a *partial* style — every field `Option`, so "un-bold" is expressible —
with a right-biased `++` (`StyleOverlay.empty` is the identity) and `StyleOverlay.of(style)` /
`ofBorder(border)` to lift a full style; `StyleMode.Merge` overlays onto each cell's current
style, `Replace` applies it to `CellStyle.default`.

**Entry points** — all take the prelude's `given FormulaSupport` (the evaluator's
`EvalFormulaSupport`, stated explicitly in the prelude because wildcard exports skip givens):

| Call | Returns | Default sheet |
|------|---------|---------------|
| `wb.edit(edits*)` | `XLResult[Workbook]` | none — a qualified target names its sheet; an unqualified one resolves only on a single-sheet book, else `SheetRequired` |
| `wb.editIn(defaultSheet)(edits*)` | `XLResult[Workbook]` | `defaultSheet` for every unqualified target — the CLI's `-s` |
| `sheet.edit(edits*)` | `XLResult[Sheet]` | the sheet itself (a one-sheet workbook under the hood); an edit naming another sheet is `SheetNotFound`; a rename of this sheet is followed |
| `Edit.applyAll(wb, Vector[Edit], scope: EditScope)` | `XLResult[Applied]` | explicit — `Applied(workbook, planned, scope)` is the edited book, one `Planned(index, edit, sheet, touched)` row per edit and the scope *after* the sequence; `applied.touchedBySheet` seeds an after-edit recalculation cone, `applied.structural` says the whole book needs one |
| `Edit.plan(wb, edits, scope)` | `XLResult[Vector[Planned]]` | the same fold with the workbook discarded — a **semantic** dry-run that fails exactly where `applyAll` would (and costs the same) |
| `Edit.validate(edit)` | `XLResult[Unit]` | the **static**, workbook-free checks of one edit: counts match the area, spans and levels are in range, a formula parses, `MoveSheet` names exactly one destination, … — a refusal is `XLError.InvalidArgument(op, reason)`, code `INVALID_ARGUMENT` (since 0.21.1, [#617](https://github.com/TJC-LP/xl/issues/617)), so a script can branch on it |
| `Edit.lower(edit, sheet)` / `Patch.toEdits(patch, sheet)` | `Option[Patch]` / `Option[Vector[Edit]]` | the bridge to the sheet-local kernel, `None` when the edit needs formula support, another sheet or the workbook |

The scope type is **`EditScope`** on the prelude and pure surfaces (its source name, `ops.Scope`,
collides with JMH's and ZIO's `Scope` in files that import both): `EditScope.none`,
`EditScope.of(sheet)`, and `scope.after(edit)` moves the default when a `RenameSheet` renames it.

**Fail-fast, all-or-nothing.** The edits apply in order; the first failure ends the run with
`Left(XLError.EditFailed(index, op, cause))` — `index` is the **1-based** position of the failing
edit, `op` its kebab name (`drag-formula`, `delete-rows`, …), `cause` the underlying error — and
the input workbook is never partially written. `err.opIndex` is `Some(index)`, `err.root` the
innermost cause (an `EditFailed` unwrapped), and `err.code` / `err.hint` / `err.candidates` are
the cause's, so a batch failure is classified by what went wrong, not by where; `err.message`
reads `op 2 (drag-formula): …`. It is the same 1-based position `xl batch` reports as
`location.opIndex` on a `BATCH_OP_FAILED`, and `orExit` prints it as the CLI's stderr diagnostic
(`Error: …`, `code:`, `hint:`). `EditSchema` is the algebra's own metadata — one `EditSpec` per
`Edit` case, 49 rows (`EditSchema.all`, `EditSchema.find("putf")`, `EditSchema.nameOf(edit)`),
each naming its `batchOp` and `cliVerb`. It has the shape of the batch schema but is not the
document `xl batch --schema` prints: that is xl-cli's `OpRegistry` (32 ops), a subset, and the two
have already drifted in detail (`copy.target` is a single cell in `EditSchema`, a cell or range in
the batch schema).

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.1
import com.tjclp.xl.scripting.{*, given}

val Data = SheetName.unsafe("Data")
val Summary = SheetName.unsafe("Summary")
val book = Workbook(
  Sheet(Data)
    .put(ref"A1", "Region").put(ref"B1", "Units").put(ref"C1", "Price")
    .put(ref"A2", "North").put(ref"B2", 12).put(ref"C2", BigDecimal("9.50"))
    .put(ref"A3", "South").put(ref"B3", 7).put(ref"C3", BigDecimal("11.25"))
)

// editIn(Data): every target with sheet = None lands on Data; a Some(...) names its own sheet.
val edited = orExit(
  book.editIn(Data)(
    Edit.put(ref"D1", "Total"),                                            // codec inference kept
    Edit.DragFormula(Area(None, ref"D2:D3"), "=B2*C2", ref"D2", None),     // D3 becomes =B3*C3
    Edit.PutFormula(Loc(None, ref"D4"), "=SUM(D2:D3)", Some(FormatHint.Explicit(NumFmt.Currency))),
    Edit.Style(Area(None, ref"A1:D1"), StyleOverlay(bold = Some(true)), StyleMode.Merge),
    Edit.AutoFit(None, None),                                              // every used column
    Edit.Freeze(Loc(None, ref"A2")),
    Edit.AddSheet(Summary, after = Some(Data), before = None),
    Edit.PutFormula(Loc(Some(Summary), ref"B2"), "=Data!D4", None)         // qualified: not the default
  )
) // a Left prints "Error: op N (<op>): …" + code/hint on stderr — the CLI's diagnostic — and exits 1

val result = Excel.writeChecked(edited, "/tmp/edited.xlsx") // caches D2:D4 and Summary!B2, writes
println(s"wrote ${edited.sheets.size} sheets; clean: ${result.isClean}")
```

`Edit.plan` and `Edit.applyAll` expose the fold itself; `Edit.validate` is the static check:

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.1
import com.tjclp.xl.scripting.{*, given}

val Data = SheetName.unsafe("Data")
val book = Workbook(Sheet(Data).put(ref"A1", 10).put(ref"A2", 20).put(ref"B1", fx"=A1*2"))

val edits = Vector(
  Edit.Fill(Area.cell(Loc(None, ref"B1")), ref"B1:B2", Edit.FillDir.Down), // B2 = A2*2
  Edit.InsertRows(None, Row.from1(1), 1),                                  // every formula shifts
  Edit.RenameSheet(Data, SheetName.unsafe("Model"))                        // the scope follows
)

edits.foreach(e => orExit(Edit.validate(e)))                       // static, no workbook needed

val planned = orExit(Edit.plan(book, edits, EditScope.of(Data)))   // semantic dry-run
planned.foreach { p =>
  println(s"${p.index} ${EditSchema.nameOf(p.edit)} on ${p.sheet.map(_.value)} touched ${p.touched.map(_.toA1)}")
}

val applied = orExit(Edit.applyAll(book, edits, EditScope.of(Data)))
println(s"default sheet now ${applied.scope.defaultSheet.map(_.value)}; structural = ${applied.structural}")

// All-or-nothing: the failing edit names its 1-based position; `book` is untouched
book.edit(
  Edit.Merge(Area(None, ref"A1:B1")),
  Edit.DeleteRows(Some(SheetName.unsafe("Nope")), Row.from1(1), 1)
) match
  case Left(err) => println(s"${err.opIndex} ${err.code}: ${err.message}") // Some(2) SHEET_NOT_FOUND: op 2 (delete-rows): …
  case Right(_) => println("unexpected")
```

Formula-aware edits (`DragFormula`, `Fill`, the structural four, `RenameSheet`) need a parser,
which xl-core does not have: the prelude supplies xl-evaluator's `EvalFormulaSupport` as the
`given`, so scripts never see the seam. `FormulaSupport.textOnly` is the explicit opt-out for a
text-only interpreter — it stores formulas as written and **refuses** those edits with
`UnsupportedCapability` rather than write a silent `#REF!`
(`book.edit(drag)(using FormulaSupport.textOnly)`).

## Formulas: build, recalculate, inspect

`wb.recalculate()` is a **total** whole-workbook recalculation: every formula on every sheet
evaluates in dependency order, cross-sheet references resolve automatically, and failures never
throw — they are collected per cell.

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.1
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

Since 0.12.1, conditional formatting is typed: `sheet.conditionalFormat(range, CfRule.cellIs(...))`
authors cellIs/expression/colorScale/dataBar/top10 rules with `Dxf` differential formats;
structural edits shift rule ranges, and rule families xl does not model survive round-trips
byte-faithfully.

Since 0.12.0 the prelude also exposes the drawing layer: `sheet.addImage(bytes, format, at)`
embeds pictures (7 formats, natural-size PNG/JPEG sniffing) and `com.tjclp.xl.charts.Chart`
authors bar/line/pie charts anchored to ranges — both round-trip through OOXML with unmodeled
content preserved byte-faithfully.

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
| `errors` | `Vector[CellEvalError]` — per-cell **host failures** (parse errors, missing sheets, cycle participants, cells blocked by a cycle). Since 0.14.0, Excel **error values** (`#DIV/0!`, `#N/A`, …) are *results*, not failures — they cache like any value and do not appear here |
| `excelErrors` | (0.14.0) `Vector[(SheetName, ARef, CellError)]` — cells whose cached result is an Excel error value, sorted; inspect when you want to surface `#DIV/0!`s without treating them as host failures |
| `isClean` | `true` when `errors.isEmpty` — a workbook full of cached `#DIV/0!`s is "clean" (the recalculation succeeded; the errors are data) |
| `toEither` | `Right(workbook)` when clean, `Left(errors)` otherwise — for fail-hard pipelines |
| `converged` | (0.20.0) `cycles.forall(_.converged)` — `false` iff some cyclic component exhausted `maxIter` without every member's \|Δ\| dropping below `maxChange`. The last-round values are kept (Excel semantics, `errors` stays empty), so gate on this after any large circular perturbation. Non-iterative runs report `true` |
| `iterationsUsed` | (0.20.0) rounds run by the WORST component: `0` when no iteration happened, `maxIter` when any component exhausted, otherwise the round it converged on |
| `cycles` | (0.20.0) `Vector[SccReport]` — one verdict per cyclic strongly-connected component actually iterated (`members`, `converged`, `rounds`, `maxDelta`, plus `render`), sorted by the component's minimum member. Empty on non-iterative and acyclic runs |
| `unconverged` | (0.20.0) `cycles.filterNot(_.converged)` — the offenders to name in a report |
| `certified` | (0.20.0) `errors.isEmpty && converged` — the single gate meaning "this workbook is at its global fixpoint" |

Reference cycles are **isolated**: the participants and their downstream dependents are reported
(e.g. `Model!A7: Formula error in '=B7': Circular reference` via `CellEvalError.render`) while
the acyclic remainder still evaluates and caches.

### One primitive: `recalculate(options)` (since 0.20.0)

The clock/rng/iterative/parallel overloads above remain, but they are forwarders for one
options-driven primitive. `RecalcOptions` carries every knob with the value the zero-argument
`recalculate()` uses, so `wb.recalculate(RecalcOptions())` is `wb.recalculate()` on an acyclic
book, byte for byte:

```scala
// since 0.20.0 — fragment, not a runnable script
val opts = RecalcOptions(
  clock = Clock.fixedDate(java.time.LocalDate.of(2026, 1, 31)),
  rng = Rng.seeded(42L),
  iterative = IterativeMode.FromCalcPr, // honour <calcPr iterate>; Off isolates cycles; Force(calc) iterates regardless
  parallelism = 4,                      // wave-parallel independent regions (a declared iteration wins and runs sequentially)
  seedTables = false                    // true also seeds data-table interiors afterwards
)
val result = wb.recalculate(opts)
println(result.summary)                // exactly the line `xl recalc` prints: "Recalculated 12 formulas"
```

Three entry points share the record — no default arguments, because extension methods reached
through the prelude's wildcard export cannot carry them:

| Method | Computes | Leaves alone |
|--------|----------|--------------|
| `wb.recalculate(opts)` | every formula on every sheet | — |
| `wb.recalculateAfterEdit(sheet, refs, opts)` | the edited cells and everything that depends on them (plus dynamic `INDIRECT`/`OFFSET` readers) — what the CLI's `put`/`putf` do after a write; falls back to a full pass when the book iterates | every other cache, byte-identical; unaffected volatile cells never touch the clock |
| `wb.recalculateUncached(opts)` | only formulas with **no** cached value, in dependency order, reading their inputs' caches as they are | every cached cell, even when its cache is wrong (the caches-are-truth doctrine); uncached cycle members are reported, not guessed |

`RecalcResult.summary` renders the result as the CLI does — formula count, `(N error values)`,
the first three failures, and the iterative verdict — so a script and `xl recalc` report the same
thing on the same file.

### Renaming a sheet rewrites its references (since 0.20.0)

`Workbook.rename` is deliberately formula-blind (xl-core has no parser): it changes the tab and
leaves `Sheet1!A1` in every other formula — a file that lints clean and opens in Excel as `#REF!`
([#559](https://github.com/TJC-LP/xl/issues/559)). `SheetRenamer.rename` is the rename that
follows through:

```scala
// since 0.20.0 — fragment, not a runnable script
val renamed: XLResult[Workbook] =
  SheetRenamer.rename(wb, SheetName.unsafe("Sheet1"), SheetName.unsafe("Q1 Data"))
// Sheet2!A1  =Sheet1!A1*2        → ='Q1 Data'!A1*2   (quoted because the name needs it)
// name Total Sheet1!$A$1         → 'Q1 Data'!$A$1     (comma unions rewritten segment by segment)
// CF  Expression("Sheet1!A1>0")  → 'Q1 Data'!A1>0     (CellIs, Expression, Cfvo.Formula)
// DV  List("Sheet1!$A$1:$A$3")   → 'Q1 Data'!$A$1:$A$3
SheetRenamer.references(wb, SheetName.unsafe("Sheet1")) // Vector[QualifiedRef]: the cells it would touch
```

Cached values and formula record kinds are preserved (a rename changes no value); a string literal
that spells the name, an external-workbook reference (`[2]Sheet1!A1`) and a sibling whose name
merely contains it (`Sheet10`) are untouched; a dependent text that mentions the sheet but cannot
be parsed refuses the whole rename (`Left(FormulaError)`) with the workbook untouched. Since 0.21.0
([#608](https://github.com/TJC-LP/xl/issues/608)) `SheetRenamer.renameLocated(wb, from, to)` is the
same rename whose `Left` is a `Refusal(site, error)` naming where that text lives — `Site.Cell(sheet,
ref)`, `ConditionalFormat(sheet)`, `DataValidation(sheet)` or `Name(name)`, spelled by `site.describe`
as `Summary!I23` — with `site = None` for `Workbook.rename`'s own `SheetNotFound`/`DuplicateSheet`;
`rename` is `renameLocated` with the site dropped. Every changed
sheet goes back through `Workbook.put`, so a workbook read from disk marks exactly the rewritten
sheets modified; because a rename also changes `workbook.xml`, the writer regenerates every
worksheet part deterministically, and sheets that never mentioned the old name come out
byte-identical. `Preserved` CF/DV/chart payloads, hyperlink locations, `_xlfn.`-prefixed functions
and 3-D ranges are not rewritten (see LIMITATIONS.md; the last two refuse the rename). The same engine is exposed string-in/string-out as
`FormulaOps.renameSheet(text, from, to)`, `FormulaOps.shift(text, dc, dr)` and
`FormulaOps.mentionsSheet(text, sheet)`, and the structural editor (`StructuralEditor.insertRowsChecked`
and friends) is reachable from the prelude too.

Since 0.13.0, **circular models are opt-in** rather than always errors: pass an `IterativeCalc` to
fixpoint declared cycles instead —
`wb.recalculate(IterativeCalc(maxIter = 100, maxChange = BigDecimal("0.001")))` runs Jacobi
iteration (each member reads previous-iteration values until every |Δ| < `maxChange`
or `maxIter` rounds; non-convergence keeps the last values with no error, per Excel). Plain
`recalculate()` still isolates cycles. Honor a file's own settings with
`wb.metadata.calcPr.filter(_.iterativeCalculation).map(IterativeCalc.fromCalcPr)`, and author them
on scratch builds with `wb.withCalcPr(CalcPr(iterativeCalculation = true, maxIterations = Some(100),
maxChange = Some(BigDecimal("0.001"))))` (emits `<calcPr iterate iterateCount iterateDelta/>`).

Since 0.20.0 an iterative recalculation walks the **SCC condensation** of the workbook graph once
in dependency-first order — a run of acyclic cells evaluates, then each cyclic component fixpoints
against those freshly computed values, and so on. Consequences worth knowing: `maxIter`/`maxChange`
are **per component** (one permanently-oscillating cycle no longer burns an unrelated cycle's
budget, and `cycles` names the offender); one pass reaches the workbook's **global** fixpoint, so
`recalculate(IterativeCalc)` is idempotent on a converged book and re-solving a cached circular
book is safe (it was not before 0.20.0); and one iterative recalculation is one volatile
generation — `TODAY()`/`NOW()` agree inside the fixpoints and in the acyclic cells between them.
Dynamic (INDIRECT/OFFSET) cycles are still invisible to Tarjan and are not covered by `converged`.

Also since 0.20.0, cycle members **warm-start from their loaded cached number** (0 for every other
shape, and as the fallback), matching Excel — a book already at its fixpoint re-solves to itself in
one round instead of being driven back through the 0-seed transient. Pass
`IterativeCalc(maxIter, maxChange, seedFromCaches = false)` for a cold start when a book's numeric
caches are known to be poisoned. Two consequences to keep in mind: for a circular book,
`recalculate(wb)` and `recalculate(wb with caches stripped)` are no longer guaranteed to agree on
a nonlinear cycle with several fixpoints; and a member whose cache is a stale **error or text**
value seeds 0 rather than itself, so such a cycle still heals (seeding it would wedge the cycle at
its own poison, since arithmetic propagates both shapes unchanged).

Also since 0.13.0, **defined names resolve** in formulas: `=IF(case=2,…)`,
`=entry_mult*ltm_ebitda`, and `=SUM(rev_range)` evaluate against workbook- and sheet-scoped names
(sheet-scoped shadows global), contribute dependency edges so `recalculate()` orders name-gated
families correctly, and round-trip byte-faithfully; unresolvable names are clean per-cell errors.

When the very next step is a write, `Excel.writeChecked(wb, path)` (since 0.21.0) fills in only
the uncached formulas (`recalculateUncached`) and writes, and `Excel.writeRecalculated(wb, path)`
(since 0.13.0) recalculates everything and writes — both write the cached workbook even on partial
failure and return the same `RecalcResult`; both take a `RecalcOptions` (since 0.21.0). Use the
explicit `recalculate().toEither` pattern above when a dirty result must abort *before* anything
lands on disk.

For one-off questions, `wb.evaluateFormula("=SUM(Data!A1:A9)", "Summary")` returns
`XLResult[CellValue]` with cross-sheet context wired automatically (108 functions supported —
see the [skill API reference](../../plugin/skills/xl-scripting/reference/API.md) for the full
list).

### Inspect a workbook: `wb.describe`, `wb.audit`, `QualifiedGraph` (since 0.20.0)

The same analyses `xl describe --full`, `xl audit` and `xl deps` print are values a script can
branch on — all pure, all total:

```scala
// since 0.20.0 — fragment, not a runnable script
val summary: WorkbookSummary = wb.describe          // one SheetSummary per sheet + names, date1904, calcPr
summary.sheets.filter(_.uncachedFormulas > 0).map(_.name.value)

val audit: WorkbookAudit = wb.audit                 // buckets, each in workbook order (sheet, row, column)
if !audit.isClean then                              // error cells, uncached/unparseable formulas, cycles, unresolved names
  audit.errorCells.foreach((ref, err) => println(s"$ref ${err.toExcel}"))
audit.volatile                                      // TODAY/NOW/RAND/RANDBETWEEN cells: a note, not a finding
audit.restrictTo(SheetName.unsafe("Summary"))       // what `xl audit -s Summary` reports

val graph = QualifiedGraph.of(wb)                   // bounded: a full-column reader expands only to occupied cells
val b4 = QualifiedRef(SheetName.unsafe("Summary"), ref"B4")
graph.precedents(b4, 2)                             // Vector of layers: exactly 1 hop, exactly 2 hops
graph.dependents(b4, 0)                             // 0 = every layer; an empty cell inside a summed range still names the sum
graph.sccs.filter(_.cyclic)                         // the circular references
```

`SheetSummary` carries `cellCount`, `formulaCount`, `uncachedFormulas`, `mergedRanges`, `comments`,
`hyperlinks`, `freeze`, `tabColor`, `autoFilter`, `tables`, `charts`, `pictures`,
`conditionalFormats`, `dataValidations`, `hiddenRows`, `hiddenCols` plus `state` and `dimension`.
`QualifiedGraph.precedentsOf`/`dependentsOf` are the single-hop sets; the `dependencies` map is the
forward graph and `rangeReaders` the symbolic range index behind the reverse question.

## Typed extraction

```scala
sheet.readTyped[BigDecimal](ref"C2")       // Either[CodecError, Option[BigDecimal]]
sheet.readTypedOr[Int](ref"B2", 0)         // total, with default
sheet.readTypedOpt[String](ref"A2")        // flat Option — mismatch and empty both None
sheet.readTypedStrict[BigDecimal](ref"C2") // like readTyped, but ANY formula cell is a TypeMismatch
```

Nine codec types: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime,
RichText. Use `readTyped` when you must distinguish a type mismatch from an empty cell;
`readTypedOr`/`readTypedOpt` when you just need a value.

**Formula cells read through their cached value** ([GH-477](https://github.com/TJC-LP/xl/issues/477)).
After `recalculate()`, `writeRecalculated`, or `Excel.read` of a book Excel saved, `B1` holds
`Formula("A1*3", Some(Number(6)), Normal())` and `readTyped[BigDecimal](ref"B1")` is
`Right(Some(6))`, `readTypedOpt` is `Some(6)` — the same value `view`/`eval` show, with no manual
`CellValue.Formula(_, Some(v), _)` unwrapping. A formula that has not been recalculated yet
(`fx"=A1*3"` straight after `put`) has no cache and therefore nothing to read: `readTyped` is
`Left(TypeMismatch(expected, formula))`, `readTypedOpt` is `None`, `readTypedOr` is the default.
`readTypedStrict` is the escape hatch that rejects *every* formula cell, cached or not — reach for
it when "is this a formula?" matters more than its result (auditing hand-entered constants, refusing
a cache that may be stale). The same see-through rule is available for hand-written matches as
`cell.effectiveValue` (and `cell.isUncachedFormula`).

## Records: `derives RowCodec` (since 0.21.0)

A case class is a row. Derive a `RowCodec` and the sheet reads and writes records directly —
field order is column order, field names are the header row, `Option[T]` fields are empty
cells — no per-cell `readTyped` loops:

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.1
import com.tjclp.xl.scripting.{*, given}
import java.time.LocalDate

final case class Order(id: Int, customer: String, qty: Int, price: BigDecimal, shipped: Option[LocalDate])
  derives RowCodec

val orders = Vector(
  Order(1, "Acme", 3, BigDecimal("9.99"), Some(LocalDate.of(2026, 1, 15))),
  Order(2, "Globex", 1, BigDecimal("120.00"), None)
)

// Write: a header row of field names at A1, one row per record below it
val placed = Sheet("Orders").putRowsWithHeader(ref"A1", orders).unsafe
val headerRange = placed.headerRange                       // Some(A1:E1)
val dataRange = placed.dataRange                           // Some(A2:E3); None when `orders` is empty
val styled = placed.sheet.style(ref"A1:E1", CellStyle.default.bold)

// Read back — by header (column order free, extra columns ignored) or by position
val byHeader: Either[RowCodecError, Vector[Order]] = styled.readRowsByHeader[Order](Row.from1(1))
val byRange: Either[RowCodecError, Vector[Order]] = styled.readRows[Order](ref"A2:E3")

// An Excel table over header + records, named and columned after the record
val table = Sheet("Orders").putTable(ref"A1", orders, "Orders").unsafe
Excel.write(Workbook(table.sheet), "/tmp/orders.xlsx")
println(s"${byHeader.map(_.size)} records; table ${table.sheet.getTable("Orders").map(_.range.toA1)}")
```

The rules, all of them:

- **Field types**: the nine codec types (String, Int, Long, Double, BigDecimal, Boolean,
  LocalDate, LocalDateTime, RichText) and `Option` of each. A field of any other type is a compile
  error naming the missing `CellCodec`; add a `given CellCodec[T]` and it flows into records too.
  A record needs at least one field.
- **Writing**: `putRows(at, records)` writes records only (append under a header you styled
  yourself); `putRowsWithHeader(at, records)` writes the field names at `at` and records below;
  `putTable(at, records, name)` adds an Excel table over header + records (`name`: letters, digits,
  `_`; it doubles as the display name; with no records the table keeps Excel's one blank data
  row). All three return `XLResult[RowsPlaced]` — `sheet`, `headerRange`, `dataRange`, `range`
  (header ∪ data), `count` — and are `OutOfBounds` when the block would run past column XFD or
  row 1048576. Codec format hints (Decimal, Date, DateTime) register as styles and merge into an
  existing cell style exactly as `put` does (the existing style wins; only a General number
  format is filled in); a `None` field leaves its cell empty and never creates one. Only the
  records' cells are written: rewriting a shorter block over a longer one leaves the rows below
  it in place, so clear the old block before regenerating a table in place.
- **Reading by position**: `readRows[A](range)` decodes one record per row of `range`, whose
  width must equal the record's (`RowCodecError.Width` otherwise). Every row is a record: a blank
  row is `Missing` unless every field is an `Option`.
- An `Option[T]` field is `None` only for an absent or `CellValue.Empty` cell. A cell holding the
  empty string — SheetJS and some exporters write `<v></v>` text cells where a person would leave
  a blank — is `Some("")` for `Option[String]` and a `TypeMismatch` for `Option[Int]`, because
  Excel distinguishes `""` from blank (`ISBLANK` is FALSE, `COUNTA` counts it). Normalise with
  `.filter(_.nonEmpty)`, or clear such cells before reading (#617).
- **Reading by header**: `readRowsByHeader[A](headerRow)` finds each field's column through
  `sheet.columnOf(field, headerRow)` — an exact header match wins, otherwise the match ignoring
  case, whitespace, `_` and `-` (`"Order ID"`, `order_id`, `orderId` agree), leftmost on ties —
  then reads the contiguous block under the header and stops at the first row whose record cells
  are all empty (Excel's current region), so a totals row after a blank line is not a record.
  `sheet.columnHeaders(row)` lists `(Column, text)` pairs for discovery.
- **Errors** (`Either[RowCodecError, Vector[A]]`, first failing cell in row-major order):
  `Field(row, column, field, cause)` for a value the field's codec rejected, `Missing(row, column,
  field)` for a required field on an empty cell, `HeaderNotFound(header, headerRow, available)`,
  `Width(expected, actual)`. `.message` is one line, cell first (`C2 (qty): expected Int, got
  Text(three)`); `.toXLError` bridges into `XLResult`.
- **Formulas** read through their cached value, like every typed read (GH-477): a recalculated
  or Excel-saved formula decodes as its result, an uncached one is a `Field` error naming the
  formula. An error cell (`#N/A`, `#DIV/0!`, …) is a `Field` error too, even under an `Option`
  field — only an empty cell is `None` — so a stray `#N/A` fails the whole read at that cell.
- **Law** (pinned in `RowCodecSpec` over generators): `readRows(putRows(at, rows).dataRange) ==
  Right(rows)` and `readRowsByHeader(putRowsWithHeader(at, rows).headerRow) == Right(rows)` for
  every codec type, required and optional.

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
CellStyle.default.withUnderline(Underline.Double)   // since 0.20.0: any Underline variant (.underline is Single)
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

## Outline groups: collapse and expand (since 0.21.0)

A collapsed Excel group is two things at once — the member rows/columns are hidden AND the summary
row/column after the span carries the `collapsed` marker that draws the "+" button. `collapseRows`
/ `collapseCols` compose both — the very fold behind `xl group-rows --collapsed` (`Sheet.groupRows`),
so on an ungrouped sheet `collapseRows(span)` equals `groupRows(span, 1, collapsed = true)` — make
ungrouped members a level-1 group, and keep a member's existing outline level; `expandRows` /
`expandCols` unhide the members and clear the marker, keeping the level, and leave rows/columns that
never had properties untouched ([#465](https://github.com/TJC-LP/xl/issues/465)).

Spans carry their axis: take a `(Row, Row)` / `(Column, Column)` pair or a `RowSpan` / `ColSpan`
(`RowSpan.parse("2:3")`, `ColSpan.parse("E:H")` — each refuses the other axis at parse time). The
`CellRange` overloads return `XLResult[Sheet]` and accept only a full-row range for the row forms
and a full-column range for the column forms: `sheet.collapseRows("E:H".asRange …)` is an
`InvalidReference`, never a million hidden rows.

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.1
import com.tjclp.xl.scripting.{*, given}

// Whole-row/column spans are runtime strings (the ref macro takes A1 / A1:B2 shapes) — parse them.
val cols = orExit(ColSpan.parse("E:H"))
val detail = Sheet("Detail")
  .put(ref"A1", "Region").put(ref"A2", "North").put(ref"A3", "South").put(ref"A4", "Total")
  .put(ref"B2", 10).put(ref"B3", 20).put(ref"B4", fx"=SUM(B2:B3)")
  .collapseRows(Row.from1(2), Row.from1(3)) // rows 2-3 hidden at level 1, row 4 marked collapsed
  .collapseCols(cols)                       // E:H hidden at level 1, column I marked collapsed

val rows = orExit(RowSpan.parse("2:3"))
val reopened = detail.expandRows(rows).expandCols(cols)
val byRange: XLResult[Sheet] = detail.expandRows(orExit("2:3".asRange)) // full-row CellRange: Right
Excel.writeChecked(Workbook(detail), "/tmp/outline.xlsx")
println(s"rows hidden: ${detail.rowProperties.count(_._2.hidden)}, reopened: ${reopened.rowProperties.count(_._2.hidden)}, byRange: ${byRange.isRight}")
```

## Print and view setup

`SheetView` (gridlines, zoom) and `PageSetup` (orientation, fit, margins, header/footer, print
area, repeat rows) live in `com.tjclp.xl.sheets` and are not part of the prelude export — import
them explicitly:

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.1
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

`orExit(result)` (since 0.21.0) is the script-shaped alternative: the value on `Right`, or the
error on stderr and exit status 1, rendered by `XLError.renderDiagnostic` — the one renderer the
CLI's own `Diagnostics` uses, so a failing script prints the same bytes as a failing `xl` call:
`Error: <message>`, then indented `code: <CODE>`, `did you mean: …` (when the error has
candidates) and `hint: …` (when it has one). `exitMessage(err)` is that text, for scripts that
report and continue.

```scala
val wb = orExit(Workbook.named("Data", "Summary")) // DuplicateSheet on a repeat → printed, exit 1
val sales = orExit(wb("Sales"))                    // SheetNotFound → printed with its hint, exit 1
val summary = orExit(Excel.readSheet("model.xlsx", "Sumary")) // 0.21.1: readSheet is an XLResult
// Error: Sheet not found: 'Sumary'. Available: Data, Summary     ← Excel.readSheet's error via orExit
//   code: SHEET_NOT_FOUND
//   did you mean: Summary
//   hint: list sheets with `xl -f <file> sheets`
```

## `Excel` vs `ExcelIO`

| | `Excel` (sync facade) | `ExcelIO` (cats-effect) |
|---|---|---|
| Style | `Excel.read("in.xlsx")` returns `Workbook`, throws at the IO edge | `ExcelIO.instance[IO].read(path)` returns `IO[Workbook]` |
| For | Scripts, REPL, quick tools | Production services, streaming, resource safety |
| Streaming | — | `readStream`/`writeStream`: `fs2.Stream[F, RowData]`, O(1) memory |

Both are in scope from the prelude. Switch to streaming above ~100k rows — `Excel.read` loads
the whole workbook:

```scala
//> using scala 3.9.0
//> using dep com.tjclp::xl:0.21.1
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
