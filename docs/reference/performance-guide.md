# Performance Guide

## Choosing the Right I/O Mode

XL provides two distinct I/O implementations with different performance characteristics. Understanding when to use each is critical for optimal performance.

## Decision Matrix

| Dataset Size | Styling Needs | Read Mode | Write Mode | Memory | Performance |
|--------------|---------------|-----------|------------|--------|-------------|
| **< 10k rows** | Any | In-memory | In-memory | ~10MB | Excellent |
| **10k-50k rows** | Full styling | In-memory | In-memory | ~25-50MB | Good |
| **50k-100k rows** | Full styling | In-memory | In-memory | ~50-100MB | Fair |
| **100k+ rows** | Minimal | **Streaming** (sequential) | **Streaming** | ~10-50MB (SST dependent) | Excellent |
| **100k+ rows** | Full styling / metadata | In-memory | `writeWorkbookStream` if already materialized | O(n) workbook, lower writer allocation | Good |

## I/O Modes Explained

### In-Memory Mode (Default)

**Use For**: Small‑to‑medium workbooks (<100k rows) or any workload that needs full styling and metadata.

**Characteristics**:
- Builds full XML trees in memory.
- Full feature support (SST, styles, merged cells, RichText).
- Deterministic output (byte‑identical structure on re‑write).
- Compact XML with DEFLATED compression by default; pretty‑printed STORED output is available for debugging via `WriterConfig.debug`.

**API**:
```scala
import com.tjclp.xl.io.ExcelIO
import cats.effect.IO

val excel = ExcelIO.instance[IO]

// Write
excel.write(workbook, path).unsafeRunSync()

// Read
excel.read(path).map { wb =>
  // Process workbook
}.unsafeRunSync()
```

**Memory Profile (approximate)**:
```
10k rows:   ~10MB
50k rows:   ~50MB
100k rows:  ~100MB
500k rows:  ~500MB (not recommended)
```

**Performance**:
- Write: ~100k rows/sec
- Read: ~55k rows/sec

---

### Row-Stream Write Mode

**Use For**: Large data generation (100k+ rows) — values, shared strings and a declared style table.

**Characteristics**:
- O(1) constant memory for worksheet data (~10MB regardless of row count), plus O(distinct strings) for the shared strings table.
- fs2‑data‑xml event streaming; no intermediate XML trees.
- Styles: `writeStream` / `writeStreamWithAutoDetect` write the minimal (default-only) table; `writeStreamStyled` / `writeStreamStyledWithAutoDetect` take a `Vector[CellStyle]` that `StyledRowData.cellStyles` indexes (GH-223, GH-675).
- **Limitations**: no API for workbook metadata such as merged cells.
- For an already-materialized workbook that needs metadata preservation with lower writer allocation, use `ExcelIO.writeWorkbookStream`.

**API**:
```scala
import com.tjclp.xl.io.Excel
import com.tjclp.xl.io.RowData
import fs2.Stream

val excel = com.tjclp.xl.io.Excel.forIO

// Generate 1M rows
Stream.range(1, 1_000_001)
  .map(i => RowData(i, Map(
    0 -> CellValue.Text(s"Row $i"),
    1 -> CellValue.Number(BigDecimal(i))
  )))
  .through(excel.writeStream(path, "Data"))
  .compile.drain
  .unsafeRunSync()

// Memory: ~10MB constant!
```

**Memory Profile**:
```
100k rows:  ~10MB (constant)
1M rows:    ~10MB (constant)
10M rows:   ~10MB (constant)
```

**Performance**:
- Write: ~88k rows/sec in benchmarks.
- In practice, ~4–5x faster and vastly more memory‑efficient than typical in‑memory POI usage for large datasets.

**Tradeoffs**:
- ✅ Constant memory for worksheet data.
- ✅ Excellent throughput at high row counts.
- ✅ Shared strings by default (`SstPolicy.Never` keeps inline strings).
- ✅ Per-cell styles from a table declared up front (`writeStreamStyled`, `writeStreamStyledWithAutoDetect`).

**Where the scratch file goes**: all row writers (`writeStream`, `writeStreamStyled`,
`writeStreamsSeq` and their auto-detect variants) consume each input once into a temporary worksheet
body, then assemble the ZIP with the metadata its dynamic-array anchors need (#749). Auto-detect
also tracks bounds to emit `<dimension>`; explicit hints are kept verbatim. Temporary disk use is
O(worksheet XML), one file per sheet, deleted when the write ends, success or failure. It lands in `java.io.tmpdir` by default. Redirect it when that directory is small,
read-only, or slower than the output volume:

```scala
val excel = ExcelIO.instance[IO].withSpillDir(fastVolume)  // must already exist
```

The `xl` CLI reads the same setting from the `XL_SPILL_DIR` environment variable — once, at its
one write interpreter — for its two-pass CSV import and for the resvg rasterizer's scratch SVG
(`docs/reference/cli.md`, "Where a streamed write spills"); the library itself never reads the
environment.

Constant-memory holds either way — the spill is disk, not heap — but a 1M-row write moves real
bytes through it, so on a container with a tiny `/tmp` this is the difference between working and
`No space left on device`. The destination ZIP is written after the input has been consumed.
The metadata accumulator itself holds at most two dynamic-array variants.

Two things you inherit by redirecting it. The scratch file holds the worksheet body in cleartext
until the write ends; it is created `rw-------` on POSIX either way, but the *directory's*
permissions become your call, and the natural targets for this setting (a shared scratch mount, a
big fast volume) are often group- or world-traversable — prefer a directory the process owns. And
the spill is deleted by the write's own bracket, not registered with `deleteOnExit` (which leaks in
long-lived JVMs), so a `SIGKILL` or OOM-kill mid-write strands it: in `java.io.tmpdir` the OS
eventually sweeps it, on a persistent volume it stays until you do.

---

### Streaming Read Mode

**Use For**: Sequential processing of large worksheets (filtering, aggregation, ETL) where you don’t need full styling information.

**Characteristics**:
- O(1) memory for worksheet data by streaming XML via `fs2.io.readInputStream` + fs2‑data‑xml.
- `SharedStrings` (if present) are parsed once up front, so memory use scales with the number of unique strings, not total rows.
- Emits `Stream[F, RowData]` with 1‑based row indices and 0‑based column indices.

**API**:
```scala
import com.tjclp.xl.io.Excel
import cats.effect.IO
import cats.effect.unsafe.implicits.global

val excel = Excel.forIO

excel.readStream(path)
  .filter(_.rowIndex > 1) // skip header
  .evalMap(row => IO.println(row))
  .compile.drain
  .unsafeRunSync()
```

**Tradeoffs**:
- ✅ True streaming (no materialization of full sheets).
- ✅ Great for ETL and analytics pipelines.
- ❌ Does not expose full formatting/metadata; for that, use the in‑memory APIs.

## Performance Optimization Techniques

### 1. Batch Operations (30-50% faster)

**Bad** (N intermediate sheets):
```scala
(1 to 10000).foldLeft(sheet) { (s, i) =>
  s.put(ref"A1".down(i - 1), i)
}
// Creates 10,000 intermediate Sheet instances
```

**Good** (fold into a Patch, apply once):
```scala
val patch = (1 to 10000).foldLeft(Patch.empty) { (p, i) =>
  p ++ (ref"A1".down(i - 1) := i)
}
sheet.put(patch)
// Creates 1 Sheet instance
```

**Best** (batch varargs put with codecs):
```scala
sheet.put(
  (1 to 10000).map(i => ARef.from0(0, i - 1) -> i)*
)
// Type-safe, auto-format inference, single allocation
```

---

### 2. Use Patch Composition (Deferred Execution)

**Bad** (eager evaluation):
```scala
var s = sheet
for i <- 1 to 1000 do
  val r = ref"A1".down(i - 1)
  s = s.put(r, s"Row $i")                  // 1000 Sheet instances
  s = s.merge(CellRange(r, r.right(1)))    // 1000 more Sheet instances
```

**Good** (deferred with Patch):
```scala
val patch = (1 to 1000).foldLeft(Patch.empty) { (p, i) =>
  val r = ref"A1".down(i - 1)
  p ++ (r := s"Row $i") ++ CellRange(r, r.right(1)).merge
}
sheet.put(patch)  // Execute once at end
```

**Impact**: Reduces intermediate allocations by 50-70%

---

### 3. Prefer Batch `Sheet.put` for Type Safety + Auto-Formatting

**Manual** (verbose, error‑prone):
```scala
import com.tjclp.xl.{*, given}
import java.time.{LocalDate, LocalDateTime}

val dateStyle = CellStyle.default.withNumFmt(NumFmt.Date)
val decimalStyle = CellStyle.default.withNumFmt(NumFmt.Decimal)

sheet
  .put(ref"A1", CellValue.Text("Revenue"))
  .put(ref"B1", CellValue.DateTime(LocalDate.of(2025, 11, 10).atStartOfDay))
  .withCellStyle(ref"B1", dateStyle)    // Must manually set format
  .put(ref"C1", CellValue.Number(BigDecimal("1000000.50")))
  .withCellStyle(ref"C1", decimalStyle) // Must manually set format
```

**Batch `put` with codecs** (clean, automatic):
```scala
import com.tjclp.xl.{*, given}
import java.time.LocalDate

sheet.put(
  ref"A1" -> "Revenue",                         // String
  ref"B1" -> LocalDate.of(2025, 11, 10),        // Auto: date format
  ref"C1" -> BigDecimal("1000000.50")           // Auto: decimal format
)
```

**Benefits**:
- Auto‑inferred number formats via `CellCodec`.
- Type‑safe (compile error if unsupported type).
- Single bulk update instead of many tiny ones.
- Cleaner code.

---

### 4. Stream from Source (Don't Materialize First)

**Bad** (materializes in memory first):
```scala
val allRows = database.fetchAll()  // Load 1M rows into memory
val cells = allRows.map(row => /* convert to cells */)
sheet.put(cells)  // Then write
```

**Good** (stream directly):
```scala
database.stream()  // fs2.Stream[IO, Row]
  .map(row => RowData(/* convert */))
  .through(excel.writeStream(path, "Data"))
  .compile.drain
  .unsafeRunSync()
```

**Impact**: Constant memory regardless of source size

---

### 5. Use Iterator for Lazy Range Operations

**Current** (after lazy optimizations):
```scala
// Already optimized! range.cells returns Iterator
sheet.fillBy(ref"A1:Z100") { (col, row) =>
  CellValue.Text(s"${col.toLetter}${row.index1}")
}
// No intermediate Vector allocation
```

**Before optimization** (old code):
```scala
val newCells = range.cells.map(...).toVector  // ❌ Unnecessary
sheet.put(newCells)
```

**Impact**: 20-30% memory reduction for large ranges (already fixed in latest version)

---

## Performance Benchmarks

### Actual Measurements (100k rows)

**XL In-Memory**:
- Write: ~1.1s @ ~100MB memory
- Read: ~1.8s @ ~100MB memory

**XL Streaming**:
- Write: ~1.1s @ ~10MB memory (constant)
- Read: ~1.8s @ ~10MB memory (constant, P6.6 complete)

**Apache POI SXSSF** (streaming):
- Write: ~5s @ ~800MB memory
- Read: ~8s @ ~1GB memory

**Improvement**:
- XL write: **4.5x faster, 80x less memory**
- XL read: **4.4x faster, 80x less memory**

### Sheet cell storage: `CellStore` against the `Map` it replaced (#891)

`Sheet.cells` is a `CellStore` (a row-radix trie of struct-of-arrays rows) where it used to be a
`Map[ARef, Cell]`. `SheetCellsBenchmark` measures both through the public `Sheet` API, so the same
source runs on either side of the swap (the one line that differs is full iteration:
`valuesIterator` on the map, `iterator` on the store).

```bash
./mill xl-benchmarks.runJmh -f 1 -wi 3 -i 5 -p shape=dense20,wide500 -p size=100000 \
  'SheetCellsBenchmark\.(getHit|getMiss|iterateAll|usedRange|singlePut|putChain10k|randomPuts5k|buildSorted|buildUnsorted)$'
```

100,000 cells per sheet, in rows of 20 columns (`dense20`) or 500 (`wide500`); average time per
operation, with JMH's 99.9% error. Measured 2026-11-19 on an Apple M4 Max (JDK 25, G1) **while
other builds were using the machine**: four runs per side, interleaved, and the load average moved
between 7 and 86 across them. The table gives each side's best run and, in brackets, its worst, so
the spread is visible; the ratio compares the best runs, since interference only adds time. Treat
a ratio within about 15% of 1 as no difference, and rerun on a quiet machine before quoting any
of this.

| Operation | Shape | Unit | `Map`: best run (worst) | `CellStore`: best run (worst) | Store / Map |
|---|---|---|---|---|---|
| `getHit` | dense20 | ns | 23.5 ±0.336 (to 45.8) | 12.3 ±0.399 (to 24.1) | 0.52× |
| `getHit` | wide500 | ns | 24.0 ±0.298 (to 59.2) | 16.0 ±0.459 (to 38.4) | 0.66× |
| `getMiss` | dense20 | ns | 11.5 ±0.628 (to 29.1) | 7.26 ±0.075 (to 12.6) | 0.63× |
| `getMiss` | wide500 | ns | 10.6 ±0.592 (to 27.0) | 4.10 ±0.031 (to 6.99) | 0.39× |
| `iterateAll` | dense20 | µs | 2,992 ±2,316 (to 10,734) | 744 ±92.2 (to 2,653) | 0.25× |
| `iterateAll` | wide500 | µs | 2,782 ±317 (to 10,617) | 532 ±12.2 (to 2,086) | 0.19× |
| `usedRange` | dense20 | µs | 3,105 ±790 (to 25,515) | 0.027 ±0.001 (to 0.388) | 1/115,305 |
| `usedRange` | wide500 | µs | 2,795 ±1,238 (to 29,021) | 0.022 ±0.000 (to 0.277) | 1/126,408 |
| `singlePut` | dense20 | ns | 80.5 ±1.40 (to 559) | 118 ±1.57 (to 243) | **1.47×** |
| `singlePut` | wide500 | ns | 75.4 ±1.62 (to 458) | 232 ±1.22 (to 1,294) | **3.07×** |
| `putChain10k` | dense20 | µs | 1,482 ±175 (to 6,320) | 1,728 ±11.5 (to 3,438) | **1.17×** |
| `putChain10k` | wide500 | µs | 1,412 ±624 (to 6,594) | 1,626 ±18.0 (to 3,090) | **1.15×** |
| `randomPuts5k` | dense20 | µs | 778 ±136 (to 2,914) | 867 ±43.5 (to 1,884) | **1.11×** |
| `randomPuts5k` | wide500 | µs | 757 ±55.1 (to 3,259) | 1,551 ±105 (to 4,337) | **2.05×** |
| `buildSorted` | dense20 | ms | 20.3 ±1.69 (to 42.9) | 20.5 ±2.01 (to 51.9) | 1.01× |
| `buildSorted` | wide500 | ms | 18.4 ±2.98 (to 40.3) | 20.0 ±0.372 (to 61.9) | **1.08×** |
| `buildUnsorted` | dense20 | ms | 20.2 ±2.75 (to 47.4) | 35.4 ±1.82 (to 62.3) | **1.75×** |
| `buildUnsorted` | wide500 | ms | 22.5 ±2.96 (to 44.8) | 39.1 ±0.361 (to 79.6) | **1.74×** |

Reads are faster everywhere: a lookup is index arithmetic on the row and column bits instead of a
hash and an equality test, iteration walks arrays, and `usedRange` reads the trie's summaries
instead of scanning every cell.

**Where the store is slower** (bold): single-cell writes. A `put` copies the row's leaf (or, in a
row of more than 64 cells, the 256-column band it lands in) and the path above it, where the map
copied one small trie node per level. On 20-column rows that is 1.1× to 1.5× the map's cost; on
500-column rows a single `put` is about 3× and 5,000 random puts 2×. `buildSorted` and
`buildUnsorted` are folds of 100,000 single `put`s, not builder runs: in row-major order they match
the map, in shuffled order they cost about 1.75× on both shapes. Code that writes many cells should
not fold single `put`s: `Sheet.withValues`, `CellStore.putAll` and `CellStore.from` /
`CellStore.newBuilder` build each row once.

The builder is the bulk path (the worksheet reader uses it), and it is not what the JMH rows above
measure. 2,000,000 cells in row-major order, same machine and caveat (`ThreadMXBean` allocation;
the smallest SerialGC heap that completes, less the 236 MiB the input cells need; best build time
of five, over three JVMs):

| Building 2M cells | Allocated per cell | Heap beyond the input | Build time |
|---|---|---|---|
| `Map.newBuilder`, 8 or 20 columns | 213 B | 136 MiB | 563–1,101 ms |
| `CellStore.newBuilder`, 8 columns | 13.8 B | 32 MiB | 17–18 ms |
| `CellStore.newBuilder`, 20 columns | 19.0 B | 40 MiB | 19–30 ms |
| `CellStore.newBuilder`, 8 columns, shuffled input | 63.7 B | 72 MiB (of 244) | 323–375 ms |

Loading a 1,000,000 × 8 numeric workbook with `XlsxReader.read` (G1, the smallest `-Xmx` in 64 MiB
steps) needs 5,120 MiB on the store, where the map needed 5,696 MiB (decimals) and 5,632 MiB
(integers). Most of what remains is the parsed worksheet, which stays live while the store is
built from it; see the memory notes on `--max-size` in the [CLI reference](cli.md).

### Recalculation benchmark (#767)

`xl-evaluator.perf` (sources in `xl-evaluator/test-perf`) builds a synthetic freight-loads workbook
(`LoadsBook`: a Loads sheet of keyed rows with whole-column defined names, Summary and Rate History
sheets of COUNTIFS/SUMIFS/MINIFS/MAXIFS over them) and recalculates it with work counters and
per-phase timing on. Shapes: `Tiny` (1,600 formulas, the CI gate), `H10` (10k rows, helper
criteria, 38k formulas), `R20` (20k rows, raw criteria, 8k formulas), `Real` (33,233 rows, about
188k formulas) and `RunTotal` (running totals over a formula column).

```bash
./mill xl-evaluator.perf.runMain com.tjclp.xl.formula.perf.LoadsBookMain shape=Tiny mode=recalc
XL_PERF_HEAP=4g ./mill xl-evaluator.perf.runMain com.tjclp.xl.formula.perf.LoadsBookMain shape=Real mode=checked
./mill xl-evaluator.perf.testForked      # the gates CI's non-blocking `perf` job runs
```

`mode` is `recalc`, `parallel`, `uncached` (what `Excel.writeChecked` computes) or `checked`
(`recalculateUncached` plus `Excel.write`, timing the write). The module forks with `-Xmx2g`
unless `XL_PERF_HEAP` says otherwise.

Baseline on 0.24.1 (`mode=recalc`, sequential, `-XX:ActiveProcessorCount=2`, M4 Max under shared
load). Phases are inclusive: Parse runs inside graph build and eval, criteria inside eval.

| Phase (s) | Tiny | H10 | R20 |
|---|---|---|---|
| Graph build | 0.33 | 48.4 | 0.75 |
| Parse (inside build and eval) | 0.14 | 0.73 | 0.40 |
| Tarjan + blocked dependents | 0.02 | 5.35 | 0.01 |
| Kahn | 0.07 | 23.7 | 0.01 |
| Eval | 0.30 | 56.6 | 116.7 |
| Criteria (inside eval) | 0.16 | 55.5 | 116.3 |
| `Sheet.put` threading | 0.01 | 0.12 | 0.09 |
| Wall | 0.8 | 134 | 118 |
| Graph edges | 186,400 | 92,520,000 | 0 |
| Matcher calls | 180,118 | 81,101,873 | 161,824,716 |

H10 needs about 25 GB of heap for its 92.5M per-cell edges; `Real` with helper criteria does not
fit in 64 GB. Writing Tiny (`mode=checked`) takes 0.25 s.

Per-pass caches (#767 PR 1): a pass caches each sheet's used range, each defined name's parsed body
and static range, and column slices of tall single-column criteria ranges (`PassCache`); a slice is
built only once the pass has evaluated every formula in its box; a column whose formulas are all done
gets one slice for every window of it, and any other window is sliced on its second request. R20
(`-Xmx4g`, same machine and load, run back to back): `mode=recalc` 151.5 s before, 12.2 s after;
`mode=checked` 160.2 s before, 13.1 s after; `mode=parallel` 6.4 s after; with 9 slices built, no
scan fallback, one used-range scan and 14 name parses (116,875 before). Matcher calls are
unchanged (161,824,716): every row is still tested; the criteria index removes them. H10 stays
bound by its legacy graph until the compact graph lands.

Whole-axis references (#747): lookups and returned references use the same `PassCache.usedRange`
as aggregates, including INDEX, INDIRECT, OFFSET and logical range reads. Bounds are scanned once
per sheet and structural generation during a recalculation, shared by sequential and parallel
readers. A structural write invalidates the bounds; a formula-cache update does not. Empty-sheet
bounds are cached too. Bounded references never request bounds, and evaluation outside an active
pass reads each sheet snapshot afresh. This removes O(formulas × populated cells) bounds work;
lookup key searching and result materialization keep their existing costs and semantics.

To run the opt-in 10,000-lookup / 200,000-cell acceptance harness (an early match isolates bounds
scanning from linear key-search cost):

```bash
XL_WHOLE_AXIS_PERF=1 ./mill -i xl-evaluator.test.testOnly com.tjclp.xl.formula.eval.WholeAxisUsedRangeSpec
```

The harness prints elapsed recalculation time, checks all 10,000 results, and requires exactly one
used-range scan; wall time is diagnostic, not a flaky pass/fail threshold. A local run on
2026-10-28 recalculated all 10,000 formulas in 0.712 s with one scan and all results correct;
workbook construction and file I/O are outside that measurement.

Criteria index (#767 PR 2): each slice gets an equality index on first use (`EqIndex`, posting lists
of the rows holding each folded text, stripped number, boolean or error), and a criteria call whose
ranges all read through slices tests only the rows every equality criterion selects, still with
`CriteriaMatcher.matches` for every criterion. Comparisons, wildcards, `<>` and `""` are not
indexed; a call with no indexed criterion, or whose smallest selection is more than half its rows,
tests every row. Candidate lists are memoised for the pass, so the Rate History columns that share
lane, carrier and year criteria share one list. `LoadsBookMain index=off` turns the index off.
Same machine, run back to back:

| Run | Index off | Index on |
|---|---|---|
| R20 `mode=recalc` | 10.9 s, 161,824,716 matcher calls | 2.3 s, 613,617 |
| R20 rows=5000 `mode=recalc` | 3.9 s, 41,418,819 | 2.3 s, 500,790 |
| R20 `mode=checked` (write 1.1 s) | | 3.4 s |
| R20 `mode=parallel` | | 2.4 s |
| Real `crit=raw mode=checked`, `-Xmx4g` (88k formulas, write 2.3 s) | | 9.4 s, 5,449,218 |
| Tiny matcher calls | 180,118 | 1,478 |

Every R20 and Real call is indexed (none window-only). With the index, R20 spends 0.9 s building the
legacy graph and 1.0 s in criteria calls.

With the compact graph (#767: one node per formula, per distinct range and per formula-column
segment, so a range read costs O(log rows) edges), H10's graph has 116,024 edges and 29,997 segment
nodes: build 0.97 s, Tarjan 0.01 s, Kahn 0.02 s, wall 49 s (eval, all criteria scans, is the rest),
heap peak 1.4 GB under `-Xmx2g`; the 40 oracle cells are unchanged. A 200k-row running total
(`RunTotal`) builds, cycle-checks and orders its 3.5M-edge graph in 1.6 s. `CompactGraphGatesSpec`
gates both graphs: `GraphEdges + VirtualNodes <= 40 x (formulas + ranges)` on H10, and Tarjan plus
Kahn walking at most `3 x (formulas + ranges) x log2(rows)` edges on `RunTotal(200k)`.

With the compact graph, the per-pass caches and the criteria index together, `LoadsBookMain` at
`-Xmx4g` (direct JVM, `ActiveProcessorCount=2`, shared machine) measures:

| Run | Wall | Phases | Counters |
|---|---|---|---|
| Real `mode=checked` (helper criteria, 187,699 formulas) | 7.6 s | graph build 2.2 s (parse 1.7 s, nested in build and eval), eval 2.9 s (criteria 1.3 s), write 1.9 s | 665,906 edges, 99,696 segment nodes, 88,000 calls indexed, 808,484 matcher calls, 0 scan fallbacks |
| Real `crit=raw mode=checked` (88,000 formulas) | 7.0 s | graph build 1.9 s, eval 3.0 s, write 1.6 s | 5,449,218 matcher calls |
| H10 `mode=recalc` | 2.4 s | graph build 0.9 s, eval 1.3 s | 75,766 matcher calls (81.1M with scans) |
| R20 `mode=recalc` | 1.9 s | | 613,617 matcher calls |

The Real helper book also completes under `-Xmx1g` (8.7 s). Every earlier engine runs out of memory
on it at `-Xmx4g`, and the compact graph without the index did not finish in 15 minutes. A full dump
of its 187,699 formula values is identical with the index on and off, and between sequential and
parallel recalculation. In a JFR profile of the Real `checked` run on 0.25.0, formula parsing is the
largest CPU cost (about 22% of samples): each formula was parsed once while the graph was built and
again when it was evaluated. Next come the XLSX write (about 24%, partly on a separate thread) and
the criteria calls (about 15%, of which about 5% is building the column slices).

Since #790 a pass reuses the graph build's parse of each formula cell (looked up by the cell and
checked against its current text) for cell evaluation, the external-reference pin check of a cached
formula containing `[`, iterative fixpoints over cycle members, the dynamic-cell scan and the pin
check of a pass with failures. The `FormulaParses` counter of a Real pass is 187,699, one per
formula. A pass parses again only a formula the build could not parse or whose text has changed
since; under the legacy engine, which keeps no parse, evaluation parses every formula again. The
data-table seeder's what-if fixpoints parse their cycle members themselves.

---

## Common Performance Pitfalls

### Pitfall 1: Fold with Put Instead of Batch Put
```scala
// ❌ Bad: O(n) sheet copies
val result = data.foldLeft(sheet) { (s, item) =>
  s.put(ARef.from0(0, item.id), item.name)
}

// ✅ Good: Single sheet copy
sheet.put(data.map(item => ARef.from0(0, item.id) -> item.name)*)
```

**Impact**: 10-20x faster for large datasets

### Pitfall 2: Calling usedRange Repeatedly
```scala
// ❌ Bad: Recomputes 4 times (was even worse before optimization)
for _ <- 1 to 4 do
  println(sheet.usedRange)  // Traverses all cells each time

// ✅ Good: Compute once
val range = sheet.usedRange
for _ <- 1 to 4 do
  println(range)
```

**Note**: `usedRange` is now single-pass (75% faster after lazy optimizations), but still recomputes on every call.

### Pitfall 3: Using the Wrong Mode for Feature Needs
```scala
// ✅ Good: Streaming read for sequential ETL (constant memory)
excel.readStream(largeFile)

// ✅ Good: In-memory read when you need styles/metadata
excel.read(largeFile)
```

**Status**: Streaming read is fixed (P6.6); choose based on feature needs, not memory concerns.

---

## When to Optimize

### Don't Optimize (Good Enough)
- **< 1000 rows**: Any approach works fine
- **Simple scripts**: Clarity > performance
- **Prototyping**: Optimize later

### Consider Optimizing
- **10k-100k rows**: Use batching (varargs `Sheet.put` / Patch folds)
- **Complex styling**: Use Patch composition
- **Repeated operations**: Cache usedRange, style registries

### Must Optimize
- **100k+ rows**: Use streaming write
- **Limited memory**: Constant-memory critical
- **Production ETL**: Performance matters

---

## Memory Budget Calculator

**Rule of Thumb**: 1 cell ≈ 100 bytes in memory (JVM overhead)

```
10k cells:    ~1 MB
100k cells:   ~10 MB
1M cells:     ~100 MB
10M cells:    ~1 GB (use streaming!)
```

**Factors that increase memory**:
- Rich styles: +50 bytes per styled cell
- Merged cells: +200 bytes per merge
- Rich text: +100 bytes per additional run
- Formulas: +150 bytes per formula cell

---

## Performance Roadmap

### Current (as of 0.12.2)

> ⚠️ Throughput figures here are indicative, captured on an earlier release; see [STATUS.md](../STATUS.md) for the benchmark tables and caveats.
- ✅ Streaming write: O(1) memory, ~88k rows/sec
- ✅ Streaming read: O(1) memory via `fs2.io.readInputStream` (+ SST materialized once)
- ✅ Lazy optimizations: 30-75% speedup (memoized canonicalKey, single-pass usedRange)

### P6.6 (complete)
- ✅ Fix streaming read to O(1) memory
- ✅ Add memory regression tests

### P6.7 (1 day)
- ✅ DEFLATED compression by default (5-10x smaller files)
- ✅ Configurable compression mode

### SAX/StAX OOXML Writer
- ✅ Lower-allocation `writeWorkbookStream` path for already-materialized workbooks
- ✅ Preserves full OOXML metadata handled by `XlsxWriter`
- ✅ Pure row-stream writers remain available for true O(1) row input

### Future (Post-1.0)
- ⬜ Parallel XML serialization (4-8 threads)
- ⬜ JMH benchmark suite vs POI
- ⬜ Automatic query optimization (builder pattern)

---

## Quick Reference

### Small Workbooks (<10k rows)
```scala
// Use simple in-memory API
val sheet = Sheet("Data")
  .put(ref"A1", "Title")
  .put(ref"B1", 100)

ExcelIO.instance[IO].write(Workbook(Vector(sheet)), path).unsafeRunSync()
```

### Medium Workbooks (10k-100k rows)
```scala
// Use batching
val cells = (1 to 100_000).map(i =>
  ARef.from0(0, i - 1) -> s"Row $i"
)
val sheet = Sheet("Data").put(cells*)

ExcelIO.instance[IO].write(Workbook(Vector(sheet)), path).unsafeRunSync()
```

### Large Workbooks (100k+ rows)
```scala
// Use streaming write
import com.tjclp.xl.io.Excel
import fs2.Stream

Stream.range(1, 1_000_001)
  .map(i => RowData(i, Map(0 -> CellValue.Text(s"Row $i"))))
  .through(Excel.forIO.writeStream(path, "Data"))
  .compile.drain
  .unsafeRunSync()

// Styled: StyledRowData.cellStyles index the table; the auto-detect variant adds <dimension>
val dateStyle = CellStyle.default.withNumFmt(NumFmt.Date)
Stream.range(1, 1_000_001)
  .map(i => StyledRowData(i, Map(0 -> CellValue.DateTime(start.plusDays(i).atStartOfDay())), Map(0 -> 0)))
  .through(ExcelIO.instance[IO].writeStreamStyledWithAutoDetect(path, "Data", Vector(dateStyle)))
  .compile.drain
  .unsafeRunSync()
```

### Reading Any Size (Streaming)
```scala
// Use streaming read for constant memory (O(1))
Excel.forIO.readStream(path)
  .evalMap(row => IO.println(row))
  .compile.drain
  .unsafeRunSync()
```

---

## Related Documentation

- [STATUS.md](../STATUS.md) - Current performance numbers
- [io-modes.md](../design/io-modes.md) - Architecture decisions
- [examples.md](examples.md) - Code samples
