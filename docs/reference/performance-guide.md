# Performance Guide

## Choosing the Right I/O Mode

XL provides two distinct I/O implementations with different performance characteristics. Understanding when to use each is critical for optimal performance.

## Decision Matrix

| Dataset Size | Styling Needs | Read Mode | Write Mode | Memory | Performance |
|--------------|---------------|-----------|------------|--------|-------------|
| **< 10k rows** | Any | In-memory | In-memory | ~10MB | Excellent |
| **10k-50k rows** | Full styling | In-memory | In-memory | ~25-50MB | Good |
| **50k-100k rows** | Full styling | In-memory | In-memory | ~50-100MB | Fair |
| **100k+ rows** | Minimal | In-memory | **Streaming** | ~10MB (write only) | Excellent for write |
| **100k+ rows** | Full styling | Not supported | Not supported | N/A | See roadmap P7.5 |

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

### Streaming Write Mode

**Use For**: Large data generation (100k+ rows) when you can live with minimal styling and inline strings.

**Characteristics**:
- O(1) constant memory for worksheet data (~10MB regardless of row count).
- fs2‑data‑xml event streaming; no intermediate XML trees.
- **Limitations**: No SST support (inline strings only), minimal styles (no rich formatting or merged cells).

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
  .through(excel.writeStreamTrue(path, "Data"))
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
- ❌ No SST (larger files if many duplicate strings).
- ❌ Minimal styles only (no rich formatting at scale).

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
  s.put(cell"A$i", CellValue.Number(i))
}
// Creates 10,000 intermediate Sheet instances
```

**Good** (single putAll):
```scala
val cells = (1 to 10000).map(i => Cell(cell"A$i", CellValue.Number(i)))
sheet.put(cells)
// Creates 1 Sheet instance
```

**Best** (putMixed with type safety):
```scala
sheet.put(
  (1 to 10000).map(i => cell"A$i" -> i)*
)
// Type-safe, auto-format inference, single allocation
```

---

### 2. Use Patch Composition (Deferred Execution)

**Bad** (eager evaluation):
```scala
var s = sheet
for i <- 1 to 1000 do
  s = s.put(cell"A$i", s"Row $i")  // 1000 Sheet instances
  s = s.merge(range"A$i:B$i")      // 1000 more Sheet instances
```

**Good** (deferred with Patch):
```scala
import com.tjclp.xl.dsl.*

val patch = (1 to 1000).foldLeft(Patch.empty: Patch) { (p, i) =>
  p ++ (cell"A$i" := s"Row $i") ++ range"A$i:B$i".merge
}
sheet.put(patch)  // Execute once at end
```

**Impact**: Reduces intermediate allocations by 50-70%

---

### 3. Prefer Batch `Sheet.put` for Type Safety + Auto-Formatting

**Manual** (verbose, error‑prone):
```scala
import com.tjclp.xl.*

val dateStyle = CellStyle.default.numFmt(NumFmt.Date)
val decimalStyle = CellStyle.default.numFmt(NumFmt.Number)

sheet
  .put(ref"A1", CellValue.Text("Revenue"))
  .put(ref"B1", CellValue.DateTime(LocalDate.of(2025, 11, 10)))
  .withCellStyle(ref"B1", dateStyle)    // Must manually set format
  .put(ref"C1", CellValue.Number(BigDecimal("1000000.50")))
  .withCellStyle(ref"C1", decimalStyle) // Must manually set format
```

**Batch `put` with codecs** (clean, automatic):
```scala
import com.tjclp.xl.*
import com.tjclp.xl.codec.syntax.*

sheet.put(
  ref"A1" -> "Revenue",                         // String
  ref"B1" -> LocalDate.of(2025, 11, 10),        // Auto: date format
  ref"C1" -> BigDecimal("1000000.50")           // Auto: number format
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
  .through(excel.writeStreamTrue(path, "Data"))
  .compile.drain
  .unsafeRunSync()
```

**Impact**: Constant memory regardless of source size

---

### 5. Use Iterator for Lazy Range Operations

**Current** (after lazy optimizations):
```scala
// Already optimized! range.cells returns Iterator
sheet.fillBy(range"A1:Z100") { (col, row) =>
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

**XL Streaming Write**:
- Write: ~1.1s @ ~10MB memory (constant)
- Read: ~1.8s @ ~100MB memory (NOT constant - bug)

**Apache POI SXSSF** (streaming):
- Write: ~5s @ ~800MB memory
- Read: ~8s @ ~1GB memory

**Improvement**:
- XL write: **4.5x faster, 80x less memory**
- XL read: **4.4x faster, but not constant-memory yet**

---

## Common Performance Pitfalls

### Pitfall 1: Fold with Put Instead of PutAll
```scala
// ❌ Bad: O(n) sheet copies
val result = data.foldLeft(sheet) { (s, item) =>
  s.put(cell"A${item.id}", item.name)
}

// ✅ Good: Single sheet copy
val cells = data.map(item => Cell(cell"A${item.id}", CellValue.Text(item.name)))
sheet.put(cells)
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

### Pitfall 3: Using Streaming Read for Large Files
```scala
// ❌ Bad: Spikes memory (known bug)
excel.readStream(largeFile)  // Uses readAllBytes() internally

// ✅ Good: Use in-memory read until P6.6 fix
excel.read[IO](largeFile)  // O(n) memory but predictable
```

**Status**: P6.6 will fix this (2-3 days)

---

## When to Optimize

### Don't Optimize (Good Enough)
- **< 1000 rows**: Any approach works fine
- **Simple scripts**: Clarity > performance
- **Prototyping**: Optimize later

### Consider Optimizing
- **10k-100k rows**: Use batching (putAll, putMixed)
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

### Current (As of 2025-11-11)
- ✅ Streaming write: O(1) memory, 88k rows/sec
- ⚠️ Streaming read: O(n) memory (broken, see P6.6)
- ✅ Lazy optimizations: 30-75% speedup (memoized canonicalKey, single-pass usedRange)

### P6.6 (2-3 days)
- ✅ Fix streaming read to O(1) memory
- ✅ Add memory regression tests

### P6.7 (1 day)
- ✅ DEFLATED compression by default (5-10x smaller files)
- ✅ Configurable compression mode

### P7.5 (3-4 weeks)
- ✅ Two-phase streaming writer with full SST/styles
- ✅ O(1) memory for writes with rich formatting
- ✅ Disk-backed SST for very large datasets

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
  .put(cell"A1", "Title")
  .put(cell"B1", 100)

ExcelIO.instance.write[IO](Workbook(Vector(sheet)), path).unsafeRunSync()
```

### Medium Workbooks (10k-100k rows)
```scala
// Use batching
import com.tjclp.xl.codec.syntax.*

val cells = (1 to 100_000).map(i =>
  cell"A$i" -> s"Row $i"
)
val sheet = Sheet("Data").put(cells*)

ExcelIO.instance.write[IO](Workbook(Vector(sheet)), path).unsafeRunSync()
```

### Large Workbooks (100k+ rows, minimal styling)
```scala
// Use streaming write
import com.tjclp.xl.io.Excel
import fs2.Stream

Stream.range(1, 1_000_001)
  .map(i => RowData(i, Map(0 -> CellValue.Text(s"Row $i"))))
  .through(Excel.forIO.writeStreamTrue(path, "Data"))
  .compile.drain
  .unsafeRunSync()
```

### Reading Any Size (Until P6.6)
```scala
// Use in-memory read (streaming read has bug)
ExcelIO.instance.read[IO](path).map {
  case Right(wb) => // Process
  case Left(err) => // Handle
}.unsafeRunSync()
```

---

## Related Documentation

- [STATUS.md](../STATUS.md) - Current performance numbers
- [streaming-improvements.md](../plan/streaming-improvements.md) - Roadmap for fixes
- [io-modes.md](../design/io-modes.md) - Architecture decisions
- [examples.md](examples.md) - Code samples
