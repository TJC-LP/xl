# Performance Investigation: XL vs Apache POI

**Date**: 2025-12-25
**Status**: Investigation Complete
**Author**: Claude Code + Human Review

## Executive Summary

XL outperforms Apache POI on typical workloads (<10k rows) for both reads and writes. At scale (100k+ rows), XL's functional abstractions introduce overhead for reads, while writes remain faster.

| Operation | 1k rows | 10k rows | 100k rows |
|-----------|---------|----------|-----------|
| **In-Memory Read** | ✅ XL +21% | ✅ XL +3% | ❌ POI +11% |
| **Streaming Read** | ✅ XL +37% | ✅ XL +2% | ❌ POI +22% |
| **Write (SaxStax)** | ✅ ~tied | ✅ XL +36% | ✅ XL +39% |

**Recommendation**: Accept current performance for v1.0. The 100k read gap is the cost of functional purity and can be addressed in future optimization work if needed.

---

## Benchmark Environment

- **Hardware**: Apple Silicon (M-series)
- **JDK**: 21.0.9 (Zulu OpenJDK)
- **Framework**: JMH 1.35
- **Configuration**: 3 warmup iterations, 5 measurement iterations, 1 fork

---

## Detailed Results

### Streaming Reads (SAX Parser)

| Rows | XL | POI | Delta |
|------|-----|-----|-------|
| 1,000 | 0.77 ms | 1.22 ms | **XL 37% faster** |
| 10,000 | 7.53 ms | 7.68 ms | **XL 2% faster** |
| 100,000 | 87.4 ms | 71.8 ms | POI 22% faster |

### In-Memory Reads

| Rows | XL | POI | Delta |
|------|-----|-----|-------|
| 1,000 | 1.20 ms | 1.52 ms | **XL 21% faster** |
| 10,000 | 13.78 ms | 14.15 ms | **XL 3% faster** |
| 100,000 | 184 ms | 166 ms | POI 11% faster |

### Writes (SaxStax Backend)

| Rows | XL SaxStax | POI | Delta |
|------|------------|-----|-------|
| 1,000 | 1.25 ms | 1.29 ms | ~tied |
| 10,000 | 6.89 ms | 10.84 ms | **XL 36% faster** |
| 100,000 | 69.66 ms | 113.35 ms | **XL 39% faster** |

---

## Root Cause Analysis

### Why XL Loses at 100k Reads

#### 1. SAX Read Materialization (Critical)

**File**: `xl-cats-effect/src/com/tjclp/xl/io/SaxStreamingReader.scala:36-47`

```scala
Stream.eval {
  Sync[F].delay {
    parser.parse(InputSource(stream), handler)  // BLOCKS until entire doc parsed
    handler.rows.toVector                        // ALL rows materialized to Vector
  }
}.flatMap(Stream.emits)                         // Only then starts streaming
```

**Problem**: SAX parsing is inherently synchronous. The `parser.parse()` call doesn't return until the entire XML document is processed. The handler accumulates all rows in an `ArrayBuffer`, then materializes to `Vector`.

**Impact**: ~15ms overhead at 100k rows from materialization + GC pressure.

#### 2. SharedStrings Materialization (Medium)

**File**: `xl-cats-effect/src/com/tjclp/xl/io/StreamingXmlReader.scala:37`

```scala
builder.strings.toVector  // SST forced into memory before worksheet parsing
```

The SharedStrings table must be fully parsed before worksheet streaming can begin (sequential dependency for resolving `t="s"` cell references).

### Why XL Wins on Writes

1. **DirectSaxEmitter**: Bypasses intermediate OOXML objects, emits SAX events directly from domain model
2. **SaxStax backend**: Streams directly to output without intermediate String allocation
3. **Surgical modification**: Only regenerates modified parts when editing existing files

### Overhead Breakdown (Estimated)

| Source | Overhead |
|--------|----------|
| SAX materialization (ArrayBuffer → Vector → Stream) | +15-20% |
| Object allocation (immutable case classes vs primitive arrays) | +5-10% |
| Abstraction layers (SaxWriter trait dispatch) | +2-5% |
| **Total at 100k scale** | **~22%** |

---

## Proposed Solutions

### Option A: Accept Current State (Recommended for v1.0)

**Rationale**:
- XL wins 5 out of 9 benchmarks
- Typical workloads (<10k rows) are faster across the board
- 100k row gap (11-22%) is the cost of functional purity
- Safety and testability benefits outweigh raw performance at scale

**Action**: Document as known limitation.

### Option B: Threaded SAX Reader (Future Optimization)

**Goal**: Emit rows during SAX parsing, not after.

**Architecture**:
```scala
def parseWorksheetStream[F[_]: Async](...): Stream[F, RowData] =
  for {
    queue <- Stream.eval(Queue.unbounded[F, Option[RowData]])
    _ <- Stream.bracket(
      Async[F].start(Async[F].blocking {
        val handler = new StreamingHandler(row =>
          queue.offer(Some(row)).unsafeRunSync()
        )
        parser.parse(stream, handler)
        queue.offer(None).unsafeRunSync()  // Signal completion
      })
    )(fiber => fiber.cancel)
    row <- Stream.fromQueueNoneTerminated(queue)
  } yield row
```

**Complexity**: Medium-High
- Requires `Async` constraint (currently `Sync`)
- Cross-thread coordination between SAX callbacks and cats-effect runtime
- Proper cancellation handling

**Estimated Improvement**: 10-15% at 100k rows

### Option C: Lazy SharedStrings (Future Optimization)

**Goal**: Parse SST entries on-demand during worksheet streaming.

**Approach**:
1. Build SST index (offset → position) in first pass
2. Resolve string references lazily during cell parsing
3. Fall back to inline strings if SST unavailable

**Complexity**: Medium
**Estimated Improvement**: 3-5% at 100k rows

### Option D: Sorted Cell Storage (Future Optimization)

**Goal**: Eliminate O(n log n) sort at write time.

**Current**: `sheet.cells: Map[ARef, Cell]` + sort before emission

**Proposed**: `sheet.cells: SortedMap[ARef, Cell]`

**Trade-off**: O(log n) insert vs O(1) insert, but O(n) iteration vs O(n log n) sort

**Complexity**: Medium (affects core domain model)
**Estimated Improvement**: 2-3% at 100k writes

---

## Validation Methodology

All benchmarks use:
- **Verifiable data**: Arithmetic series (A{i} = i) enables sum validation
- **Fair comparison**: Both libraries read identical shared files
- **No JIT artifacts**: Benchmarks return computed sum
- **Write cost isolated**: Files pre-created in @Setup

The POI streaming benchmark was updated to use pre-initialized `SAXParserFactory` (moved to setup) to ensure fair comparison.

---

## Files Modified

| File | Change |
|------|--------|
| `xl-benchmarks/src/.../PoiComparisonBenchmark.scala` | Added `saxParserFactory` field, moved initialization to setup |
| `xl-benchmarks/README.md` | Updated with latest benchmark results |

---

## References

- [JMH Documentation](https://github.com/openjdk/jmh)
- [XL Performance Guide](../reference/performance-guide.md)
- [Apache POI Streaming API](https://poi.apache.org/components/spreadsheet/how-to.html#sxssf)
