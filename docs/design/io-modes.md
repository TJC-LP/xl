# I/O Modes and Architecture

## Overview

XL has **two distinct I/O implementations** with fundamentally different architectures and tradeoffs. This document explains the design decisions, implementation details, and when to use each.

## The Two Modes

### Mode 1: In-Memory (xl-ooxml)

**Implementation**: Traditional OOXML library approach
- Parse ZIP into memory
- Build full scala.xml trees
- Traverse and transform
- Serialize back to XML strings
- Write to ZIP

**Modules**:
- `xl-ooxml`: OOXML data model (OoxmlWorkbook, OoxmlWorksheet, etc.)
- `XlsxReader`: ZIP → Domain model
- `XlsxWriter`: Domain model → ZIP

**Key Files**:
- `xl-ooxml/src/com/tjclp/xl/ooxml/XlsxReader.scala`
- `xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala`
- `xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala`
- `xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala`
- `xl-ooxml/src/com/tjclp/xl/ooxml/SharedStrings.scala`

---

### Mode 2: Streaming (xl-cats-effect)

**Implementation**: Event-based streaming with fs2
- Stream ZIP bytes via fs2-data-xml
- Emit XML events directly
- No intermediate trees
- Constant memory

**Modules**:
- `xl-cats-effect`: Effect interpreters
- `StreamingXmlWriter`: Domain model → XML events → ZIP
- `StreamingXmlReader`: ZIP → XML events → Domain model

**Key Files**:
- `xl-cats-effect/src/com/tjclp/xl/io/Excel.scala`
- `xl-cats-effect/src/com/tjclp/xl/io/ExcelIO.scala`
- `xl-cats-effect/src/com/tjclp/xl/io/StreamingXmlWriter.scala`
- `xl-cats-effect/src/com/tjclp/xl/io/StreamingXmlReader.scala`

---

## Architecture Comparison

### In-Memory Path

Read (simplified):
- `ExcelIO.read(path)` calls `XlsxReader.readWithWarnings`.
- The reader walks the ZIP, parses only *known* parts into `scala.xml.Elem`s, and turns them into OOXML model types.
- It then converts OOXML types into the pure domain `Workbook`, optionally attaching a `SourceContext` for surgical writes.

Write (simplified):
- `ExcelIO.write(workbook, path)` calls `XlsxWriter.writeWith` with `WriterConfig.default`.
- The writer builds a shared `StyleIndex` and optional `SharedStrings` table, converts domain objects into OOXML parts, and serializes them as XML into a new ZIP.
- Output is compact XML (no pretty‑printing) with `Compression.Deflated` by default; `WriterConfig.debug` switches to pretty‑printed XML and STORED compression for debugging.

Characteristics:
- Memory: **O(n)** in the number of cells.
- Features: **Full** (SST, styles, merged cells, RichText, etc.).
- Deterministic output: stable ordering for reliable diffs.

---

### Streaming Path

Write (`writeStreamTrue` / `writeStreamsSeqTrue`):
- Static parts (`[Content_Types].xml`, workbook relationships, minimal `styles.xml`) are written once up front.
- For each streamed `RowData`, `StreamingXmlWriter` emits XML events directly to a `ZipOutputStream` without building intermediate XML trees.
- Output is compact XML with `Compression.Deflated` by default; by design it uses inline strings (no SST) and a minimal style set.

Read (`readStream` / `readSheetStream` / `readStreamByIndex`):
- The ZIP is opened as a `ZipFile`, and the target worksheet entry is streamed through fs2‑data‑xml as a stream of XML events.
- Those events are converted into `RowData` records on the fly.
- The `SharedStrings` table (if present) is parsed once up front and held in memory; worksheet data itself is streamed.

Characteristics:
- Memory: **O(1)** for worksheet data (plus the in‑memory SST and minimal bookkeeping).
- Features:
  - Write: inline strings only, default styles, no merged cells or advanced sheet metadata.
  - Read: values and basic types; you typically use it for ETL/analytics rather than formatting‑preserving workflows.

```
Read:
User calls readStream(path)
  ↓
Open ZIP file
  ↓
For worksheet entry:
  1. getInputStream(entry)
  2. ❌ readAllBytes() ← BUG! Materializes in memory
  3. Stream.emits(bytes)
  4. fs2-data-xml parser
  5. Convert events → RowData
  ↓
Return Stream[F, RowData]

Memory: ❌ O(n) (should be O(1), see P6.6)
Time: Fast (streaming parser)
Features: Limited (reads values only, minimal style info)
```

---

## Design Decisions

### ADR-011: Why Two Modes Instead of One?

**Decision**: Maintain both in-memory and streaming implementations

**Context**:
- Small files need full features (SST, styles, formulas)
- Large files need constant memory (streaming)
- No single implementation satisfies both

**Alternatives Considered**:
1. **Streaming only**: Would lose SST/styles (larger files, no formatting)
2. **In-memory only**: Would OOM on large files
3. **Two-phase streaming**: Adds complexity, still under development (P7.5)

**Chosen**: Two modes with clear guidance on when to use each

**Consequences**:
- ✅ Best-of-both-worlds (full features OR constant memory)
- ✅ Users choose based on needs
- ❌ Two implementations to maintain
- ❌ Some confusion about which to use

**Mitigation**: Clear documentation (this doc, performance-guide.md)

---

### ADR-012: Why Not Streaming-First?

**Decision**: In-memory is the **default** API, streaming is opt-in

**Rationale**:
1. **Most use cases are small** (<10k rows, full features needed)
2. **Streaming has limitations** (no SST/styles currently)
3. **Streaming read is broken** (P6.6 - uses readAllBytes)
4. **In-memory is simpler** (easier API, full features)

**Consequences**:
- ✅ Users get full features by default
- ✅ Simple API for common cases
- ❌ Users might not discover streaming option
- ❌ Default might not scale for their use case

**Mitigation**: Prominent documentation in README (mode selection table)

---

### ADR-013: Streaming Read Bug - Why Not Fixed Yet?

**Status**: Historical only – the original streaming reader used `readAllBytes()` and was O(n) in memory. As of P6.6 it has been rewritten on top of `fs2.io.readInputStream` and fs2‑data‑xml and now achieves the same constant‑memory characteristics as the streaming writer.

Today:
- `ExcelIO.readStream` / `readSheetStream` / `readStreamByIndex` are safe to use for large files.
- The remaining trade‑off is **feature coverage**, not memory use (you do not get full styling/metadata preservation from the streaming APIs).

---

## Feature Support Matrix

| Feature | In-Memory | Streaming Write | Streaming Read (Fixed) |
|---------|-----------|-----------------|------------------------|
| **SST** | ✅ Full | ❌ None (inline only) | ✅ Full |
| **Styles** | ✅ Full | ⚠️ Default only | ⚠️ Minimal |
| **Formulas** | ✅ Store | ✅ Store | ✅ Read |
| **Merged Cells** | ⚠️ Track (not serialized) | ❌ No | ❌ No |
| **Rich Text** | ✅ Full | ✅ Full | ✅ Full |
| **Column/Row Props** | ⚠️ Track (not serialized) | ❌ No | ❌ No |
| **Memory** | O(n) | O(1) | O(1) after P6.6 |
| **Speed** | ~88k rows/sec | ~88k rows/sec | ~55k rows/sec |

**Legend**:
- ✅ Full support
- ⚠️ Partial (tracked but not serialized, or limitations)
- ❌ Not supported

---

## Implementation Details

### In-Memory: Why scala.xml?

**Decision**: Use stdlib `scala.xml.Elem` for in-memory trees

**Pros**:
- Stdlib (no dependencies)
- Familiar API
- Easy to debug (inspect trees in REPL)
- Pattern matching on structure

**Cons**:
- Mutable internals (but we don't expose mutability)
- Memory overhead (full tree representation)
- Not streaming-friendly

**Alternative Considered**: Build custom immutable XML AST
**Rejected Because**: Reinventing stdlib, marginal benefits

---

### Streaming: Why fs2-data-xml?

**Decision**: Use fs2-data-xml for event streaming

**Pros**:
- True streaming (no tree building)
- Constant memory
- Backpressure support
- Pure functional (fs2)

**Cons**:
- External dependency
- Event-based API (more complex than tree)
- Harder to debug (can't inspect tree)

**Alternative Considered**: Use javax.xml.stream.XMLStreamWriter
**Rejected Because**: Imperative, side-effecting, not fs2-compatible

---

## Why Not a Unified Implementation?

**Could we make streaming support full features?**

**Challenge 1**: SST requires string deduplication
- Need to see all strings before writing sharedStrings.xml
- But [Content_Types].xml written first (before strings known)
- **Solution**: Two-phase approach (P7.5) or optimistic SST inclusion

**Challenge 2**: Styles require deduplication across workbook
- Need to merge all sheet style registries
- Assign global indices
- Write styles.xml with unified index
- **Solution**: Two-phase or incremental registry (complex)

**Challenge 3**: ZIP part ordering
- OOXML doesn't require specific order, but:
  - [Content_Types].xml typically first
  - Relationships typically before referenced parts
  - Changing order might break some readers (untested)

**Conclusion**: Streaming with full features is possible (P7.5) but requires:
- Two-pass approach (scan data, write with indices)
- OR disk-backed registries for SST/styles
- OR optimistic overhead (include SST/styles even if empty)

**Timeline**: 3-4 weeks of implementation (deferred to post-MVP)

---

## Testing Strategy

### In-Memory Tests
```scala
test("in-memory round-trip preserves all data"):
  val original = createComplexWorkbook()  // Styles, SST, formulas, merged

  val path = Files.createTempFile("test", ".xlsx")
  ExcelIO.instance.write[IO](original, path).unsafeRunSync()

  val reloaded = ExcelIO.instance.read[IO](path).map(_.toOption.get).unsafeRunSync()

  assertEquals(reloaded.sheets.size, original.sheets.size)
  assertEquals(reloaded.sheets(0).cells, original.sheets(0).cells)
```

### Streaming Write Tests
```scala
test("streaming write uses constant memory"):
  val memBefore = currentHeapUsage()

  Stream.range(1, 1_000_001)
    .map(i => RowData(i, Map(0 -> CellValue.Number(i))))
    .through(Excel.forIO.writeStreamTrue(path, "Data"))
    .compile.drain
    .unsafeRunSync()

  val memAfter = currentHeapUsage()
  val memUsed = (memAfter - memBefore) / (1024 * 1024)  // MB

  assert(memUsed < 50, s"Used $memUsed MB (expected < 50 MB)")
```

### Streaming Read Tests (After P6.6 Fix)
```scala
test("streaming read uses constant memory"):
  val largePath = generateLargeFile(1_000_000)
  val memBefore = currentHeapUsage()

  Excel.forIO.readStream(largePath)
    .take(1000)
    .compile.drain
    .unsafeRunSync()

  val memAfter = currentHeapUsage()
  val memUsed = (memAfter - memBefore) / (1024 * 1024)

  assert(memUsed < 50, s"Used $memUsed MB (expected < 50 MB)")
```

---

## Migration Path

### Current State (As of 2025-11-11)
- In-memory: Production-ready for <100k rows
- Streaming write: Production-ready for >100k rows (minimal styling)
- Streaming read: Broken (use in-memory)

### P6.6: Fix Streaming Read (2-3 days)
- Replace `readAllBytes()` with `fs2.io.readInputStream`
- Achieve true O(1) memory for reads
- Add memory regression tests

### P6.7: Compression Control (1 day)
- Add WriterConfig(compression, prettyPrint)
- Default to DEFLATED + compact (smaller files)
- Keep STORED + pretty for debugging

### P7.5: Two-Phase Streaming Writer (3-4 weeks)
- Support SST and styles in streaming mode
- Two-pass approach: scan → write
- Disk-backed registries for very large datasets
- Achieves O(1) memory with full features

---

## Performance Targets

### In-Memory Mode
- **Small (<10k)**: < 100ms write, < 10MB memory
- **Medium (10k-100k)**: < 2s write, < 100MB memory
- **Large (>100k)**: Not recommended (use streaming)

### Streaming Mode (After Fixes)
- **Write**: O(1) memory (~10MB), 88k rows/sec ✅
- **Read**: O(1) memory (~10MB), 55k rows/sec (after P6.6)
- **Scalability**: 10M+ rows without OOM

---

## Related Documents

- [performance-guide.md](../reference/performance-guide.md) - User guidance
- [streaming-improvements.md](../plan/streaming-improvements.md) - Roadmap
- [purity-charter.md](purity-charter.md) - Why pure core matters
- [STATUS.md](../STATUS.md) - Current limitations

---

## Author

Documented 2025-11-11
