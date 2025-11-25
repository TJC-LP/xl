# WI-17: SAX Streaming Write - Zero-Allocation OOXML Serialization

**Status**: ✅ Complete
**Priority**: Delivered
**Last Updated**: 2025-11-25

> **Archive Notice**: This plan has been fully implemented. SAX streaming write is operational with significantly reduced memory allocation compared to Scala XML tree-building. This document is retained for historical reference.

---

**Dependencies**: WI-15 (SAX read proven) ✅
**Module**: xl-ooxml, xl-cats-effect
**Actual Effort**: ~20 hours
**Result**: SAX streaming write integrated

---

## Problem Statement

### Current Performance Gap (WI-15 Baseline)

**Write Performance** @ 10k rows:
- Apache POI: 10.8ms ± 0.4ms
- XL (Scala XML): 15.4ms ± 1.5ms
- **Gap**: POI is 42% faster

**Memory Allocation** (JMH GC Profiler):
- Apache POI: 7.2 MB per write
- XL (Scala XML): 63.5 MB per write
- **Gap**: XL allocates 8.8x MORE

**Root Cause**: Scala XML tree-building architecture creates massive intermediate allocations:

| Component | Allocations per 10k rows |
|-----------|-------------------------|
| Cell Elem objects | 25-30 MB (100k cells × 250 bytes) |
| Row Elem objects | 8-10 MB (10k rows × 800 bytes) |
| XML tree → String | 15-20 MB (materialization) |
| .toString() on primitives | 5-8 MB (150k+ String allocations) |
| Collection intermediates | 5-10 MB (groupBy, toSeq, zipWithIndex) |
| **Total** | **63.5 MB** (vs POI's 7.2 MB) |

**GC Impact**:
- XL: 28 GC collections per write
- POI: 9 GC collections per write
- XL: 118ms total GC pause time
- POI: 10ms total GC pause time

---

## Proposed Solution: SAX Streaming with Pure FP Wrapper

### Architecture Principle: Effect Isolation

**Core Constraint** (from CLAUDE.md):
> Core (xl-core, xl-ooxml) is 100% pure; only xl-cats-effect has IO

**Solution**: Keep SAX (imperative) in xl-cats-effect, expose pure abstractions in xl-ooxml.

### Design Pattern: Pure Protocol + Effectful Interpreter

#### **xl-ooxml** (Pure Layer)

```scala
// Pure abstraction - no XMLStreamWriter dependency
trait SaxWriter:
  def startElement(name: String): Unit
  def writeAttribute(name: String, value: String): Unit
  def writeCharacters(text: String): Unit
  def endElement(): Unit

// Pure serialization protocol
trait SaxSerializable:
  def writeSax(writer: SaxWriter): Unit

// Domain types implement pure protocol
extension (cell: OoxmlCell)
  def writeSax(writer: SaxWriter): Unit =
    writer.startElement("c")
    writer.writeAttribute("r", cell.ref.toA1)
    cell.styleIndex.foreach(s => writer.writeAttribute("s", s.toString))
    // ... value serialization ...
    writer.endElement()
```

**Key Properties**:
- ✅ 100% pure (no side effects in xl-ooxml)
- ✅ Total functions (no exceptions)
- ✅ Deterministic (attribute ordering in protocol)
- ✅ Testable (mock SaxWriter implementation)

#### **xl-cats-effect** (Effect Layer)

```scala
// Effectful interpreter - wraps javax.xml.stream.XMLStreamWriter
class XmlStreamSaxWriter[F[_]: Sync](underlying: XMLStreamWriter) extends SaxWriter:
  def startElement(name: String): Unit = underlying.writeStartElement(name)
  def writeAttribute(name: String, value: String): Unit = underlying.writeAttribute(name, value)
  def writeCharacters(text: String): Unit = underlying.writeCharacters(text)
  def endElement(): Unit = underlying.writeEndElement()

// Public API in ExcelIO
def writeFast(wb: Workbook, path: Path): F[Unit] =
  Sync[F].delay {
    val out = new BufferedOutputStream(new FileOutputStream(path.toFile))
    try {
      val xmlWriter = createXMLStreamWriter(out)
      val saxWriter = new XmlStreamSaxWriter(xmlWriter)
      SaxStreamingWriter.writeWorkbook(wb, saxWriter)
    } finally out.close()
  }
```

**Key Properties**:
- ✅ Effect boundary explicit (Sync[F].delay)
- ✅ Resource safety (try-finally, bracket if needed)
- ✅ XMLStreamWriter encapsulated (not exposed to xl-ooxml)

---

## Expected Performance Gains

### Allocation Reduction

**Current (Scala XML)**:
- 63.5 MB allocations per 10k row write
- 28 GC collections
- 118ms GC pause time

**After SAX Streaming**:
- **8-12 MB** allocations (buffer + ZIP compression only)
- **5-8 GC collections** (6x reduction)
- **15-20ms GC pause time** (6x reduction)

### Performance Projection

Based on SAX read optimization (3.8x speedup from eliminating XmlEvent allocations):

| Rows | Current XL | After SAX | POI | XL vs POI |
|------|------------|-----------|-----|-----------|
| **1,000** | 1.97ms | **0.9-1.2ms** | 1.38ms | **XL 15-35% faster** ✨ |
| **10,000** | 15.4ms | **6-9ms** | 10.8ms | **XL 17-45% faster** ✨ |

**Confidence**: High (85%) - same optimization pattern as read path

---

## Implementation Phases

### **Phase A: Foundation** (4-6 hours)

**Create**: `xl-ooxml/src/com/tjclp/xl/ooxml/SaxWriter.scala`

```scala
package com.tjclp.xl.ooxml

/** Pure abstraction for SAX-style XML writing */
trait SaxWriter:
  def startDocument(): Unit
  def endDocument(): Unit
  def startElement(name: String): Unit
  def startElement(name: String, namespace: String): Unit
  def writeAttribute(name: String, value: String): Unit
  def writeCharacters(text: String): Unit
  def endElement(): Unit
  def flush(): Unit

/** Types that can serialize via SAX */
trait SaxSerializable:
  def writeSax(writer: SaxWriter): Unit

object SaxWriter:
  /** Ensure attribute determinism - sort before writing */
  def withAttributes(writer: SaxWriter, attrs: (String, String)*)(body: => Unit): Unit =
    attrs.sortBy(_._1).foreach { (k, v) => writer.writeAttribute(k, v) }
    body
```

**Tests**: Create `SaxWriterSpec.scala` with mock implementation:
```scala
class MockSaxWriter extends SaxWriter:
  val events = mutable.ArrayBuffer[String]()
  def startElement(name: String): Unit = events += s"<$name>"
  def writeAttribute(name: String, value: String): Unit = events += s" $name=$value"
  // ...

test("SaxWriter mock captures events") {
  val mock = MockSaxWriter()
  mock.startElement("row")
  mock.writeAttribute("r", "1")
  mock.endElement()
  assertEquals(mock.events.mkString, "<row> r=1</row>")
}
```

---

### **Phase B: Worksheet Serialization** (10-14 hours)

**Modify**: `xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala`

**Add** (keep existing toXml for backward compatibility):

```scala
extension (ws: OoxmlWorksheet)
  def writeSax(writer: SaxWriter, sst: Option[SharedStrings], styleRemapping: Map[Int, Int]): Unit =
    writer.startElement("worksheet")
    writer.writeAttribute("xmlns", nsSpreadsheetML)
    // ... namespace attributes ...

    writer.startElement("sheetData")

    // Group cells by row (reuse existing logic)
    val cellsByRow = // ... same as current ...

    cellsByRow.foreach { case (rowIdx, cells) =>
      writer.startElement("row")
      SaxWriter.withAttributes(writer,
        "r" -> rowIdx.toString,
        // ... other row attributes ...
      ) {
        cells.foreach { cell =>
          cell.writeSax(writer, sst, styleRemapping)
        }
      }
      writer.endElement() // row
    }

    writer.endElement() // sheetData

    // ... other worksheet elements ...
    writer.endElement() // worksheet

extension (cell: OoxmlCell)
  def writeSax(writer: SaxWriter, sst: Option[SharedStrings], styleRemapping: Map[Int, Int]): Unit =
    writer.startElement("c")

    val attrs = Seq.newBuilder[(String, String)]
    attrs += ("r" -> cell.ref.toA1)
    val globalStyleIdx = cell.styleId.flatMap(localId =>
      styleRemapping.get(localId.value).orElse(Some(0))
    )
    globalStyleIdx.foreach(idx => attrs += ("s" -> idx.toString))

    val (cellType, value) = cell.value match
      case CellValue.Text(s) =>
        sst.flatMap(_.indexOf(s)) match
          case Some(idx) => ("s", CellValue.Text(idx.toString))
          case None => ("inlineStr", cell.value)
      case _ => // ... existing logic ...

    if cellType.nonEmpty then attrs += ("t" -> cellType)

    SaxWriter.withAttributes(writer, attrs.result()*)

    // Write value element
    value match
      case CellValue.Text(s) if cellType == "inlineStr" =>
        writer.startElement("is")
        writer.startElement("t")
        writer.writeCharacters(s)
        writer.endElement() // t
        writer.endElement() // is
      case CellValue.Text(s) | CellValue.Number(s) =>
        writer.startElement("v")
        writer.writeCharacters(s.toString)
        writer.endElement() // v
      case CellValue.RichText(rt) =>
        writeRichTextSax(writer, rt)
      case _ => ()

    writer.endElement() // c
```

**Tests**:
- Modify `OoxmlRoundTripSpec` to compare SAX output vs Scala XML output (byte-identical after parse)
- Add `SaxWorksheetSerializationSpec` with golden file tests

**Impact**: 25-30 MB allocation reduction (biggest win!)

---

### **Phase C: Styles Serialization** (8-10 hours)

**Modify**: `xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala`

**Add**:

```scala
extension (styles: OoxmlStyles)
  def writeSax(writer: SaxWriter): Unit =
    writer.startElement("styleSheet")
    writer.writeAttribute("xmlns", nsSpreadsheetML)

    // numFmts (optional)
    if styles.index.numFmts.nonEmpty then
      writer.startElement("numFmts")
      SaxWriter.withAttributes(writer, "count" -> styles.index.numFmts.size.toString)
      styles.index.numFmts.sortBy(_._1).foreach { (id, fmt) =>
        writer.startElement("numFmt")
        writer.writeAttribute("formatCode", fmt.formatCode)
        writer.writeAttribute("numFmtId", id.toString)
        writer.endElement() // numFmt
      }
      writer.endElement() // numFmts

    // Fonts
    writer.startElement("fonts")
    writer.writeAttribute("count", styles.index.fonts.size.toString)
    styles.index.fonts.foreach(font => writeFontSax(writer, font))
    writer.endElement() // fonts

    // Fills, Borders, cellXfs...
    // ... similar pattern for each component ...

    writer.endElement() // styleSheet

private def writeFontSax(writer: SaxWriter, font: Font): Unit =
  writer.startElement("font")
  // Bold
  if font.bold then
    writer.startElement("b")
    writer.endElement()
  // Italic, size, color, name...
  writer.endElement() // font
```

**Tests**:
- `StyleIntegrationSpec` with SAX comparison
- Verify deduplication still works (fonts/fills/borders)

**Impact**: 2-5 MB allocation reduction

---

### **Phase D: SharedStrings Serialization** (6-8 hours)

**Modify**: `xl-ooxml/src/com/tjclp/xl/ooxml/SharedStrings.scala`

**Add**:

```scala
extension (sst: SharedStrings)
  def writeSax(writer: SaxWriter): Unit =
    writer.startElement("sst")
    SaxWriter.withAttributes(writer,
      "xmlns" -> nsSpreadsheetML,
      "count" -> sst.totalCount.toString,
      "uniqueCount" -> sst.strings.size.toString
    )

    sst.strings.foreach {
      case Left(text) =>
        writer.startElement("si")
        writer.startElement("t")
        if needsXmlSpacePreserve(text) then
          writer.writeAttribute("xml:space", "preserve")
        writer.writeCharacters(text)
        writer.endElement() // t
        writer.endElement() // si

      case Right(richText) =>
        writer.startElement("si")
        richText.runs.foreach { run =>
          writeTextRunSax(writer, run)
        }
        writer.endElement() // si
    }

    writer.endElement() // sst

private def writeTextRunSax(writer: SaxWriter, run: TextRun): Unit =
  writer.startElement("r")

  // rPr (font properties)
  run.font.foreach { f =>
    writer.startElement("rPr")
    if f.bold then
      writer.startElement("b")
      writer.endElement()
    // ... other font properties ...
    writer.endElement() // rPr
  }

  // Text content
  writer.startElement("t")
  if needsXmlSpacePreserve(run.text) then
    writer.writeAttribute("xml:space", "preserve")
  writer.writeCharacters(run.text)
  writer.endElement() // t

  writer.endElement() // r
```

**Impact**: 2-8 MB allocation reduction (depends on RichText usage)

---

### **Phase E: Workbook and Other Parts** (4-6 hours)

**Files**:
- `OoxmlWorkbook.scala` - workbook.xml
- `ContentTypes.scala` - [Content_Types].xml
- `Relationships.scala` - .rels files
- `VmlDrawing.scala` - VML for comment indicators

**Pattern**: Same extension method approach (writeSax)

**Impact**: 1-3 MB allocation reduction

---

### **Phase F: Integration** (6-8 hours)

**Modify**: `xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala`

**Add**:

```scala
def writeWithSax(
  wb: Workbook,
  target: OutputTarget,
  config: WriterConfig
): Either[XLError, Unit] =
  // ... existing ZIP setup ...

  // NEW: Use SAX for all parts
  def writePartSax[A: SaxSerializable](path: String, data: A): Unit =
    zip.putNextEntry(new ZipEntry(path))
    val buffered = new BufferedOutputStream(zip)  // Buffer for XMLStreamWriter
    val xmlWriter = createXMLStreamWriter(buffered)
    val saxWriter = new PureSaxWriter(xmlWriter)  // Wrap for safety

    data.writeSax(saxWriter)

    xmlWriter.flush()
    buffered.flush()  // Don't close - ZIP stream remains open
    zip.closeEntry()

  // Write all parts with SAX
  writePartSax("xl/workbook.xml", ooxmlWorkbook)
  writePartSax("xl/styles.xml", ooxmlStyles)
  ooxmlSST.foreach(sst => writePartSax("xl/sharedStrings.xml", sst))
  sheets.foreach { case (sheet, idx) =>
    writePartSax(s"xl/worksheets/sheet${idx + 1}.xml", sheet)
  }
  // ...
```

**Create**: `xl-cats-effect/src/com/tjclp/xl/io/SaxStreamingWriter.scala`

```scala
package com.tjclp.xl.io

import com.tjclp.xl.ooxml.{SaxWriter, OoxmlWorksheet, OoxmlStyles, SharedStrings}
import javax.xml.stream.{XMLOutputFactory, XMLStreamWriter}
import java.io.OutputStream

/** Pure SAX writer implementation (wraps imperative XMLStreamWriter) */
class PureSaxWriter(underlying: XMLStreamWriter) extends SaxWriter:
  def startDocument(): Unit = underlying.writeStartDocument("UTF-8", "1.0")
  def endDocument(): Unit = underlying.writeEndDocument()
  def startElement(name: String): Unit = underlying.writeStartElement(name)
  def startElement(name: String, namespace: String): Unit =
    underlying.writeStartElement("", name, namespace)
  def writeAttribute(name: String, value: String): Unit =
    underlying.writeAttribute(name, value)
  def writeCharacters(text: String): Unit = underlying.writeCharacters(text)
  def endElement(): Unit = underlying.writeEndElement()
  def flush(): Unit = underlying.flush()

object SaxStreamingWriter:
  def createXMLStreamWriter(out: OutputStream): XMLStreamWriter =
    val factory = XMLOutputFactory.newInstance()
    factory.setProperty(XMLOutputFactory.IS_REPAIRING_NAMESPACES, false)
    factory.createXMLStreamWriter(out, "UTF-8")
```

**Tests**:
- `SaxStreamingWriterSpec` with round-trip verification
- Byte-compare SAX output vs Scala XML output

---

### **Phase G: Public API** (2-3 hours)

**Modify**: `xl-cats-effect/src/com/tjclp/xl/io/ExcelIO.scala`

**Add**:

```scala
/** Fast write using SAX streaming (40-60% faster, zero-allocation) */
def writeFast(wb: Workbook, path: Path): F[Unit] =
  Sync[F].delay(XlsxWriter.writeWithSax(wb, path, WriterConfig.default))
    .flatMap {
      case Right(_) => Async[F].unit
      case Left(err) => Async[F].raiseError(new Exception(s"Failed to write XLSX: ${err.message}"))
    }

/** Fast write with custom config */
def writeFastWith(wb: Workbook, path: Path, config: WriterConfig): F[Unit] =
  Sync[F].delay(XlsxWriter.writeWithSax(wb, path, config))
    .flatMap {
      case Right(_) => Async[F].unit
      case Left(err) => Async[F].raiseError(new Exception(s"Failed to write XLSX: ${err.message}"))
    }
```

**Decision**: Should `write()` default to SAX or Scala XML?

**Option 1 (Recommended)**: Switch default to SAX
```scala
def write(wb: Workbook, path: Path): F[Unit] =
  writeFast(wb, path)  // Use SAX by default

def writeCompat(wb: Workbook, path: Path): F[Unit] =
  writeWith(wb, path, WriterConfig.default)  // Scala XML for compatibility
```

**Option 2 (Conservative)**: Keep Scala XML default, add opt-in SAX
```scala
def write(wb: Workbook, path: Path): F[Unit] =
  writeWith(wb, path, WriterConfig.default)  // Scala XML (existing)

def writeFast(wb: Workbook, path: Path): F[Unit] =
  // SAX (opt-in for performance)
```

---

### **Phase H: Validation** (4-6 hours)

**Benchmark Suite**:
```bash
# Full comparison
./mill xl-benchmarks.runJmh 'PoiComparisonBenchmark.*(xlWrite|poiWrite)' \
  -p rows=1000,10000 -wi 5 -i 10 -f 1 -prof gc

# Expected results:
# XL write @ 10k: 6-9ms (vs POI 10.8ms)
# XL allocations: 8-12 MB (vs POI 7.2 MB)
```

**Test Suite**:
```bash
./mill __.test  # All 731+ tests must pass
```

**Regression Tests**:
- Round-trip: write with SAX → read → compare to original
- Byte-comparison: SAX output vs Scala XML output (parse both, compare trees)
- Excel compatibility: Open generated files in Excel, verify no corruption

---

## Maintaining Backward Compatibility

### Dual Serialization Paths

**For Tests** (keep Scala XML):
```scala
// Tests can still use Elem-based assertions
test("worksheet has correct structure") {
  val elem = worksheet.toXml
  assertEquals(elem.label, "worksheet")
  assertEquals((elem \ "sheetData" \ "row").size, 10)
}
```

**For Production** (use SAX):
```scala
// Production code uses fast path
ExcelIO.instance[IO].writeFast(workbook, path).unsafeRunSync()
```

### Migration Strategy

**Phase 1**: Add SAX alongside Scala XML (both work)
- xl-ooxml: has both `toXml` and `writeSax`
- ExcelIO: has both `write()` (Scala XML) and `writeFast()` (SAX)
- Users can choose

**Phase 2** (6-12 months later): Deprecate Scala XML
- Mark `toXml` as deprecated
- Make `write()` use SAX by default
- Keep `toXml` for tests only

**Phase 3** (eventual): Remove Scala XML
- Delete all `toXml` methods
- Tests use SAX + byte comparison

---

## Risk Assessment

### **HIGH RISK: Attribute Ordering**

**Issue**: XMLStreamWriter doesn't sort attributes automatically
**Mitigation**: `SaxWriter.withAttributes()` helper sorts before writing
**Test**: DeterminismSpec verifies byte-identical output across writes

### **MEDIUM RISK: Namespace Handling**

**Issue**: SAX namespace API is complex (prefix binding)
**Mitigation**: Use `writeDefaultNamespace()` and `setDefaultNamespace()` explicitly
**Test**: WorksheetNamespaceSpec verifies correct xmlns attributes

### **MEDIUM RISK: Test Migration**

**Issue**: 731+ tests currently use `toXml: Elem`
**Mitigation**: Keep both paths, tests use Elem, production uses SAX
**Test**: Round-trip tests compare SAX vs Scala XML byte-for-byte

### **LOW RISK: Performance Regression**

**Issue**: SAX could theoretically be slower (buffer overhead)
**Mitigation**: Benchmark each phase, revert if slower
**Test**: Performance regression test ensures writes stay <10ms @ 10k rows

---

## Success Criteria

### Definition of Done

- [ ] All 731+ tests passing
- [ ] Write benchmark @ 10k rows: <9.5ms (beat POI by 12%+)
- [ ] GC allocations: <12 MB per write (6x reduction from 63.5 MB)
- [ ] GC collections: <10 per write (3x reduction from 28)
- [ ] Round-trip tests: SAX write → read → byte-identical
- [ ] Excel compatibility: Files open in Excel without errors
- [ ] Documentation updated: CLAUDE.md, STATUS.md, README.md

### Acceptance Criteria

**Minimum** (must achieve):
- XL write <10.5ms @ 10k rows (within 3% of POI)
- Zero test regressions
- Excel files open correctly

**Target** (goal):
- XL write <9.5ms @ 10k rows (beat POI by 12%)
- Allocations <10 MB (match POI)

**Stretch** (aspirational):
- XL write <8.5ms @ 10k rows (beat POI by 21%)
- Allocations <8 MB (beat POI!)

---

## Alternative Considered: fs2-data-xml

**Why NOT fs2-data-xml**:
1. **Allocation overhead**: 300k XmlEvent objects = 18-30 MB (vs SAX's 8-10 MB)
2. **Render step cost**: Pattern-matching events = 2-5ms overhead
3. **Performance**: Expected 10-12ms (competitive, but doesn't beat POI)
4. **Benefit**: Pure FP (but so is SaxWriter abstraction!)

**Conclusion**: SAX with pure wrapper gives BOTH purity AND maximum performance.

---

## Timeline and Effort

### Phased Delivery

**MVP** (Phases A-B-F-H): 24-32 hours
- Foundation + Worksheet only
- Expected: 8-10ms @ 10k rows (17-26% faster than POI)
- Ships partial SAX path (biggest win)

**Complete** (All phases): 44-58 hours
- All parts using SAX
- Expected: 6-9ms @ 10k rows (17-45% faster than POI)
- Total dominance achieved

### Recommended Approach

**Session-based implementation** (5 sessions × 6-8 hours):
1. Foundation + Worksheet (50%)
2. Worksheet complete + Styles
3. SharedStrings + Workbook
4. Integration + Testing
5. Validation + Documentation

---

## Post-Implementation Messaging

### **Before** (WI-15):
- Reads: XL 35% faster than POI ✨
- Writes: POI 42% faster ❌

### **After** (WI-17):
- Reads: XL 35% faster than POI ✨
- Writes: **XL 20-40% faster than POI** ✨✨

### **Positioning**:
"XL is the FASTEST Excel library in any language - pure functional architecture with zero-allocation SAX streaming achieves dominance over imperative Apache POI across ALL operations."

---

## References

- **SAX Read Optimization**: Commit `aca827d` - 3.8x speedup, proven pattern
- **Allocation Analysis**: WI-15 GC profiling - 8.8x allocation gap identified
- **Pure FP Protocol**: Similar to Cats Effect IOApp (pure algebra + effectful runner)
- **Scala XML Limitations**: scala.xml.Elem inherently allocates tree structures

---

## Open Questions for Implementation

1. **Should SaxWriter be synchronous or support F[_]?**
   - Current design: Synchronous (wrapped in Sync[F].delay at call site)
   - Alternative: Make SaxWriter methods return F[Unit]
   - Recommendation: Keep synchronous (simpler, effect boundary at ExcelIO level)

2. **How to handle preserved metadata (surgical writes)?**
   - Option A: Parse preserved parts back to domain, serialize with SAX
   - Option B: Copy preserved parts verbatim (current approach)
   - Recommendation: Option B (maintains current surgical write benefits)

3. **Should we keep toXml methods indefinitely or deprecate?**
   - Option A: Keep both forever (dual API)
   - Option B: Deprecate after 6 months, remove after 12 months
   - Recommendation: Option B (reduce maintenance burden)

---

## Dependencies

**Before WI-17**:
- ✅ WI-15: SAX read optimization complete (proves SAX pattern works)
- ✅ WI-15: JMH benchmark infrastructure complete
- ✅ WI-15: GC profiling identified bottleneck (allocation gap)

**Parallel with WI-17** (safe):
- WI-07: Formula parsing (xl-evaluator, no conflicts)
- WI-08: Formula evaluation (xl-evaluator, no conflicts)

**After WI-17**:
- WI-18: Streaming improvements can build on SAX foundation
- WI-19: Query API benefits from SAX performance

---

## Conclusion

SAX streaming write with pure FP wrapper is the **optimal solution** for achieving total performance dominance while maintaining XL's purity charter. The pattern is proven (SAX reads achieved 3.8x speedup), the allocation gap is identified (8.8x difference), and the expected gain is substantial (40-60% speedup).

**Status**: Design complete, ready for implementation when prioritized.
**Recommendation**: Implement after formula evaluation (WI-07, WI-08) to balance features vs performance work.
