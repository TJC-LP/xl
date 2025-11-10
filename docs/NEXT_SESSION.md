# Next Session Guide - XL Project

**Last Updated**: 2025-11-10 (End of Session 2)
**Current State**: 85% complete, 171/171 tests passing ✅
**Branch**: `main` (all work merged)

---

## What's Been Completed (Session 2)

### P5: Complete Streaming I/O (~800 LOC)
✅ True streaming write with fs2-data-xml
✅ True streaming read with pull parsing
✅ Multi-sheet sequential streaming
✅ Arbitrary sheet access (by index or name)
✅ 100k rows in ~3s, ~50MB constant memory

### P4: Complete Formatting (~510 LOC)
✅ StyleRegistry for coordinated style tracking
✅ High-level API: `withCellStyle`, `withRangeStyle`, `getCellStyle`
✅ `Patch.SetCellStyle` for declarative styling
✅ Unified style index with per-sheet remapping
✅ DateTime serialization with Excel serial numbers
✅ **Styles and dates work in Excel!**

### Documentation (~1000 LOC)
✅ LIMITATIONS.md - Comprehensive roadmap
✅ Session docs (STATUS.md, NEXT_STEPS.md, SESSION_SUMMARY.md)
✅ Updated README with streaming examples
✅ .gitignore for test files

---

## Immediate Priorities (Next Session Start Here)

### Priority 1: Update Documentation (~2 hours)

**README.md needs style examples** - Users need to see how to use the new formatting API:

```scala
// Add to README.md "Styling Example" section:
val headerStyle = CellStyle.default
  .withFont(Font("Arial", 14.0, bold = true))
  .withFill(Fill.Solid(Color.Rgb(0xFFCCCCCC)))
  .withAlign(Align(HAlign.Center, VAlign.Middle))

val sheet = Sheet("Report")
  .put(cell"A1", "Revenue Report")
  .withCellStyle(cell"A1", headerStyle)
  // OR with range styling:
  .withRangeStyle(range"A1:D1", headerStyle)
```

**CLAUDE.md needs updates** - Add sections for:
- StyleRegistry usage patterns
- Streaming patterns (reference https://fs2.io/#/guide)
- Multi-sheet patterns with `writeStreamsSeqTrue`

**Update STATUS.md** - Current file on main is outdated (says streaming not complete)

---

### Priority 2: Implement Merged Cells XML Serialization (~3 hours, ~40 LOC)

**Impact**: HIGH - Currently tracked but not written to XML

**Current State**:
- `Sheet.mergedRanges: Set[CellRange]` exists
- `Patch.Merge(range)` and `Patch.Unmerge(range)` work
- BUT: Not serialized to `<mergeCells>` in worksheet.xml

**Implementation**:
File: `xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala`

Add to `OoxmlWorksheet.toXml` (around line 60):
```scala
// After sheetData element
val mergeCellsElem = if mergedRanges.nonEmpty then
  Some(elem("mergeCells", "count" -> mergedRanges.size.toString)(
    mergedRanges.toSeq.sortBy(_.start.toA1).map { range =>
      elem("mergeCell", "ref" -> range.toA1)()
    }*
  ))
else None

// Include in worksheet children
elem("worksheet", ...)(
  sheetData +: mergeCellsElem.toSeq*
)
```

**Add to domain conversion**:
- Extract `sheet.mergedRanges` when creating OoxmlWorksheet
- Pass through fromDomain/fromDomainWithSST

**Tests**:
- Verify `<mergeCells>` appears in XML
- Verify Excel recognizes merged cells
- Round-trip preservation

**Reference**: See docs/plan/11-ooxml-mapping.md for spec

---

### Priority 3: Column/Row Properties XML Serialization (~4 hours, ~60 LOC)

**Impact**: MEDIUM - Width, height, hidden state not preserved

**Files to change**:
- `xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala`

Add `<cols>` element:
```xml
<cols>
  <col min="1" max="1" width="15.5" customWidth="1"/>
</cols>
```

Add row attributes:
```xml
<row r="1" ht="25.5" customHeight="1">
  ...
</row>
```

**Extract from**:
- `sheet.columnProperties: Map[Column, ColumnProperties]`
- `sheet.rowProperties: Map[Row, RowProperties]`

**Reference**: docs/plan/11-ooxml-mapping.md

---

### Priority 4: Performance Profiling (~2-3 hours)

**Goal**: Identify bottlenecks and optimize hot paths

**Tasks**:
1. Profile 100k row write/read
2. Identify allocation hotspots
3. Consider additional inline annotations
4. Benchmark vs Apache POI

**Tools**:
- JMH benchmarks
- VisualVM profiling
- Memory profiling

**Target**: Beat POI by 3-5x on throughput, 5-10x on memory

---

## Medium-Term Roadmap (P6-P7, ~4-6 weeks)

### P6: Codecs & Named Tuples (2-3 weeks, ~500 LOC)

**High value**: Eliminates boilerplate for data binding

**Goal**:
```scala
case class Person(name: String, age: Int, salary: BigDecimal)

given Codec[Person] = Codec.derived

// Read: Stream[F, RowData] → Stream[F, Person]
excel.readStream(path).through(Codec[Person].decode)

// Write: Stream[F, Person] → XLSX
people.through(Codec[Person].encode)
  .through(excel.writeStreamTrue(path, "People"))
```

**Tasks**:
1. Type-class derivation with Scala 3 deriving
2. Header row binding
3. Column mapping (by index or name)
4. Error accumulation for invalid rows
5. Optional field handling

**Reference**: docs/plan/09-codecs-and-named-tuples.md

---

### P7: Advanced Macros (1-2 weeks, ~200 LOC)

**Path macro** for symbolic references:
```scala
object Paths:
  val totalRevenue = path"Summary!B10"

sheet.put(Paths.totalRevenue, formula"=SUM(Sales!A:A)")
```

**Style literal** for inline formatting:
```scala
val headerStyle = style"font-weight: bold; background: #CCCCCC"
```

**Reference**: docs/plan/17-macros-and-syntax.md

---

## Long-Term Roadmap (P8-P11, ~6-12 months)

### P8: Drawings (2-3 weeks)
- Image embedding
- Shapes and text boxes
- Drawing relationships

### P9: Charts (4-6 weeks)
- 15+ chart types
- Complex XML structure
- Chart relationships

### P10: Tables & Advanced (3-4 weeks)
- Excel Tables
- Pivot Tables
- Conditional formatting
- Data validation

### P11: Security & Safety (1-2 weeks) - **Critical before 1.0**
- ZIP bomb detection
- XXE prevention
- Formula injection guards
- Resource limits
- Security audit

---

## Known Issues & Quirks

### 1. Merged Cells Not Serialized
**Impact**: Merge ranges tracked but lost on write
**Fix**: Priority 2 above (~3 hours)

### 2. Column/Row Properties Not Serialized
**Impact**: Width, height, hidden state lost
**Fix**: Priority 3 above (~4 hours)

### 3. Theme Colors Return 0
**Impact**: Theme-based colors don't work
**Workaround**: Use `Color.Rgb()` directly
**Fix**: Parse theme1.xml (~1-2 days)

### 4. No SST in Streaming Write
**Impact**: Larger file sizes for repeated strings
**Workaround**: Acceptable (disk is cheap)
**Fix**: Two-pass or online SST building (~4-5 days)

### 5. Formula Evaluation Not Implemented
**Impact**: Cannot calculate formula results
**Workaround**: Let Excel recalculate on open
**Fix**: P11 (~60-90 days - very complex)

---

## Important Context for Next Session

### Build Commands
```bash
# Compile everything
./mill __.compile

# Run all tests (171 tests)
./mill __.test

# Test specific module
./mill xl-core.test
./mill xl-ooxml.test
./mill xl-cats-effect.test

# Format code
./mill __.reformat

# Check formatting (CI)
./mill __.checkFormat

# Clean
./mill clean
```

### Project Structure
```
xl/
├── xl-core/          # Pure domain (Cell, Sheet, Workbook, Patch, Style, StyleRegistry)
├── xl-macros/        # Compile-time DSL (cell"", range"", batch put, formatted literals)
├── xl-ooxml/         # Pure XML serialization (XlsxReader, XlsxWriter, OOXML parts)
├── xl-cats-effect/   # IO interpreters (Excel[F], ExcelIO, StreamingXmlReader/Writer)
├── xl-evaluator/     # Formula evaluator (future)
└── xl-testkit/       # Test infrastructure (future)
```

### Key Files to Reference

**Planning**:
- `docs/plan/18-roadmap.md` - Full implementation roadmap (P0-P11)
- `docs/LIMITATIONS.md` - Current limitations with effort estimates
- `docs/STATUS.md` - Detailed current state (may need update)

**Implementation**:
- `CLAUDE.md` - AI assistant context and patterns
- `README.md` - User documentation

**Session History**:
- `docs/SESSION_SUMMARY.md` - Session 1 achievements
- This file - Session 2 context

---

## Style API Quick Reference

```scala
// Create styles
val boldStyle = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))
val redStyle = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xFFFF0000)))

// Apply to cells
sheet
  .put(cell"A1", "Header")
  .withCellStyle(cell"A1", boldStyle)

// Apply to ranges
sheet.withRangeStyle(range"A1:D1", headerStyle)

// With patches
val styling = (Patch.SetCellStyle(cell"A1", boldStyle): Patch) |+|
              (Patch.SetCellStyle(cell"B1", redStyle): Patch)
sheet.applyPatch(styling)

// Retrieve style
sheet.getCellStyle(cell"A1")  // Returns Option[CellStyle]
```

---

## Streaming API Quick Reference

```scala
// Write single sheet
rows.through(excel.writeStreamTrue(path, "Data", sheetIndex = 1))
  .compile.drain

// Write multiple sheets
excel.writeStreamsSeqTrue(
  path,
  Seq("Sheet1" -> rows1, "Sheet2" -> rows2, "Sheet3" -> rows3)
)

// Read first sheet
excel.readStream(path).compile.toList

// Read by index
excel.readStreamByIndex(path, 2).compile.toList

// Read by name
excel.readSheetStream(path, "Sales").compile.toList
```

---

## Testing Patterns

### Unit Tests
```scala
// Property-based (ScalaCheck)
property("round-trip") {
  forAll { (ref: ARef) =>
    ARef.parse(ref.toA1) == Right(ref)
  }
}

// Integration tests
test("end-to-end workflow") {
  val wb = Workbook("Test")...
  XlsxWriter.write(wb, path)
  val read = XlsxReader.read(path)
  // Verify
}
```

### Manual Testing
```bash
# Generate test file
./mill xl-ooxml.test.runMain com.tjclp.xl.ooxml.ManualStyleTest

# Opens test-styles.xlsx with various formatting
```

---

## Code Style Reminders

1. **Purity**: No side effects in xl-core, xl-ooxml
2. **Totality**: Return Either/Option, never throw
3. **Inline hot paths**: Use `inline` for performance-critical methods
4. **Opaque types**: Zero-overhead wrappers (Column, Row, ARef, Pt, Px, Emu)
5. **Monoid composition**: Enable `|+|` for Patch and StylePatch
6. **Deterministic output**: Sort attributes/elements for stable diffs
7. **Law-based testing**: Property tests for all algebras

---

## Common Pitfalls

### 1. Monoid Syntax Requires Type Ascription
```scala
// ❌ Fails
val p = Patch.Put(...) |+| Patch.SetStyle(...)

// ✅ Works
val p = (Patch.Put(...): Patch) |+| (Patch.SetStyle(...): Patch)
```

### 2. Sheet Constructor Needs All Fields
```scala
// When creating Sheet directly, include styleRegistry
Sheet(name, cells, mergedRanges, colProps, rowProps, None, None, StyleRegistry.default)

// Or use apply method
Sheet("Name").getOrElse(...)  // Handles defaults automatically
```

### 3. Extension Method Naming
Use `@annotation.targetName` to avoid erasure conflicts

---

## Recommended Next Steps

### Option A: Complete P4 Cleanup (~1 day)
1. Implement merged cells XML serialization (~3 hours)
2. Implement column/row properties XML (~4 hours)
3. Update README with full style examples (~1 hour)
4. Update STATUS.md to reflect current state (~30 min)

**Result**: P4 fully complete with all XML parts

---

### Option B: Start P6 Codecs (~2-3 weeks)
**High value**: Eliminates boilerplate for 90% of use cases

**Phase 1: Basic Derivation** (~1 week)
- Type-class for Encoder/Decoder
- Derive for case classes
- Header row binding

**Phase 2: Advanced Features** (~1 week)
- Column mapping (by name or index)
- Optional field handling
- Error accumulation

**Phase 3: Testing & Docs** (~3-5 days)
- Comprehensive tests
- Examples and documentation

**Reference**: docs/plan/09-codecs-and-named-tuples.md

---

### Option C: Performance Optimization (~3-5 days)
1. Profile with JMH
2. Identify allocation hotspots
3. Add more inline annotations
4. Benchmark vs POI
5. Document performance characteristics

**Goal**: Achieve 5-10x memory advantage, 3-5x throughput advantage vs POI

---

### Option D: Quick Wins (~2-3 days each)
- **Hyperlinks**: Add clickable links to cells (~2 days)
- **Comments**: Add cell notes (~2 days)
- **Named Ranges**: Define reusable cell/range names (~1-2 days)
- **Document Properties**: Add metadata (author, title, etc.) (~1 day)

---

## Current Limitations (From LIMITATIONS.md)

### 🔴 High Impact
1. **Update Existing Workbooks**: Cannot modify one sheet while preserving others (P6, 15-20 days)

### 🟡 Medium Impact
1. **Merged Cells Not Serialized**: Easy fix (3 hours) - **Recommend Priority 1**
2. **Column/Row Properties**: Easy fix (4 hours) - **Recommend Priority 1**
3. **Formula Evaluation**: Very complex (60-90 days) - Defer to P11
4. **Theme Colors**: Moderate (1-2 days)
5. **SST in Streaming**: Complex (4-5 days)

### 🟢 Low Impact
- Hyperlinks, Comments, Conditional Formatting, Data Validation, Charts, Drawings, Tables

### 🟣 Security (Critical before 1.0)
- ZIP bomb detection, XXE prevention, Formula injection, Resource limits

---

## File Locations Reference

### Core Implementation
- **Addressing**: `xl-core/src/com/tjclp/xl/addressing.scala`
- **Cell/CellValue**: `xl-core/src/com/tjclp/xl/cell.scala`
- **Sheet/Workbook**: `xl-core/src/com/tjclp/xl/sheet.scala`
- **Patch System**: `xl-core/src/com/tjclp/xl/patch.scala`
- **Styles**: `xl-core/src/com/tjclp/xl/style.scala`
- **Errors**: `xl-core/src/com/tjclp/xl/error.scala`

### OOXML Layer
- **XlsxWriter**: `xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala`
- **XlsxReader**: `xl-ooxml/src/com/tjclp/xl/ooxml/XlsxReader.scala`
- **Worksheet**: `xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala`
- **Styles**: `xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala`
- **SharedStrings**: `xl-ooxml/src/com/tjclp/xl/ooxml/SharedStrings.scala`

### Streaming Layer
- **Excel Algebra**: `xl-cats-effect/src/com/tjclp/xl/io/Excel.scala`
- **ExcelIO**: `xl-cats-effect/src/com/tjclp/xl/io/ExcelIO.scala`
- **StreamingXmlWriter**: `xl-cats-effect/src/com/tjclp/xl/io/StreamingXmlWriter.scala`
- **StreamingXmlReader**: `xl-cats-effect/src/com/tjclp/xl/io/StreamingXmlReader.scala`

---

## Test Patterns

### Property-Based Tests
```scala
import munit.ScalaCheckSuite
import org.scalacheck.Prop.*

class MySpec extends ScalaCheckSuite:
  property("law name") {
    forAll { (input: Type) =>
      // Test invariant
      assert(property holds)
    }
  }
```

### Integration Tests
```scala
import munit.FunSuite
import java.nio.file.{Files, Path}

class IntegrationSpec extends FunSuite:
  val tempDir: FunFixture[Path] = FunFixture[Path](
    setup = _ => Files.createTempDirectory("xl-test-"),
    teardown = dir => Files.walk(dir).sorted(...).forEach(Files.delete)
  )

  tempDir.test("test name") { dir =>
    // Test using temp directory
  }
```

---

## Dependencies

**Current versions**:
- Scala: 3.7.3
- cats-effect: 3.5.7
- fs2-core: 3.11.0
- fs2-data-xml: 1.11.1
- munit: 1.2.0
- scalacheck: 1.18.1

**Build tool**: Mill 0.12.x

---

## Useful Commands

```bash
# Find usage of a symbol
./mill xl-core.compile && grep -r "withCellStyle" xl-core/

# Run single test
./mill xl-core.test.testOnly com.tjclp.xl.StyleRegistrySpec

# Generate test file
./mill xl-ooxml.test.runMain com.tjclp.xl.ooxml.ManualStyleTest

# Check specific file formatting
./mill xl-core.reformat

# Clean and rebuild
./mill clean && ./mill __.compile
```

---

## Performance Benchmarks (Current)

**Streaming Write**:
- 100k rows: 1.1s (~88k rows/sec)
- 30k rows (3 sheets): 0.32s (~94k rows/sec)
- Memory: O(1) constant (~50MB)

**Streaming Read**:
- 100k rows: 1.8s (~55k rows/sec)
- Memory: O(1) constant (~50MB)

**vs Apache POI** (estimated):
- Memory: 16x better (50MB vs 800MB for 100k rows)
- Throughput: Comparable (XL: ~70k rows/sec, POI: ~30-40k rows/sec)

---

## Success Criteria for Next Session

### Must Have
- [ ] All existing tests continue to pass
- [ ] No regressions in functionality
- [ ] Code formatted (./mill __.checkFormat)
- [ ] Documentation updated for changes

### Nice to Have
- [ ] Performance improvements measured
- [ ] New features have comprehensive tests
- [ ] Examples in README/CLAUDE.md

---

## Quick Start Checklist for Next Session

1. Read this file (NEXT_SESSION.md)
2. Read LIMITATIONS.md for current state
3. Run `./mill __.compile` and `./mill __.test` to verify clean state
4. Choose priority from "Recommended Next Steps" section above
5. Reference relevant docs/plan/*.md file for detailed spec
6. Implement with comprehensive tests
7. Update documentation
8. Commit with detailed message

---

## Contact Points

**Project Goal**: Pure functional, mathematically rigorous Excel library for Scala 3

**Non-Negotiables**:
- Purity (no side effects in core)
- Totality (no exceptions)
- Determinism (stable output)
- Law-governed (property-tested)
- Performance (beat POI on memory, match on throughput)

**Philosophy**: Functional programming + spreadsheets = ❤️

---

*This document will be updated at the end of each session to guide the next one.*
