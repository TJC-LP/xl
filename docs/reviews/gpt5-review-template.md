# GPT-5-Pro Review Prompt: XL Query API Design

## Context: About XL

**XL** is a purely functional, mathematically rigorous Excel (OOXML) library for Scala 3.7. The design prioritizes **purity, totality, determinism, and law-governed semantics** with zero-overhead opaque types and compile-time DSLs.

### Core Philosophy (Non-Negotiables)

From `docs/plan/01-purity-charter.md`:

1. **Core purity & effect isolation**
   - `xl-core` and `xl-ooxml` expose **pure functions only**; no side-effects, no clocks, no randomness.
   - `xl-cats-effect` contains **the only interpreters** for: ZIP I/O, file system, streams.
   - All IO return types are **`F[_]`** with typeclass constraints (`Sync`, `Async`) and **no hidden global state**.

2. **Totality & defensive programming**
   - Abolish `null`; use `Option`.
   - No partial functions; exhaustive `match` on enums.
   - All validation is **first class**: return `Either[XLError, A]` or `ValidatedNec[Validation, A]`.

3. **Laws & reasoning**
   - **Lens laws** for optics, **Monoid laws** for patches & style patches, **Action law** for `applyPatch`.
   - **Determinism:** same inputs → same outputs including **byte-level equivalence** after canonicalization.
   - **No surprises:** the `Workbook` model is a finite algebra with explicit defaults; nothing implicit.

### Current Status

- **Complete**: P0-P3 (Bootstrap, Addressing, Core+Patches, Styles)
- **In Progress**: P4 (OOXML MVP) + P5 (Streaming with fs2)
- **77 tests passing**: Property-based tests with ScalaCheck for all core algebras

---

## Current Streaming Implementation

### Excel API (`xl-cats-effect/src/com/tjclp/xl/io/Excel.scala`)

```scala
package com.tjclp.xl.io

import cats.effect.Async
import fs2.Stream
import java.nio.file.Path
import com.tjclp.xl.api.*

/** Row-level streaming data for efficient processing */
case class RowData(
  rowIndex: Int, // 1-based row number
  cells: Map[Int, CellValue] // 0-based column index → value
)

/**
 * Excel algebra for pure functional XLSX operations.
 *
 * Provides both in-memory and streaming APIs:
 *   - read/write: Load entire workbook into memory (good for <10k rows)
 *   - readStream/writeStream: Constant-memory streaming (good for 100k+ rows)
 */
trait Excel[F[_]]:
  def read(path: Path): F[Workbook]
  def write(wb: Workbook, path: Path): F[Unit]

  // Streaming read operations
  def readStream(path: Path): Stream[F, RowData]
  def readSheetStream(path: Path, sheetName: String): Stream[F, RowData]
  def readStreamByIndex(path: Path, sheetIndex: Int): Stream[F, RowData]

  // Streaming write operations
  def writeStream(path: Path, sheetName: String): fs2.Pipe[F, RowData, Unit]
  def writeStream(
    path: Path,
    sheetName: String,
    sheetIndex: Int = 1
  ): fs2.Pipe[F, RowData, Unit]
  def writeStreamsSeq(
    path: Path,
    sheets: Seq[(String, Stream[F, RowData])]
  ): F[Unit]
```

**Key Points**:
- **O(1) constant memory**: Handles 1M+ rows with ~50-100MB memory
- **Performance**: 55k rows/second read, 88k rows/second write
- **Pure functional**: All operations are referentially transparent
- **Already supports fs2 combinators**: `.filter()`, `.map()`, `.fold()` work directly

### Streaming XML Reader Architecture

From `xl-cats-effect/src/com/tjclp/xl/io/StreamingXmlReader.scala` (246 LOC):

```scala
/**
 * True streaming XML reader using fs2-data-xml for constant-memory reads.
 *
 * Parses XML events incrementally without materializing the full document tree.
 */
object StreamingXmlReader:

  /**
   * Stream worksheet rows incrementally from XML byte stream.
   *
   * Memory usage is O(1) regardless of total row count.
   */
  def parseWorksheetStream[F[_]: Sync](
    xmlBytes: Stream[F, Byte],
    sst: Option[SharedStrings]
  ): Stream[F, RowData] =
    xmlBytes
      .through(fs2.text.utf8.decode)
      .through(xml.events[F, String]())
      .through(worksheetEventsToRows(sst))

  /**
   * Convert XML events to RowData using a state machine.
   *
   * Tracks current row index and accumulates cells until </row> event.
   */
  private def worksheetEventsToRows[F[_]](
    sst: Option[SharedStrings]
  ): Pipe[F, XmlEvent, RowData] =
    _.scan(RowBuilder.empty) { (builder, event) =>
      builder.process(event, sst)
    }
      .collect { case RowBuilder.Complete(rowData) =>
        rowData
      }

  // State machine with sealed trait: Idle, InRow, Complete
  private sealed trait RowBuilder:
    def process(event: XmlEvent, sst: Option[SharedStrings]): RowBuilder
```

**Architecture**:
- **fs2-data-xml pull parsing**: Processes XML events incrementally
- **State machine**: `Idle` → `InRow` → `Complete` → `Idle` (next row)
- **No DOM materialization**: Emits `RowData` as each `</row>` tag is encountered
- **Backpressure**: fs2 handles memory backpressure automatically

---

## Testing Patterns

From `xl-core/test/src/com/tjclp/xl/StylePatchSpec.scala` (288 LOC total):

```scala
import munit.ScalaCheckSuite
import org.scalacheck.{Arbitrary, Gen, Prop}
import org.scalacheck.Prop.*
import cats.syntax.all.*

/** Property tests for StylePatch monoid laws */
class StylePatchSpec extends ScalaCheckSuite:

  // Generator for StylePatch
  val genStylePatch: Gen[StylePatch] = Gen.oneOf(
    genFont.map(StylePatch.SetFont.apply),
    genFill.map(StylePatch.SetFill.apply),
    genBorder.map(StylePatch.SetBorder.apply),
    genNumFmt.map(StylePatch.SetNumFmt.apply),
    genAlign.map(StylePatch.SetAlign.apply)
  )

  given Arbitrary[StylePatch] = Arbitrary(genStylePatch)
  import StylePatch.{given, *}

  // ========== Monoid Law Tests ==========

  property("StylePatch Monoid: left identity (empty |+| p == p)") {
    forAll { (patch: StylePatch) =>
      val combined = (StylePatch.empty: StylePatch) |+| (patch: StylePatch)
      val style = CellStyle.default
      val result1 = applyPatch(style, combined)
      val result2 = applyPatch(style, patch)
      assertEquals(result1, result2)
      true
    }
  }

  property("StylePatch Monoid: associativity") {
    forAll { (p1: StylePatch, p2: StylePatch, p3: StylePatch) =>
      val left = ((p1: StylePatch) |+| (p2: StylePatch)) |+| (p3: StylePatch)
      val right = (p1: StylePatch) |+| ((p2: StylePatch) |+| (p3: StylePatch))
      val style = CellStyle.default
      assertEquals(applyPatch(style, left), applyPatch(style, right))
      true
    }
  }

  test("SetFont patch updates font") {
    val style = CellStyle.default
    val newFont = Font("Arial", 14.0, bold = true)
    val patch = StylePatch.SetFont(newFont)
    val result = applyPatch(style, patch)
    assertEquals(result.font, newFont)
    assertEquals(result.fill, style.fill) // Other properties unchanged
  }
```

**Testing Strategy**:
- **Property-based tests**: ScalaCheck for algebraic laws (Monoid identity, associativity)
- **Unit tests**: Specific behavior verification
- **Integration tests**: Large files (100k+ rows) with memory profiling
- **Current**: 77 tests passing (17 addressing + 19 patch + 41 style)
- **Target**: 90%+ coverage with law-based tests

---

## CellValue Type (Core Domain)

From `xl-core/src/com/tjclp/xl/CellValue.scala`:

```scala
enum CellValue derives CanEqual:
  case Empty
  case Text(s: String)
  case Number(n: BigDecimal)
  case Bool(b: Boolean)
  case Formula(expr: String)
  case Error(err: CellError)
  case DateTime(dt: java.time.LocalDateTime)

enum CellError derives CanEqual:
  case Div0, NA, Name, Null, Num, Ref, Value
```

---

## Proposed Query API Design

**Full specification**: `docs/plan/30-query-api.md` (1000 lines)

### Executive Summary

The Query API extends XL's streaming infrastructure with **pure functional, constant-memory query operations** over Excel data using fs2 Pipes.

**Design Principles**:
- **Zero materialization**: Queries process rows incrementally via fs2 streams
- **Pure functional**: All operations are referentially transparent
- **Type-safe**: Leverages CellValue ADT for safe pattern matching
- **Composable**: fs2 Pipes combine via `.through()` and `.andThen()`
- **Columnar focus**: No header assumptions; queries target column indices

**Performance Profile**:
- **Memory**: O(1) constant ~50-100MB
- **Throughput**: 40-50k rows/second with regex matching
- **Scalability**: Handles 1M+ rows

### Module Location

**File**: `xl-cats-effect/src/com/tjclp/xl/io/RowQuery.scala` (~200 LOC)

**Rationale**: Co-locate with streaming infrastructure for shared types and consistent effect handling.

### Core API

```scala
object RowQuery:

  // Regex matching
  def matchingRegex[F[_]](
    regex: Regex,
    columns: Set[Int] = Set.empty  // Empty = all columns
  ): Pipe[F, RowData, RowData]

  // Simple text search
  def containsText[F[_]](
    text: String,
    ignoreCase: Boolean = true
  ): Pipe[F, RowData, RowData]

  // Column-specific predicates
  def whereColumn[F[_]](
    col: Int
  )(
    predicate: CellValue => Boolean
  ): Pipe[F, RowData, RowData]

  // Row-level predicates
  def whereAny[F[_]](
    predicate: CellValue => Boolean
  ): Pipe[F, RowData, RowData]

  def whereAll[F[_]](
    predicate: CellValue => Boolean
  ): Pipe[F, RowData, RowData]

  // Multi-column queries
  case class ColumnQuery(
    col: Int,
    predicate: CellValue => Boolean
  )

  def matchingAll[F[_]](
    queries: List[ColumnQuery]
  ): Pipe[F, RowData, RowData]

  def matchingAny[F[_]](
    queries: List[ColumnQuery]
  ): Pipe[F, RowData, RowData]
```

### Cell Value Stringification (for regex/text matching)

```scala
private def cellToString(cell: CellValue): String = cell match
  case CellValue.Text(s) => s
  case CellValue.Number(n) => n.toString
  case CellValue.Formula(expr) => expr
  case CellValue.DateTime(dt) => dt.toString  // ISO-8601
  case CellValue.Bool(b) => b.toString
  case CellValue.Empty => ""
  case CellValue.Error(_) => ""
```

**Rationale**:
- **Text**: Primary search target
- **Number**: Enables numeric pattern matching (e.g., find 5-digit IDs)
- **Formula**: Search formula expressions (e.g., find `SUM` usages)
- **DateTime**: Date pattern matching (e.g., find dates in 2024)
- **Bool**: Boolean literal search (rare but complete)
- **Empty/Error**: Skip (no meaningful content)

### Usage Examples

```scala
// Example 1: Log analysis
excel.readStream("logs.xlsx")
  .through(RowQuery.matchingRegex("ERROR|FATAL".r, Set(2)))  // Column C
  .through(excel.writeStream("errors.xlsx", "Errors"))
  .compile.drain

// Example 2: Data validation
val emailRegex = """^[\w\.-]+@[\w\.-]+\.\w+$""".r
excel.readStream("customers.xlsx")
  .through(RowQuery.whereColumn(3) {  // Column D: email
    case CellValue.Text(s) => emailRegex.findFirstIn(s).isEmpty
    case _ => true  // Missing = invalid
  })
  .map(row => row.copy(
    cells = row.cells + (10 -> CellValue.Text("INVALID_EMAIL"))
  ))
  .through(excel.writeStream("errors.xlsx", "Errors"))
  .compile.drain

// Example 3: Multi-column ETL
val queries = List(
  ColumnQuery(4, {  // Column E: status
    case CellValue.Text(s) => s.equalsIgnoreCase("COMPLETE")
    case _ => false
  }),
  ColumnQuery(5, {  // Column F: amount
    case CellValue.Number(n) => n > 1000
    case _ => false
  })
)

excel.readStream("orders.xlsx")
  .through(RowQuery.matchingAll(queries))
  .through(excel.writeStream("high_value.xlsx", "Orders"))
  .compile.drain

// Example 4: Aggregation
val totalSales = excel.readStream("sales.xlsx")
  .through(RowQuery.containsText("APPROVED"))
  .map(_.cells.get(7))  // Column H
  .collect { case Some(CellValue.Number(n)) => n }
  .compile.fold(BigDecimal(0))(_ + _)
```

### Implementation Strategy

```scala
def matchingRegex[F[_]](
  regex: Regex,
  columns: Set[Int] = Set.empty
): Pipe[F, RowData, RowData] =
  _.filter { row =>
    val cellsToSearch =
      if columns.isEmpty then row.cells.values
      else row.cells.view.filterKeys(columns.contains).values

    cellsToSearch.exists { cell =>
      val str = cellToString(cell)
      str.nonEmpty && regex.findFirstIn(str).isDefined
    }
  }

def containsText[F[_]](
  text: String,
  ignoreCase: Boolean = true
): Pipe[F, RowData, RowData] =
  val searchText = if ignoreCase then text.toLowerCase else text

  _.filter { row =>
    row.cells.values.exists { cell =>
      val cellStr = cellToString(cell)
      val compareStr = if ignoreCase then cellStr.toLowerCase else cellStr
      compareStr.contains(searchText)
    }
  }

def whereColumn[F[_]](
  col: Int
)(
  predicate: CellValue => Boolean
): Pipe[F, RowData, RowData] =
  _.filter { row =>
    row.cells.get(col).exists(predicate)
  }

def matchingAll[F[_]](
  queries: List[ColumnQuery]
): Pipe[F, RowData, RowData] =
  _.filter { row =>
    queries.forall { query =>
      row.cells.get(query.col).exists(query.predicate)
    }
  }
```

### Testing Strategy

**Unit Tests** (~10-12 tests):
```scala
test("matchingRegex filters matching rows"):
  val rows = Stream(
    RowData(1, Map(0 -> CellValue.Text("ERROR: failed"))),
    RowData(2, Map(0 -> CellValue.Text("INFO: success"))),
    RowData(3, Map(0 -> CellValue.Text("ERROR: timeout")))
  )

  val filtered = rows
    .through(RowQuery.matchingRegex("ERROR".r))
    .compile.toList

  assertEquals(filtered.map(_.rowIndex), List(1, 3))
```

**Property-Based Tests** (~5-7 tests):
```scala
property("matchingRegex preserves streaming"):
  forAll { (rows: List[RowData], regex: Regex) =>
    val stream = Stream.emits(rows)
    val filtered = stream.through(RowQuery.matchingRegex(regex))
    filtered.compile.count.unsafeRunSync() <= rows.size
  }

property("chaining queries is associative"):
  forAll(genRowDataList) { rows =>
    val stream = Stream.emits(rows)
    val q1 = RowQuery.containsText("A")
    val q2 = RowQuery.matchingRegex("B".r)

    val result1 = stream.through(q1).through(q2).compile.toList
    val result2 = stream.through(q1.andThen(q2)).compile.toList

    result1 == result2
  }
```

**Integration Test** (large file):
```scala
test("filter 100k rows with constant memory"):
  val excel = Excel[IO]()

  // Write 100k rows
  val rows = Stream.range(1, 100001).map { i =>
    RowData(i, Map(
      0 -> CellValue.Number(i),
      1 -> CellValue.Text(if i % 10 == 0 then "ERROR" else "INFO")
    ))
  }

  for
    _ <- rows.through(excel.writeStream(testFile, "Data")).compile.drain
    count <- excel.readStream(testFile)
      .through(RowQuery.containsText("ERROR"))
      .compile.count
  yield
    assertEquals(count, 10000L)
```

### Composition Patterns

```scala
// Chain multiple filters (AND logic)
excel.readStream(data)
  .through(RowQuery.containsText("ACTIVE"))
  .through(RowQuery.matchingRegex("""\d{4}""".r, Set(0)))
  .through(RowQuery.whereColumn(2) {
    case CellValue.Number(n) => n > 100
    case _ => false
  })
  .compile.toList

// Filter + Aggregate (streaming)
excel.readStream(sales)
  .through(RowQuery.containsText("COMPLETE"))
  .map(row => row.cells.get(5))
  .collect { case Some(CellValue.Number(n)) => n }
  .compile.fold(BigDecimal(0))(_ + _)

// Filter + Transform + Write (constant memory)
excel.readStream(input)
  .through(RowQuery.matchingRegex("ERROR".r))
  .map(addTimestamp)
  .through(excel.writeStream(output, "Errors"))
  .compile.drain
```

### Future Extensions (Optional)

1. **Column Name Resolution** (~50 LOC):
   ```scala
   case class RowQueryWithHeaders[F[_]](headers: Map[String, Int]):
     def whereColumn(name: String)(pred: CellValue => Boolean): Pipe[F, RowData, RowData]
   ```

2. **Aggregation Helpers** (~100 LOC):
   ```scala
   object RowAggregation:
     def sumColumn[F[_]](col: Int): Pipe[F, RowData, BigDecimal]
     def avgColumn[F[_]](col: Int): Pipe[F, RowData, BigDecimal]
     def countDistinct[F[_]](col: Int): Pipe[F, RowData, Long]
   ```

3. **Query DSL** (~200 LOC):
   ```scala
   import com.tjclp.xl.query.dsl.*
   val query = column(0).matches("ERROR".r) &&
               column(1).isNumber &&
               column(1).gt(100)
   ```

### Known Limitations

1. **No column name awareness**: Queries use 0-based indices, not names
2. **No cross-row context**: Can't reference previous/next rows directly
3. **No aggregation in Pipe**: Must aggregate after filtering
4. **ReDoS risk**: Malicious regex patterns can cause exponential backtracking (documented in scaladoc)

### Security Considerations

**Regex Denial of Service (ReDoS)**:
- **Risk**: Patterns like `(a+)+$` can cause exponential backtracking
- **Mitigation**: Document in API; recommend safe patterns; future timeout support

**Formula Injection**:
- **Not applicable**: `CellValue.Formula` is explicitly typed; queries don't interpret formulas

---

## Review Questions for GPT-5-Pro

Please provide detailed feedback on the following aspects of the Query API design:

### 1. API Design & Ergonomics

- **fs2 Pipes approach**: Is the pure `Pipe[F, RowData, RowData]` design optimal? Would extension methods provide better ergonomics?
- **Method naming**: Are names like `matchingRegex`, `containsText`, `whereColumn` clear and intuitive?
- **Parameter choices**: Are defaults appropriate (e.g., `ignoreCase = true`, `columns = Set.empty`)?
- **Composition**: Are there better patterns for combining queries beyond `.through()`?

### 2. Type Safety & Purity

- **Adherence to philosophy**: Does the design maintain XL's purity guarantees?
- **Type signatures**: Are effect constraints appropriate (`F[_]` with no constraints vs `F[_]: Sync`)?
- **Totality**: Are there any partial functions or unsafe operations?
- **Referential transparency**: Can any operations break RT?

### 3. Performance & Scalability

- **Memory guarantees**: Will the implementation truly maintain O(1) memory?
- **Throughput estimates**: Are 40-50k rows/sec realistic for regex matching?
- **Optimization opportunities**: Any low-hanging fruit for performance improvements?
- **Backpressure**: Does fs2 handle backpressure correctly in these patterns?

### 4. Cell Value Stringification

- **Conversion rules**: Are the rules for converting `CellValue` to `String` sensible?
- **Formula handling**: Should we match formula expressions or computed results?
- **DateTime format**: Is ISO-8601 the right choice for date matching?
- **Error handling**: Should we skip errors or convert to error string like `"#DIV/0!"`?

### 5. Testing Strategy

- **Coverage**: Does the test strategy adequately cover edge cases?
- **Property-based tests**: Are the proposed properties sufficient for law verification?
- **Integration tests**: Should we test with files larger than 100k rows?
- **Memory profiling**: How to reliably test constant memory usage?

### 6. Multi-Column Queries

- **ColumnQuery design**: Is the case class + List approach optimal?
- **AND/OR logic**: Should we provide more sophisticated boolean combinators?
- **Performance**: Will multiple column checks impact throughput significantly?
- **Alternatives**: Are there better patterns for multi-column filtering?

### 7. Implementation Completeness

- **Missing operations**: Are there obvious query types we should include in v1?
- **Error handling**: Should queries return `Either[XLError, ...]` or just filter?
- **Logging/debugging**: Any observability concerns for large-scale filtering?
- **Edge cases**: What happens with empty rows, missing columns, malformed data?

### 8. Future Extensions

- **Column name resolution**: Is the proposed design practical? Should it be in v1?
- **Aggregation helpers**: Should we provide these, or is manual aggregation sufficient?
- **Query DSL**: Is a DSL worth the complexity, or is the Pipe API enough?
- **Parallel processing**: Should we support parallel sheet queries in v1?

### 9. Documentation & Examples

- **API docs**: Are the examples clear and comprehensive?
- **Usage patterns**: Are the example use cases realistic and valuable?
- **Cookbook**: What additional examples would be helpful?
- **Pitfalls**: Are there common mistakes users might make?

### 10. Comparison with Alternatives

- **Comparison with in-memory filtering**: When should users prefer streaming queries?
- **Comparison with external tools**: How does this compare to Excel's built-in filtering or tools like `grep`?
- **Integration with other libraries**: Could we leverage existing libraries (e.g., better regex engines)?

---

## Additional Context

### Project Structure

```
xl/
├── xl-core/          → Pure domain model (Cell, Sheet, Workbook)
├── xl-macros/        → Compile-time validated literals
├── xl-ooxml/         → Pure XML serialization
├── xl-cats-effect/   → IO interpreters (includes streaming)
├── xl-evaluator/     → Formula evaluator [future]
└── xl-testkit/       → Test laws, generators [future]
```

### Dependencies

- **Scala**: 3.7.3
- **Cats Effect**: 3.x
- **fs2**: 3.x
- **fs2-data-xml**: For XML streaming
- **MUnit**: Test framework
- **ScalaCheck**: Property-based testing

### Build Tool

- **Mill**: 0.12.x
- Commands: `./mill __.compile`, `./mill __.test`, `./mill __.reformat`

---

## Request

Please provide:

1. **Overall Assessment**: Is this design sound? Does it fit XL's philosophy?
2. **Critical Issues**: Any deal-breakers or major flaws?
3. **Improvements**: Concrete suggestions for better API design, performance, or testing
4. **Implementation Priorities**: What should be in v1 vs future extensions?
5. **Alternative Approaches**: Are there fundamentally better ways to solve this problem?

**Format**: Please structure your response with clear sections addressing each review question. Include code examples where helpful.

Thank you for your thorough review!
