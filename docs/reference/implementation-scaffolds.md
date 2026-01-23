# Implementation Scaffolds: Idiomatic Scala 3 Code Examples

**Purpose**: Complete, runnable code scaffolds for AI agents and developers
**Last Updated**: 2026-01-23
**Companion To**: [docs/plan/roadmap.md](../plan/roadmap.md)

> **Note**: These are idiomatic Scala 3 scaffolds with enough structure for AI agents to expand. Assumes modules: `xl-core`, `xl-ooxml`, `xl-cats-effect`, `xl-benchmarks`.

---

## Table of Contents

- [A. Core Polish & Streaming](#a-core-polish--streaming)
  - [A.1 Round-trip & Error-path Tests](#a1-round-trip--error-path-tests)
  - [A.2 Two-Phase Streaming Writer](#a2-two-phase-streaming-writer)
  - [A.3 Benchmarks (JMH, POI Comparison)](#a3-benchmarks-jmh-poi-comparison)
- [B. AST Completeness & OOXML Coverage](#b-ast-completeness--ooxml-coverage)
  - [B.1 Comments (Already Complete ✅)](#b1-comments-already-complete-)
  - [B.2 Hyperlinks, Named Ranges, CalcChain](#b2-hyperlinks-named-ranges-calcchain)
  - [B.3 Document Properties & Page Setup](#b3-document-properties--page-setup)
  - [B.4 Conditional Formatting & Data Validation](#b4-conditional-formatting--data-validation)
  - [B.5 Tables & Pivot Tables](#b5-tables--pivot-tables)
  - [B.6 Drawings & Charts](#b6-drawings--charts)
  - [B.7 Part Registry](#b7-part-registry)

---

## A. Core Polish & Streaming

### A.1 Round-trip & Error-path Tests

#### Property-based Round-trip: `Workbook` ↔ `.xlsx`

**Goal**: Write a workbook, read it back, get observational equality.

```scala
// xl-core-test/src/test/scala/com/tjclp/xl/RoundTripSpec.scala
package com.tjclp.xl

import org.scalatest.funsuite.AnyFunSuite
import org.scalatestplus.scalacheck.ScalaCheckPropertyChecks.*
import org.scalacheck.{Arbitrary, Gen}
import cats.syntax.all.*
import java.nio.file.{Files, Path}

// You likely already have these types
final case class Workbook(sheets: Vector[Sheet])
final case class Sheet(name: String, cells: Map[ARef, Cell])
final case class Cell(ref: ARef, value: CellValue /* ... */)

class RoundTripSpec extends AnyFunSuite:

  import Generators.*

  test("Workbook round-trips through XLSX read/write") {
    forAll { (wb: Workbook) =>
      val path = writeTempXlsx(wb)

      val readBack: Either[XLError, Workbook] =
        ExcelIO.read(path)         // or however your API looks

      // You might need weaker equality that ignores style id renumbering
      assert(readBack == Right(normalizeWorkbook(wb)))
    }
  }

  private def writeTempXlsx(wb: Workbook): Path =
    val tmp = Files.createTempFile("xl-roundtrip-", ".xlsx")
    ExcelIO.write(wb, tmp)        // effectful wrapper around XlsxWriter
    tmp

  private def normalizeWorkbook(wb: Workbook): Workbook =
    // Sort sheets, normalize style ids, ranges, etc.
    wb
```

**Generators Module:**

```scala
// xl-core-test/src/test/scala/com/tjclp/xl/Generators.scala
package com.tjclp.xl

import org.scalacheck.{Arbitrary, Gen}
import scala.util.chaining.*

object Generators:

  // Very small sheet for test speed; scale via parameters
  val genARef: Gen[ARef] =
    for
      col <- Gen.chooseNum(0, 10)   // A..K
      row <- Gen.chooseNum(0, 50)   // 1..51
    yield ARef(col, row)

  val genCellValue: Gen[CellValue] =
    Gen.oneOf(
      Gen.const(CellValue.Empty),
      Gen.chooseNum(-1000, 1000).map(CellValue.Number(_)),
      Gen.asciiStr.map(CellValue.Text(_)),
      Gen.oneOf(true, false).map(CellValue.Bool(_))
    )

  val genCell: Gen[Cell] =
    for
      ref   <- genARef
      value <- genCellValue
    yield Cell(ref, value)

  val genSheet: Gen[Sheet] =
    for
      name  <- Gen.alphaStr.suchThat(_.nonEmpty)
      cells <- Gen.listOf(genCell).map { cs =>
                 cs.map(c => c.ref -> c).toMap
               }
    yield Sheet(name.take(31), cells)  // Excel sheet name limit

  val genWorkbook: Gen[Workbook] =
    for
      sheetCount <- Gen.chooseNum(1, 5)
      sheets     <- Gen.listOfN(sheetCount, genSheet)
    yield Workbook(sheets.toVector)

  given Arbitrary[Workbook] =
    Arbitrary(genWorkbook)
```

**Expansion points for AI agents:**
- Add more `CellValue` variants (dates, errors, formulas)
- Tune workbook sizes and distributions
- Add law tests (style registry invariants, merge ranges normalization)

#### Error-path Tests: "Good Errors, Not Crashes"

**Goal**: When there's malformed XML/ZIP/missing parts, return structured `XLError`.

```scala
// xl-ooxml-test/src/test/scala/com/tjclp/xl/XlsxReaderErrorSpec.scala
package com.tjclp.xl

import org.scalatest.funsuite.AnyFunSuite
import java.nio.file.*

class XlsxReaderErrorSpec extends AnyFunSuite:

  test("returns XLError.ZipBombDetected for suspicious compression ratio") {
    val bomb = resourcePath("zip-bombs/bad.xlsx")

    val res = ExcelIO.read(bomb)

    assert(res.left.exists {
      case XLError.ZipBombDetected(path, ratio) => path.endsWith("bad.xlsx")
      case _                                    => false
    })
  }

  test("returns XLError.MissingWorkbookPart for missing xl/workbook.xml") {
    val broken = resourcePath("malformed/missing-workbook.xlsx")

    val res = ExcelIO.read(broken)

    assert(res.left.exists(_.isInstanceOf[XLError.MissingWorkbookPart]))
  }

  private def resourcePath(rel: String): Path =
    Paths.get(getClass.getClassLoader.getResource(rel).toURI)
```

**Expansion points:**
- Add fixtures with missing `[Content_Types].xml`, corrupt sharedStrings
- Tighten `XLError` ADT: `XmlParseError`, `MissingPart`, `SecurityViolation`

---

### A.2 Two-Phase Streaming Writer

**Goal**: First pass scans rows building SST + column stats, second pass writes XLSX.

#### Configuration & Modes

```scala
// xl-ooxml/src/main/scala/com/tjclp/xl/io/StreamingConfig.scala
package com.tjclp.xl.io

enum SharedStringsMode derives CanEqual:
  case None               // inline strings only
  case BuildInMemory      // build SST in memory
  case BuildOnDisk(dir: java.nio.file.Path)

enum ColumnWidthStrategy derives CanEqual:
  case Fixed(width: Double)
  case AutoFitFromSample(sampleRows: Int)
  case None

final case class StreamingWriteConfig(
  sharedStringsMode: SharedStringsMode = SharedStringsMode.BuildInMemory,
  columnWidthStrategy: ColumnWidthStrategy = ColumnWidthStrategy.AutoFitFromSample(200),
  theme: Option[Theme] = None,
  compressionLevel: Int = 6
)
```

#### Abstractions: Rows + Two-Phase Writer

```scala
// xl-core
final case class RowData(
  rowIndex: Int,
  cells: Vector[Cell]  // or (ARef, Cell) depending on your design
)
```

```scala
// xl-cats-effect/src/main/scala/com/tjclp/xl/io/StreamingWorkbookWriter.scala
package com.tjclp.xl.io

import cats.effect.kernel.Concurrent
import fs2.Stream
import java.nio.file.Path

trait StreamingWorkbookWriter[F[_]]:
  def writeSingleSheet(
    rows: Stream[F, RowData],
    dest: Path,
    config: StreamingWriteConfig
  ): F[Unit]

  // multi-sheet variant if needed
```

#### Two-Phase Implementation

```scala
// xl-cats-effect/src/main/scala/com/tjclp/xl/io/TwoPhaseStreamingWorkbookWriter.scala
package com.tjclp.xl.io

import cats.effect.kernel.Concurrent
import cats.syntax.all.*
import fs2.{Stream, Pipe}
import java.nio.file.{Files, Path}

final class TwoPhaseStreamingWorkbookWriter[F[_]: Concurrent](
  ooxmlWriter: OoxmlWriter[F]
) extends StreamingWorkbookWriter[F]:

  def writeSingleSheet(
    rows: Stream[F, RowData],
    dest: Path,
    config: StreamingWriteConfig
  ): F[Unit] =
    config.sharedStringsMode match
      case SharedStringsMode.None =>
        // Single pass, inline strings
        onePassInline(rows, dest, config)

      case _ =>
        // Two-pass: first accumulate SST + column stats, then write
        for
          tmpSheet <- Concurrent[F].blocking(Files.createTempFile("xl-rows-", ".bin"))
          acc      <- firstPassAccumulate(rows, tmpSheet, config)
          _        <- secondPassWrite(tmpSheet, dest, config, acc)
          _        <- Concurrent[F].blocking(Files.deleteIfExists(tmpSheet))
        yield ()

  private final case class Accumulator(
    sharedStrings: SharedStringsBuilder,
    columnStats: ColumnStats
  )

  // Phase 1: Stream through, write compact row representation to temp file
  // while updating SST + stats
  private def firstPassAccumulate(
    rows: Stream[F, RowData],
    tmpSheet: Path,
    config: StreamingWriteConfig
  ): F[Accumulator] =
    val builder = SharedStringsBuilder.empty(config.sharedStringsMode)
    val stats0  = ColumnStats.empty

    rows.evalMap { row =>
      // 1. Update SST and column stats
      val (builder2, stats2) = row.cells.foldLeft(builder -> stats0) {
        case ((b, s), cell) =>
          val (b2, maybeIndex) = b.onCell(cell)  // returns updated builder and maybe SST index
          val s2               = s.onCell(cell)
          (b2, s2)
      }

      // 2. Append row to tmpSheet (e.g. as protobuf/CBOR/your binary format)
      RowBinaryIO.appendRow(tmpSheet, row).as(builder2 -> stats2)
    }
    .fold(Accumulator(builder, stats0)) {
      case (acc, (b, s)) =>
        acc.copy(sharedStrings = b, columnStats = acc.columnStats combine s)
    }
    .compile
    .lastOrError

  // Phase 2: Re-stream rows from tmpSheet and write actual XLSX
  private def secondPassWrite(
    tmpSheet: Path,
    dest: Path,
    config: StreamingWriteConfig,
    acc: Accumulator
  ): F[Unit] =
    val sst     = acc.sharedStrings.result   // produce final SharedStrings part
    val columns = acc.columnStats.toColumnDefs(config.columnWidthStrategy)

    val rows: Stream[F, RowData] = RowBinaryIO.readRows(tmpSheet)

    ooxmlWriter.writeWorkbook(
      dest = dest,
      sharedStrings = Some(sst),
      sheets = List(OneSheetSpec("Sheet1", rows, columns)),
      config = OoxmlWriteConfig.fromStreaming(config)
    )
```

**Expansion points for AI agents:**
- Implement `SharedStringsBuilder` with mode-dependent representation
- Implement `ColumnStats` & `RowBinaryIO` (length measurements + binary row dump/load)
- Implement `OoxmlWriter.writeWorkbook` to accept streaming sheet input

---

### A.3 Benchmarks (JMH, POI Comparison)

```scala
// xl-benchmarks/src/jmh/scala/com/tjclp/xl/WriteBenchmark.scala
package com.tjclp.xl.benchmarks

import org.openjdk.jmh.annotations.*
import java.util.concurrent.TimeUnit
import java.nio.file.*
import com.tjclp.xl.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream

@State(Scope.Benchmark)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@BenchmarkMode(Array(Mode.AverageTime))
class WriteBenchmark:

  @Param(Array("1000", "10000", "100000"))
  var rows: Int = _

  var workbookXL: Workbook = _
  var workbookPOI: XSSFWorkbook = _

  @Setup(Level.Trial)
  def setup(): Unit =
    workbookXL  = buildTestWorkbook(rows)
    workbookPOI = buildPoiWorkbook(rows)

  @Benchmark
  def xl_write(): Path =
    val path = Files.createTempFile("xl-bench-", ".xlsx")
    ExcelIO.write(workbookXL, path)
    path

  @Benchmark
  def poi_write(): Path =
    val path = Files.createTempFile("poi-bench-", ".xlsx")
    val out  = new FileOutputStream(path.toFile)
    try workbookPOI.write(out)
    finally out.close()
    path

  private def buildTestWorkbook(n: Int): Workbook =
    val cells = (0 until n).map { i =>
      val ref = ARef(col = 0, row = i)
      Cell(ref, CellValue.Number(i.toDouble))
    }.toVector
    val sheet = Sheet("bench", cells.map(c => c.ref -> c).toMap)
    Workbook(Vector(sheet))

  private def buildPoiWorkbook(n: Int): XSSFWorkbook =
    val wb    = new XSSFWorkbook()
    val sheet = wb.createSheet("bench")
    var i     = 0
    while i < n do
      val row  = sheet.createRow(i)
      val cell = row.createCell(0)
      cell.setCellValue(i.toDouble)
      i += 1
    wb
```

**Expansion points:**
- Add read benchmarks and streaming writer benchmarks
- Add parameters for columns, styles, formulas
- Plug into CI to fail on regressions beyond threshold

---

## B. AST Completeness & OOXML Coverage

**Pattern**: Each OOXML part gets strongly-typed AST with:
- One `XmlCodec[A]` for conversion
- Optional `otherXml: Seq[Elem]`/`unknownAttributes` for forward-compat

**Helper trait:**

```scala
// xl-ooxml/src/main/scala/com/tjclp/xl/ooxml/XmlCodec.scala
package com.tjclp.xl.ooxml

import scala.xml.Elem
import com.tjclp.xl.XLError

trait XmlCodec[A]:
  def encode(a: A): Elem
  def decode(e: Elem): Either[XLError, A]
```

---

### B.1 Comments (Already Complete ✅)

**Status**: ✅ Fully implemented in codebase
- AST: `Comment.scala` in xl-core
- OOXML: `OoxmlComment`, `OoxmlComments` classes
- Tests: `CommentsSpec.scala` (12+ tests)
- Round-trip: xl/commentsN.xml + VML drawings

**No scaffolding needed** - reference existing implementation.

---

### B.2 Hyperlinks, Named Ranges, CalcChain

#### Hyperlinks in Worksheet

```scala
// xl-ooxml/src/main/scala/com/tjclp/xl/ooxml/hyperlinks.scala
package com.tjclp.xl.ooxml

import scala.xml.Elem
import com.tjclp.xl.{ARef, XLError}

final case class OoxmlHyperlink(
  ref: ARef,
  location: Option[String],         // internal (e.g. "Sheet2!A1")
  display: Option[String],
  tooltip: Option[String],
  relId: Option[String],            // external link via relationships
  otherAttrs: Map[String, String] = Map.empty
)

final case class OoxmlHyperlinks(
  hyperlinks: Vector[OoxmlHyperlink]
)

object OoxmlHyperlinks:

  given XmlCodec[OoxmlHyperlinks] with
    def decode(e: Elem): Either[XLError, OoxmlHyperlinks] =
      if e.label != "hyperlinks" then Left(XLError.XmlStructure("Expected <hyperlinks>", e.position))
      else
        val hs = (e \ "hyperlink").collect { case el: Elem => el }.toVector
        hs.traverse(decodeHyperlink).map(OoxmlHyperlinks(_))

    private def decodeHyperlink(e: Elem): Either[XLError, OoxmlHyperlink] =
      for
        refStr <- e.attribute("ref").map(_.text).toRight(XLError.XmlStructure("Missing @ref", e.position))
        ref    <- ARef.parse(refStr).left.map(err => XLError.InvalidRef(refStr, err.message))
      yield
        val known = Set("ref", "location", "display", "tooltip", "r:id")
        val attrs = e.attributes.asAttrMap.filterNot { case (k, _) => known.contains(k) }
        OoxmlHyperlink(
          ref      = ref,
          location = e.attribute("location").map(_.text),
          display  = e.attribute("display").map(_.text),
          tooltip  = e.attribute("tooltip").map(_.text),
          relId    = e.attribute("id", "r").map(_.text),
          otherAttrs = attrs
        )

    def encode(hs: OoxmlHyperlinks): Elem =
      <hyperlinks>
        {hs.hyperlinks.map(encodeHyperlink)}
      </hyperlinks>

    private def encodeHyperlink(h: OoxmlHyperlink): Elem =
      var base = <hyperlink ref={ARef.show(h.ref)}></hyperlink>

      def setAttr(name: String, value: String): Elem =
        base % new scala.xml.UnprefixedAttribute(name, value, base.attributes)

      base = h.location.fold(base)(loc => setAttr("location", loc))
      base = h.display.fold(base)(d => setAttr("display", d))
      base = h.tooltip.fold(base)(t => setAttr("tooltip", t))
      base = h.relId.fold(base)(id =>
        base % new scala.xml.PrefixedAttribute("r", "id", id, base.attributes)
      )

      h.otherAttrs.foldLeft(base) { case (el, (k, v)) =>
        el % new scala.xml.UnprefixedAttribute(k, v, el.attributes)
      }
```

#### Named Ranges

```scala
// xl-ooxml/src/main/scala/com/tjclp/xl/ooxml/names.scala
package com.tjclp.xl.ooxml

import scala.xml.Elem
import com.tjclp.xl.XLError

final case class OoxmlDefinedName(
  name: String,
  localSheetId: Option[Int],
  hidden: Boolean,
  comment: Option[String],
  refersTo: String,                 // formula-like text, e.g. "Sheet1!$A$1:$B$10"
  otherAttrs: Map[String, String] = Map.empty
)

final case class OoxmlDefinedNames(
  names: Vector[OoxmlDefinedName]
)

object OoxmlDefinedNames:

  given XmlCodec[OoxmlDefinedNames] with
    def decode(e: Elem): Either[XLError, OoxmlDefinedNames] =
      if e.label != "definedNames" then Left(XLError.XmlStructure("Expected <definedNames>", e.position))
      else
        val ns = (e \ "definedName").collect { case el: Elem => el }.toVector
        ns.traverse(decodeDefinedName).map(OoxmlDefinedNames(_))

    private def decodeDefinedName(e: Elem): Either[XLError, OoxmlDefinedName] =
      for
        name <- e.attribute("name").map(_.text).toRight(XLError.XmlStructure("Missing @name", e.position))
        text  = e.text.trim
      yield
        val localSheetId = e.attribute("localSheetId").map(_.text.toInt)
        val hidden       = e.attribute("hidden").exists(a => a.text == "1" || a.text == "true")
        val comment      = e.attribute("comment").map(_.text)
        val known        = Set("name", "localSheetId", "hidden", "comment")
        val attrs        = e.attributes.asAttrMap.filterNot { case (k, _) => known.contains(k) }

        OoxmlDefinedName(
          name       = name,
          localSheetId = localSheetId,
          hidden     = hidden,
          comment    = comment,
          refersTo   = text,
          otherAttrs = attrs
        )

    def encode(ds: OoxmlDefinedNames): Elem =
      <definedNames>
        {ds.names.map(encodeDefinedName)}
      </definedNames>

    private def encodeDefinedName(d: OoxmlDefinedName): Elem =
      var el = <definedName name={d.name}>{d.refersTo}</definedName>

      d.localSheetId.foreach { id =>
        el = el % new scala.xml.UnprefixedAttribute("localSheetId", id.toString, el.attributes)
      }
      if d.hidden then
        el = el % new scala.xml.UnprefixedAttribute("hidden", "1", el.attributes)
      d.comment.foreach { c =>
        el = el % new scala.xml.UnprefixedAttribute("comment", c, el.attributes)
      }

      d.otherAttrs.foldLeft(el) { case (e, (k, v)) =>
        e % new scala.xml.UnprefixedAttribute(k, v, e.attributes)
      }
```

#### CalcChain

```scala
// xl-ooxml/src/main/scala/com/tjclp/xl/ooxml/calcchain.scala
package com.tjclp.xl.ooxml

import scala.xml.Elem
import com.tjclp.xl.{ARef, XLError}

final case class CalcCell(
  ref: ARef,
  sheetId: Option[Int],
  inChain: Boolean,
  newThread: Boolean,
  otherAttrs: Map[String, String] = Map.empty
)

final case class CalcChain(
  cells: Vector[CalcCell]
)

object CalcChain:

  given XmlCodec[CalcChain] with
    def decode(e: Elem): Either[XLError, CalcChain] =
      if e.label != "calcChain" then Left(XLError.XmlStructure("Expected <calcChain>", e.position))
      else
        val cs = (e \ "c").collect { case el: Elem => el }.toVector
        cs.traverse(decodeCell).map(CalcChain(_))

    private def decodeCell(e: Elem): Either[XLError, CalcCell] =
      for
        refStr <- e.attribute("r").map(_.text).toRight(XLError.XmlStructure("Missing @r", e.position))
        ref    <- ARef.parse(refStr).left.map(err => XLError.InvalidRef(refStr, err.message))
      yield
        val sheetId  = e.attribute("i").map(_.text.toInt)
        val inChain  = e.attribute("s").exists(a => a.text == "1")
        val newThread= e.attribute("t").exists(a => a.text == "1")
        val known    = Set("r", "i", "s", "t")
        val attrs    = e.attributes.asAttrMap.filterNot { case (k, _) => known.contains(k) }

        CalcCell(ref, sheetId, inChain, newThread, attrs)

    def encode(cc: CalcChain): Elem =
      <calcChain>
        {cc.cells.map(encodeCell)}
      </calcChain>

    private def encodeCell(c: CalcCell): Elem =
      var el = <c r={ARef.show(c.ref)}></c>
      c.sheetId.foreach { id =>
        el = el % new scala.xml.UnprefixedAttribute("i", id.toString, el.attributes)
      }
      if c.inChain then
        el = el % new scala.xml.UnprefixedAttribute("s", "1", el.attributes)
      if c.newThread then
        el = el % new scala.xml.UnprefixedAttribute("t", "1", el.attributes)

      c.otherAttrs.foldLeft(el) { case (e, (k, v)) =>
        e % new scala.xml.UnprefixedAttribute(k, v, e.attributes)
      }
```

---

### B.3 Document Properties & Page Setup

#### Document Properties

```scala
// xl-ooxml/src/main/scala/com/tjclp/xl/ooxml/docprops.scala
package com.tjclp.xl.ooxml

import scala.xml.Elem
import com.tjclp.xl.XLError

final case class CoreProperties(
  title: Option[String],
  subject: Option[String],
  creator: Option[String],
  description: Option[String],
  lastModifiedBy: Option[String],
  revision: Option[String],
  created: Option[java.time.Instant],
  modified: Option[java.time.Instant],
  otherChildren: Seq[Elem] = Seq.empty
)

object CoreProperties:

  given XmlCodec[CoreProperties] with
    def decode(e: Elem): Either[XLError, CoreProperties] =
      if e.label != "coreProperties" then Left(XLError.XmlStructure("Expected <coreProperties>", e.position))
      else
        def find(label: String) = (e \ label).headOption.map(_.text.trim).filter(_.nonEmpty)

        Right(
          CoreProperties(
            title          = find("title"),
            subject        = find("subject"),
            creator        = find("creator"),
            description    = find("description"),
            lastModifiedBy = find("lastModifiedBy"),
            revision       = find("revision"),
            created        = find("created").flatMap(parseInstant),
            modified       = find("modified").flatMap(parseInstant),
            otherChildren  = e.child.collect {
              case el: Elem if !Set("title", "subject", "creator", "description",
                                   "lastModifiedBy", "revision", "created", "modified").contains(el.label) => el
            }.toVector
          )
        )

    private def parseInstant(s: String): Option[java.time.Instant] =
      scala.util.Try(java.time.Instant.parse(s)).toOption

    def encode(p: CoreProperties): Elem =
      <coreProperties>
        {p.title.map(t => <title>{t}</title>).toSeq}
        {p.subject.map(s => <subject>{s}</subject>).toSeq}
        {p.creator.map(c => <creator>{c}</creator>).toSeq}
        {p.description.map(d => <description>{d}</description>).toSeq}
        {p.lastModifiedBy.map(m => <lastModifiedBy>{m}</lastModifiedBy>).toSeq}
        {p.revision.map(r => <revision>{r}</revision>).toSeq}
        {p.created.map(d => <created>{d.toString}</created>).toSeq}
        {p.modified.map(d => <modified>{d.toString}</modified>).toSeq}
        {p.otherChildren}
      </coreProperties>
```

#### Page Setup

```scala
// xl-ooxml/src/main/scala/com/tjclp/xl/ooxml/pagesetup.scala
package com.tjclp.xl.ooxml

import scala.xml.Elem
import com.tjclp.xl.XLError

final case class PageMargins(
  left: Double,
  right: Double,
  top: Double,
  bottom: Double,
  header: Double,
  footer: Double
)

final case class PageSetup(
  orientation: Option[String],
  paperSize: Option[Int],
  scale: Option[Int],
  fitToWidth: Option[Int],
  fitToHeight: Option[Int],
  otherAttrs: Map[String, String] = Map.empty
)

object PageSetupXml:

  def decodeMargins(e: Elem): Either[XLError, PageMargins] =
    def d(attr: String) =
      e.attribute(attr).map(_.text.toDouble).toRight(XLError.XmlStructure(s"Missing @$attr", e.position))

    for
      left   <- d("left")
      right  <- d("right")
      top    <- d("top")
      bottom <- d("bottom")
      header <- d("header")
      footer <- d("footer")
    yield PageMargins(left, right, top, bottom, header, footer)

  def encodeMargins(m: PageMargins): Elem =
    <pageMargins
      left={m.left.toString}
      right={m.right.toString}
      top={m.top.toString}
      bottom={m.bottom.toString}
      header={m.header.toString}
      footer={m.footer.toString}
    />

  def decodeSetup(e: Elem): Either[XLError, PageSetup] =
    val known = Set("orientation", "paperSize", "scale", "fitToWidth", "fitToHeight")
    val attrs = e.attributes.asAttrMap

    Right(
      PageSetup(
        orientation = attrs.get("orientation"),
        paperSize   = attrs.get("paperSize").map(_.toInt),
        scale       = attrs.get("scale").map(_.toInt),
        fitToWidth  = attrs.get("fitToWidth").map(_.toInt),
        fitToHeight = attrs.get("fitToHeight").map(_.toInt),
        otherAttrs  = attrs.filterNot { case (k, _) => known.contains(k) }
      )
    )

  def encodeSetup(s: PageSetup): Elem =
    var el = <pageSetup></pageSetup>
    def set(name: String, value: String): Elem =
      el % new scala.xml.UnprefixedAttribute(name, value, el.attributes)

    el = s.orientation.fold(el)(o => set("orientation", o))
    el = s.paperSize.fold(el)(p => set("paperSize", p.toString))
    el = s.scale.fold(el)(sc => set("scale", sc.toString))
    el = s.fitToWidth.fold(el)(w => set("fitToWidth", w.toString))
    el = s.fitToHeight.fold(el)(h => set("fitToHeight", h.toString))

    s.otherAttrs.foldLeft(el) { case (e, (k, v)) =>
      e % new scala.xml.UnprefixedAttribute(k, v, e.attributes)
    }
```

---

### B.4 Conditional Formatting & Data Validation

#### Conditional Formatting

```scala
// xl-ooxml/src/main/scala/com/tjclp/xl/ooxml/conditional.scala
package com.tjclp.xl.ooxml

import scala.xml.Elem
import com.tjclp.xl.{CellRange, XLError}

enum ConditionType derives CanEqual:
  case CellIs, Expression, ColorScale, DataBar, IconSet, Other(raw: String)

final case class CfRule(
  ruleType: ConditionType,
  priority: Int,
  stopIfTrue: Boolean,
  formula: List[String],
  dxfId: Option[Int],
  otherAttrs: Map[String, String] = Map.empty,
  otherChildren: Seq[Elem] = Seq.empty
)

final case class ConditionalFormatting(
  ranges: List[CellRange],
  rules: List[CfRule]
)

object ConditionalFormattingXml:

  def decode(e: Elem): Either[XLError, ConditionalFormatting] =
    for
      sqref <- e.attribute("sqref").map(_.text).toRight(XLError.XmlStructure("Missing @sqref", e.position))
      ranges <- CellRange.parseSqref(sqref)
      rules <- (e \ "cfRule").collect { case el: Elem => el }.toList.traverse(decodeRule)
    yield ConditionalFormatting(ranges, rules)

  private def decodeRule(e: Elem): Either[XLError, CfRule] =
    val rawType  = e.attribute("type").map(_.text)
    val ct       = rawType.map {
                     case "cellIs"      => ConditionType.CellIs
                     case "expression"  => ConditionType.Expression
                     case "colorScale"  => ConditionType.ColorScale
                     case "dataBar"     => ConditionType.DataBar
                     case "iconSet"     => ConditionType.IconSet
                     case other         => ConditionType.Other(other)
                   }.getOrElse(ConditionType.Other(""))

    val priority = e.attribute("priority").map(_.text.toInt).getOrElse(0)
    val stop     = e.attribute("stopIfTrue").exists(a => a.text == "1" || a.text == "true")
    val dxfId    = e.attribute("dxfId").map(_.text.toInt)
    val formula  = (e \ "formula").map(_.text).toList

    val known = Set("type", "priority", "stopIfTrue", "dxfId")
    val attrs = e.attributes.asAttrMap.filterNot { case (k, _) => known.contains(k) }
    val kids  = e.child.collect { case el: Elem if el.label != "formula" => el }.toVector

    Right(CfRule(ct, priority, stop, formula, dxfId, attrs, kids))

  def encode(cf: ConditionalFormatting): Elem =
    <conditionalFormatting sqref={cf.ranges.map(_.toA1).mkString(" ")}>
      {cf.rules.map(encodeRule)}
    </conditionalFormatting>

  private def encodeRule(r: CfRule): Elem =
    val typeStr = r.ruleType match
      case ConditionType.CellIs     => "cellIs"
      case ConditionType.Expression => "expression"
      case ConditionType.ColorScale => "colorScale"
      case ConditionType.DataBar    => "dataBar"
      case ConditionType.IconSet    => "iconSet"
      case ConditionType.Other(raw) => raw

    var el = <cfRule type={typeStr} priority={r.priority.toString}>
      {r.formula.map(f => <formula>{f}</formula>)}
      {r.otherChildren}
    </cfRule>

    if r.stopIfTrue then
      el = el % new scala.xml.UnprefixedAttribute("stopIfTrue", "1", el.attributes)
    r.dxfId.foreach { id =>
      el = el % new scala.xml.UnprefixedAttribute("dxfId", id.toString, el.attributes)
    }

    r.otherAttrs.foldLeft(el) { case (e, (k, v)) =>
      e % new scala.xml.UnprefixedAttribute(k, v, e.attributes)
    }
```

**Note**: Data validation follows similar pattern with `<dataValidations>` / `<dataValidation>` elements.

---

### B.5 Tables & Pivot Tables

#### Tables

```scala
// xl-ooxml/src/main/scala/com/tjclp/xl/ooxml/tables.scala
package com.tjclp.xl.ooxml

import scala.xml.Elem
import com.tjclp.xl.{CellRange, XLError}

final case class TableColumn(
  id: Long,
  name: String
)

final case class TableAutoFilter(
  range: CellRange
)

final case class OoxmlTable(
  id: Long,
  name: String,
  displayName: String,
  ref: CellRange,
  headerRowCount: Int,
  totalsRowCount: Int,
  columns: Vector[TableColumn],
  autoFilter: Option[TableAutoFilter],
  otherChildren: Seq[Elem] = Seq.empty,
  otherAttrs: Map[String, String] = Map.empty
)

object OoxmlTable:

  given XmlCodec[OoxmlTable] with
    def decode(e: Elem): Either[XLError, OoxmlTable] =
      for
        idStr   <- e.attribute("id").map(_.text).toRight(XLError.XmlStructure("Missing @id", e.position))
        id      <- Right(idStr.toLong)
        name    <- e.attribute("name").map(_.text).toRight(XLError.XmlStructure("Missing @name", e.position))
        disp    <- e.attribute("displayName").map(_.text).toRight(XLError.XmlStructure("Missing @displayName", e.position))
        refStr  <- e.attribute("ref").map(_.text).toRight(XLError.XmlStructure("Missing @ref", e.position))
        range   <- CellRange.parse(refStr).left.map(err => XLError.InvalidRange(refStr, err.message))
      yield
        val header = e.attribute("headerRowCount").map(_.text.toInt).getOrElse(1)
        val totals = e.attribute("totalsRowCount").map(_.text.toInt).getOrElse(0)
        val cols   = (e \ "tableColumns" \ "tableColumn").map { col =>
          TableColumn(
            id   = col.attribute("id").map(_.text.toLong).getOrElse(0L),
            name = col.attribute("name").map(_.text).getOrElse("")
          )
        }.toVector

        val afOpt = (e \ "autoFilter").headOption.collect { case el: Elem => el }.flatMap { af =>
          CellRange.parse(af.attribute("ref").map(_.text).getOrElse("")).toOption.map(TableAutoFilter(_))
        }

        val knownAttrs = Set("id", "name", "displayName", "ref", "headerRowCount", "totalsRowCount")
        val attrs      = e.attributes.asAttrMap.filterNot { case (k, _) => knownAttrs.contains(k) }
        val knownKids  = (e \ "tableColumns") ++ (e \ "autoFilter")
        val others     = e.child.collect { case el: Elem if !knownKids.contains(el) => el }.toVector

        OoxmlTable(id, name, disp, range, header, totals, cols, afOpt, others, attrs)

    def encode(t: OoxmlTable): Elem =
      <table
        id={t.id.toString}
        name={t.name}
        displayName={t.displayName}
        ref={t.ref.toA1}
        headerRowCount={t.headerRowCount.toString}
        totalsRowCount={t.totalsRowCount.toString}>
        <tableColumns count={t.columns.size.toString}>
          {t.columns.map(c => <tableColumn id={c.id.toString} name={c.name}/>)}
        </tableColumns>
        {t.autoFilter.map(af => <autoFilter ref={af.range.toA1}/>).toSeq}
        {t.otherChildren}
      </table>
```

**Note**: Pivot tables/pivot cache follow similar ADT patterns (verbose but mechanical).

---

### B.6 Drawings & Charts

#### Drawings

```scala
// xl-ooxml/src/main/scala/com/tjclp/xl/ooxml/drawings.scala
package com.tjclp.xl.ooxml

import scala.xml.Elem
import com.tjclp.xl.XLError

enum Anchor derives CanEqual:
  case TwoCell(fromCol: Int, fromRow: Int, toCol: Int, toRow: Int)
  case OneCell(col: Int, row: Int)
  case Absolute(x: Int, y: Int)

final case class Picture(
  relId: String,
  name: Option[String],
  descr: Option[String],
  anchor: Anchor
)

final case class OoxmlDrawing(
  pictures: Vector[Picture],
  otherChildren: Seq[Elem] = Seq.empty
)

object OoxmlDrawing:

  given XmlCodec[OoxmlDrawing] with
    def decode(e: Elem): Either[XLError, OoxmlDrawing] =
      // Parse <xdr:twoCellAnchor>, <xdr:pic>, etc.
      Right(OoxmlDrawing(Vector.empty))  // TODO: fill in

    def encode(d: OoxmlDrawing): Elem =
      // Write <xdr:wsDr>...
      <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
        {d.pictures.map(encodePicture)}
        {d.otherChildren}
      </xdr:wsDr>

    private def encodePicture(p: Picture): Elem =
      // TODO: full implementation
      <xdr:twoCellAnchor/>
```

#### Charts

```scala
// xl-ooxml/src/main/scala/com/tjclp/xl/ooxml/charts.scala
package com.tjclp.xl.ooxml

import scala.xml.Elem
import com.tjclp.xl.{CellRange, XLError}

enum ChartType derives CanEqual:
  case Bar, Line, Pie, Scatter, Other(raw: String)

final case class Series(
  nameFormula: Option[String],
  valuesFormula: String,
  categoryFormula: Option[String]
)

final case class OoxmlChart(
  chartType: ChartType,
  series: Vector[Series],
  title: Option[String],
  otherChildren: Seq[Elem] = Seq.empty
)

final case class OoxmlChartPart(
  charts: Vector[OoxmlChart],
  otherChildren: Seq[Elem] = Seq.empty
)

object OoxmlChartPart:

  given XmlCodec[OoxmlChartPart] with
    def decode(e: Elem): Either[XLError, OoxmlChartPart] =
      // Parse <c:chartSpace>/<c:chart>...
      Right(OoxmlChartPart(Vector.empty))  // TODO

    def encode(c: OoxmlChartPart): Elem =
      // Write <c:chartSpace>...
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        {c.otherChildren}
      </c:chartSpace>
```

**Expansion points**: Flesh out OOXML details (namespaces, element structure), wire to RelationshipGraph.

---

### B.7 Part Registry

```scala
// xl-ooxml/src/main/scala/com/tjclp/xl/ooxml/PartRegistry.scala
package com.tjclp.xl.ooxml

import scala.xml.Elem
import com.tjclp.xl.XLError

enum WorkbookPartKind derives CanEqual:
  case WorkbookXml
  case WorksheetXml(index: Int)
  case Styles
  case SharedStrings
  case Comments(sheetIndex: Int)
  case Hyperlinks(sheetIndex: Int)
  case Table(sheetIndex: Int, tableId: Long)
  case PivotTable(sheetIndex: Int)
  case Drawing(sheetIndex: Int)
  case Chart
  case CoreProps
  case AppProps
  case CalcChain
  case Custom(name: String)

final case class WorkbookPart[A](
  kind: WorkbookPartKind,
  path: String,
  contentType: String,
  codec: XmlCodec[A]
)

object PartRegistry:

  // Central registry for all known parts
  val comments: WorkbookPart[OoxmlComments] =
    WorkbookPart(
      kind         = WorkbookPartKind.Comments(sheetIndex = -1),
      path         = "xl/comments1.xml",
      contentType  = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
      codec        = summon[XmlCodec[OoxmlComments]]
    )

  // ... styles, shared strings, tables, etc.

  def lookup(path: String, contentType: String): Option[WorkbookPart[?]] =
    // Build map or pattern match on path prefixes
    ???
```

**Writer usage**: Use registry to decide which AST to parse/emit for each part, and which unknown parts to copy verbatim for surgical preservation.

---

## Usage with AI Agents

### For A (Core + Streaming):
1. Implement `SharedStringsBuilder`, `ColumnStats`, `RowBinaryIO` around two-phase writer skeleton
2. Add round-trip + error-path tests, extend to more cases
3. Plug JMH benchmark into build and wire in real `Workbook`/`ExcelIO`

### For B (AST Coverage):
1. Lift patterns for comments, hyperlinks, named ranges into real package structure
2. Extend ASTs to fully cover OOXML spec where it matters (use `otherAttrs`/`otherChildren` as escape hatches)
3. Build `PartRegistry` and hook into existing `XlsxReader`/`XlsxWriter`

---

**Last Updated**: 2026-01-23
**Maintained By**: XL Core Team
**For Strategic Context**: See [docs/plan/roadmap.md](../plan/roadmap.md)
