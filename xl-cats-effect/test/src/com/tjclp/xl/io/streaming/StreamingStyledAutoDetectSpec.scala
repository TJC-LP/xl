package com.tjclp.xl.io.streaming

import java.io.ByteArrayOutputStream
import java.nio.file.{Files, Path}
import java.time.LocalDate
import java.util.zip.{ZipEntry, ZipFile, ZipInputStream, ZipOutputStream}

import scala.jdk.CollectionConverters.*

import cats.effect.IO
import cats.effect.unsafe.implicits.global
import munit.{CatsEffectSuite, ScalaCheckSuite}
import org.scalacheck.{Gen, Prop}

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.{BoundsAccumulator, ExcelIO, RowData, StreamingXmlWriter, StyledRowData}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.{OoxmlStyles, StyleIndex, XlsxReader, XmlSecurity, XmlUtil}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.sheets.styleSyntax.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.workbooks.Workbook

/** Shared fixtures for the GH-675 styled two-pass writer suites. */
private object StyledAutoDetect:

  val excel: ExcelIO[IO] = ExcelIO.instance[IO]

  val date: CellStyle = CellStyle.default.withNumFmt(NumFmt.Date)
  val bold: CellStyle = CellStyle.default.copy(font = Font.default.copy(bold = true))
  val percent: CellStyle = CellStyle.default.withNumFmt(NumFmt.Percent)
  val custom: CellStyle = CellStyle.default.withNumFmt(NumFmt.Custom("0.000"))

  def tempXlsx(label: String): Path =
    val p = Files.createTempFile(s"xl-styled-autodetect-$label-", ".xlsx")
    p.toFile.deleteOnExit()
    p

  def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), "UTF-8")
        case None => throw new AssertionError(s"zip entry $name not found in $path")
    finally zip.close()

  def entryTimes(path: Path): Vector[(String, Long)] =
    val zip = new ZipFile(path.toFile)
    try zip.entries().asScala.map(e => e.getName -> e.getTime).toVector
    finally zip.close()

  /** What a `setTime(0L)` entry reads back as (the DOS epoch in the local zone). */
  val epochEntryTime: Long =
    val bytes = ByteArrayOutputStream()
    val zos = ZipOutputStream(bytes)
    val entry = ZipEntry("x")
    entry.setTime(0L)
    zos.putNextEntry(entry)
    zos.closeEntry()
    zos.close()
    val zis = ZipInputStream(java.io.ByteArrayInputStream(bytes.toByteArray))
    try Option(zis.getNextEntry).map(_.getTime).getOrElse(-1L)
    finally zis.close()

  /** Each emitted `<c>`'s `r` → its `s` attribute (None when absent), in document order. */
  def styleAttrs(path: Path): Vector[(String, Option[String])] =
    val xml = entryText(path, "xl/worksheets/sheet1.xml")
    val root = XmlSecurity
      .parseSafe(xml, "sheet1.xml")
      .fold(e => throw new AssertionError(e.message), identity)
    (root \\ "c").map(c => (c \@ "r") -> Option(c \@ "s").filter(_.nonEmpty)).toVector

  def dateCell(d: LocalDate): CellValue = CellValue.DateTime(d.atStartOfDay())

  def run(rows: List[StyledRowData], pipe: fs2.Pipe[IO, StyledRowData, Unit]): IO[Unit] =
    fs2.Stream.emits(rows).covary[IO].through(pipe).compile.drain

/**
 * GH-675: `writeStreamStyledWithAutoDetect` — the styled sibling of the two-pass writer
 * (`writeStreamWithAutoDetect`). Same style-table semantics as `writeStreamStyled` (the default is
 * cellXf 0, duplicates collapse by canonical key, row indices remap before emission, an index
 * outside the table falls back to the default), plus the computed `<dimension>`.
 */
class StreamingStyledAutoDetectSpec extends CatsEffectSuite:
  import StyledAutoDetect.*

  private val ledger = List(
    StyledRowData(1, Map(0 -> CellValue.Text("name"), 1 -> CellValue.Text("when"))),
    StyledRowData(
      2,
      Map(0 -> CellValue.Text("Alpha"), 1 -> dateCell(LocalDate.of(2026, 1, 15))),
      Map(1 -> 0)
    ),
    StyledRowData(
      3,
      Map(
        0 -> CellValue.Text("Beta"),
        1 -> dateCell(LocalDate.of(2026, 2, 1)),
        2 -> CellValue.Number(BigDecimal(42))
      ),
      Map(1 -> 0)
    )
  )

  private def styleOf(sheet: Sheet, at: ARef): Option[CellStyle] =
    sheet.cells.get(at).flatMap(_.styleId).flatMap(sheet.styleRegistry.get)

  test("T1: dimension, SST and a two-xf styles.xml; the date cell reads back as a date") {
    val path = tempXlsx("t1")
    run(ledger, excel.writeStreamStyledWithAutoDetect(path, "Ledger", Vector(date))).map { _ =>
      val sheetXml = entryText(path, "xl/worksheets/sheet1.xml")
      assert(sheetXml.contains("""<dimension ref="A1:C3"/>"""), sheetXml)
      val s = styleAttrs(path).toMap
      assertEquals(s.get("B2"), Some(Some("1")), sheetXml)
      assertEquals(s.get("B3"), Some(Some("1")), sheetXml)
      assertEquals(s.get("A2"), Some(None), sheetXml)
      assertEquals(s.get("C3"), Some(None), sheetXml)
      assert(entryText(path, "xl/sharedStrings.xml").contains("Alpha"))
      val styles = entryText(path, "xl/styles.xml")
      assert(styles.contains("""<cellXfs count="2">"""), styles)
      assert(styles.contains("""numFmtId="14""""), styles)
      assert(!styles.contains("<numFmts"), s"a built-in date format declares no numFmt: $styles")

      val wb = XlsxReader.read(path).fold(e => fail(e.message), identity)
      val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
      assertEquals(styleOf(sheet, ref"B2").map(_.numFmt), Some(NumFmt.Date))
      assertEquals(styleOf(sheet, ref"B3").map(_.numFmt), Some(NumFmt.Date))
      assertEquals(sheet.cells.get(ref"A2").flatMap(_.styleId), None)
      assertEquals(sheet.cells.get(ref"C3").flatMap(_.styleId), None)
      // the reader keeps the serial; the date lives in the style's format
      val when = sheet.cells.get(ref"B2").map(_.value).collect {
        case CellValue.Number(serial) =>
          CellValue.excelSerialToDateTime(serial.toDouble).toLocalDate
        case CellValue.DateTime(dt) => dt.toLocalDate
      }
      assertEquals(when, Some(LocalDate.of(2026, 1, 15)))
    }
  }

  test("T3: out-of-range and default-equal indices emit no s=; duplicates collapse to one xf") {
    val path = tempXlsx("t3")
    val rows = List(
      StyledRowData(
        1,
        Map(
          0 -> CellValue.Number(BigDecimal(1)),
          1 -> CellValue.Number(BigDecimal(2)),
          2 -> CellValue.Number(BigDecimal(3)),
          3 -> CellValue.Number(BigDecimal(4)),
          4 -> CellValue.Number(BigDecimal(5))
        ),
        Map(0 -> -1, 1 -> 4, 2 -> 1, 3 -> 0, 4 -> 2)
      )
    )
    // index 1 is default-equal; indices 0 and 2 are the same style
    val table = Vector(date, CellStyle.default, date, bold)
    run(rows, excel.writeStreamStyledWithAutoDetect(path, "S", table)).map { _ =>
      assertEquals(
        styleAttrs(path),
        Vector(
          "A1" -> None, // -1 falls back to the default
          "B1" -> None, // styles.size falls back to the default
          "C1" -> None, // a default-equal table entry is the default
          "D1" -> Some("1"),
          "E1" -> Some("1") // the duplicate date shares the xf
        )
      )
      // default + date + bold: the duplicate date and the default-equal entry add nothing
      val styles = entryText(path, "xl/styles.xml")
      assert(styles.contains("""<cellXfs count="3">"""), styles)
    }
  }

  test("T4: re-running the same compiled IO writes byte-identical files (styled and plain)") {
    val styled = tempXlsx("t4-styled")
    val plain = tempXlsx("t4-plain")
    val styledIo = run(ledger, excel.writeStreamStyledWithAutoDetect(styled, "L", Vector(date)))
    val plainIo = fs2.Stream
      .emits(ledger.map(_.toRowData))
      .covary[IO]
      .through(excel.writeStream(plain, "L"))
      .compile
      .drain
    for
      _ <- styledIo
      styled1 = Files.readAllBytes(styled).toVector
      _ <- plainIo
      plain1 = Files.readAllBytes(plain).toVector
      _ <- styledIo
      styled2 = Files.readAllBytes(styled).toVector
      _ <- plainIo
      plain2 = Files.readAllBytes(plain).toVector
    yield
      assertEquals(styled2, styled1, "the styled two-pass archive is byte-reproducible")
      assertEquals(plain2, plain1, "the single-pass archive is byte-reproducible")
      // the entry times carry no wall clock: every entry reads back as setTime(0L)
      (entryTimes(styled) ++ entryTimes(plain)).foreach { (name, time) =>
        assertEquals(time, epochEntryTime, s"$name carries a wall-clock time")
      }
  }

  test("T4: every streaming writer stamps its entries with time 0") {
    val auto = tempXlsx("t4-auto")
    val seq = tempXlsx("t4-seq")
    val seqAuto = tempXlsx("t4-seq-auto")
    val rows = fs2.Stream.emits(ledger.map(_.toRowData)).covary[IO]
    for
      _ <- rows.through(excel.writeStreamWithAutoDetect(auto, "A")).compile.drain
      _ <- excel.writeStreamsSeq(seq, Seq("A" -> rows, "B" -> rows))
      _ <- excel.writeStreamsSeqWithAutoDetect(seqAuto, Seq("A" -> rows, "B" -> rows))
    yield Vector(auto, seq, seqAuto).flatMap(entryTimes).foreach { (name, time) =>
      assertEquals(time, epochEntryTime, s"$name carries a wall-clock time")
    }
  }

  test("T5: the unstyled two-pass writer keeps the minimal styles and ignores RowData.cellStyles") {
    val path = tempXlsx("t5")
    val rows = fs2.Stream
      .emits(List(RowData(1, Map(0 -> CellValue.Number(BigDecimal(1))), Map(0 -> 5))))
      .covary[IO]
    rows.through(excel.writeStreamWithAutoDetect(path, "S")).compile.drain.map { _ =>
      assertEquals(
        entryText(path, "xl/styles.xml"),
        XmlUtil.compact(OoxmlStyles.minimal.toXml),
        "styles.xml is the minimal table"
      )
      val sheetXml = entryText(path, "xl/worksheets/sheet1.xml")
      assert(!sheetXml.contains(" s="), s"a source xf index is not emitted: $sheetXml")
    }
  }

  test("T6: writeStreamStyled re-declares a stale source numFmtId at 164 (GH-471 parity)") {
    val path = tempXlsx("t6")
    val stale = CellStyle.default.withNumFmtId(194, NumFmt.Custom("0.000"))
    val rows = List(StyledRowData(1, Map(0 -> CellValue.Number(BigDecimal("1.5"))), Map(0 -> 0)))
    run(rows, excel.writeStreamStyled(path, "S", Vector(stale))).map { _ =>
      val styles = entryText(path, "xl/styles.xml")
      assert(styles.contains("""<numFmt numFmtId="164" formatCode="0.000"/>"""), styles)
      assert(styles.contains("""<xf xfId="0" numFmtId="164" """), s"the xf references 164: $styles")
      assert(!styles.contains("194"), s"the source id must not ship: $styles")
      val wb = XlsxReader.read(path).fold(e => fail(e.message), identity)
      val sheet = wb.sheets.headOption.getOrElse(fail("no sheet"))
      assertEquals(styleOf(sheet, ref"A1").map(_.numFmt), Some(NumFmt.Custom("0.000")))
    }
  }

  test("T7: buildStyleTable agrees with the in-memory StyleIndex of a fresh workbook") {
    val styled = CellStyle.default
      .copy(font = Font("Arial", 9.0, bold = true))
      .withNumFmt(NumFmt.Custom("#,##0.0"))
    val (streaming, remap) = StreamingXmlWriter.buildStyleTable(Vector(styled))
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(ref"A1", CellValue.Number(BigDecimal(1)))
      .withCellStyle(ref"A1", styled)
    val (inMemory, _) = StyleIndex.fromWorkbook(Workbook(Vector(sheet)))
    assertEquals(remap, Map(0 -> 1))
    assertEquals(streaming.index.cellStyles, inMemory.cellStyles)
    assertEquals(streaming.index.fonts, inMemory.fonts)
    assertEquals(streaming.index.fills, inMemory.fills)
    assertEquals(streaming.index.borders, inMemory.borders)
    assertEquals(streaming.index.numFmts, inMemory.numFmts)
  }

  test("T8: 100k rows with one date style complete with a two-xf styles table") {
    val path = tempXlsx("t8")
    val rows = fs2.Stream
      .range(1, 100_001)
      .map { i =>
        StyledRowData(
          i,
          Map(
            0 -> CellValue.Text(s"row-${i % 50}"),
            1 -> dateCell(LocalDate.of(2026, 1, 1).plusDays((i % 365).toLong))
          ),
          Map(1 -> 0)
        )
      }
      .covary[IO]
    rows
      .through(excel.writeStreamStyledWithAutoDetect(path, "Big", Vector(date)))
      .compile
      .drain
      .map { _ =>
        val styles = entryText(path, "xl/styles.xml")
        assert(styles.contains("""<cellXfs count="2">"""), styles)
        val sheetXml = entryText(path, "xl/worksheets/sheet1.xml")
        assert(sheetXml.contains("""<dimension ref="A1:B100000"/>"""), sheetXml.take(400))
      }
  }

  test("a sheet index below 1 is refused before anything is written") {
    val path = tempXlsx("idx")
    run(
      ledger,
      excel.writeStreamStyledWithAutoDetect(path, "S", Vector(date), sheetIndex = 0)
    ).attempt
      .map(r => assert(r.isLeft, s"expected a failure, got $r"))
  }

/**
 * T2, the equivalence law: over generated rows and style tables, the two-pass styled writer emits
 * exactly the parts `writeStreamStyled` emits when given the true bounds as its dimension hint.
 */
class StreamingStyledAutoDetectLawSpec extends ScalaCheckSuite:
  import StyledAutoDetect.*

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(25)

  private val genValue: Gen[CellValue] = Gen.oneOf(
    Gen.oneOf("a", "b", "c & d", "<e>").map(CellValue.Text(_)),
    Gen.choose(-1000, 1000).map(n => CellValue.Number(BigDecimal(n))),
    Gen.oneOf(true, false).map(CellValue.Bool(_)),
    Gen.choose(0L, 3650L).map(d => dateCell(LocalDate.of(2020, 1, 1).plusDays(d)))
  )

  private val genTable: Gen[Vector[CellStyle]] =
    Gen.listOfN(4, Gen.oneOf(CellStyle.default, date, bold, percent, custom)).flatMap { pool =>
      Gen.choose(0, 4).map(n => pool.take(n).toVector)
    }

  private def genRows(tableSize: Int): Gen[List[StyledRowData]] =
    for
      n <- Gen.choose(1, 8)
      gaps <- Gen.listOfN(n, Gen.choose(1, 3))
      rows <- Gen.sequence[List[StyledRowData], StyledRowData](
        gaps.scanLeft(0)(_ + _).drop(1).map { rowIndex =>
          for
            cols <- Gen.nonEmptyListOf(Gen.choose(0, 5)).map(_.distinct)
            values <- Gen.listOfN(cols.size, genValue)
            styled <- Gen.listOfN(cols.size, Gen.option(Gen.choose(-1, tableSize)))
          yield StyledRowData(
            rowIndex,
            cols.zip(values).toMap,
            cols.zip(styled).collect { case (c, Some(s)) => c -> s }.toMap
          )
        }
      )
    yield rows

  private val parts = Vector("xl/styles.xml", "xl/worksheets/sheet1.xml", "xl/sharedStrings.xml")

  property("T2: writeStreamStyledWithAutoDetect == writeStreamStyled with the true dimension") {
    Prop.forAll(genTable.flatMap(t => genRows(t.size).map(t -> _))) { (table, rows) =>
      val bounds = BoundsAccumulator()
      rows.foreach(r => bounds.update(r.toRowData))
      val auto = tempXlsx("law-auto")
      val hinted = tempXlsx("law-hinted")
      run(rows, excel.writeStreamStyledWithAutoDetect(auto, "S", table)).unsafeRunSync()
      run(rows, excel.writeStreamStyled(hinted, "S", table, dimension = bounds.dimension))
        .unsafeRunSync()
      parts.foreach(p => assertEquals(entryText(auto, p), entryText(hinted, p), p))
      Files.deleteIfExists(auto)
      Files.deleteIfExists(hinted)
      true
    }
  }
