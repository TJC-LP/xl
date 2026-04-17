package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Sheet, Workbook}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.sheets.FreezePane

/**
 * Integration tests for `freeze` / `unfreeze` CLI commands.
 *
 * The domain model stores a `FreezePane` override on the sheet. The OOXML write layer turns that
 * override into a `<pane>` element inside `<sheetViews>`. These tests exercise the full read/write
 * round-trip by writing an xlsx and parsing the resulting `<sheetViews>` XML to verify the pane
 * attributes.
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class FreezePaneSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val outputPath: Path = Files.createTempFile("freeze-test-", ".xlsx")
  val config: WriterConfig = WriterConfig.default

  override def afterEach(context: AfterEach): Unit =
    if Files.exists(outputPath) then Files.delete(outputPath)

  /** Extract the `<pane>` element from `<sheetViews>` of the first sheet, if present. */
  private def readPaneXml(path: Path): Option[scala.xml.Elem] =
    readAllPaneXmls(path).headOption.flatMap(_._2)

  /**
   * Return a map from worksheet entry name (e.g. "xl/worksheets/sheet1.xml") to its `<pane>`
   * element, if present. Lets tests check per-sheet pane state.
   */
  private def readAllPaneXmls(path: Path): Vector[(String, Option[scala.xml.Elem])] =
    val bytes = Files.readAllBytes(path)
    val zip = new java.util.zip.ZipInputStream(new java.io.ByteArrayInputStream(bytes))
    try
      val result = Vector.newBuilder[(String, Option[scala.xml.Elem])]
      LazyList
        .continually(Option(zip.getNextEntry))
        .takeWhile(_.isDefined)
        .flatten
        .foreach { entry =>
          val name = entry.getName
          if name.startsWith("xl/worksheets/sheet") && name.endsWith(".xml") then
            val buf = new java.io.ByteArrayOutputStream()
            val data = new Array[Byte](4096)
            LazyList
              .continually(zip.read(data))
              .takeWhile(_ != -1)
              .foreach(n => buf.write(data, 0, n))
            val xml = scala.xml.XML.loadString(buf.toString("UTF-8"))
            val pane = (xml \\ "pane").headOption.collect { case e: scala.xml.Elem => e }
            result += ((name, pane))
        }
      result.result().sortBy(_._1)
    finally zip.close()

  // =========================================================================
  // freeze
  // =========================================================================

  test("freeze B2: creates pane with xSplit=1, ySplit=1, bottomRight active") {
    val sheet = Sheet("Test").put(ARef.from0(0, 0), CellValue.Text("header"))
    val wb = Workbook(sheet)

    WriteCommands
      .freeze(wb, Some(sheet), "B2", outputPath, config)
      .unsafeRunSync()

    val pane = readPaneXml(outputPath)
    assert(pane.isDefined, "Expected <pane> element in sheetViews")
    assertEquals(pane.get.attribute("xSplit").map(_.text), Some("1"))
    assertEquals(pane.get.attribute("ySplit").map(_.text), Some("1"))
    assertEquals(pane.get.attribute("topLeftCell").map(_.text), Some("B2"))
    assertEquals(pane.get.attribute("activePane").map(_.text), Some("bottomRight"))
    assertEquals(pane.get.attribute("state").map(_.text), Some("frozen"))
  }

  test("freeze A3: rows-only freeze produces ySplit without xSplit") {
    val sheet = Sheet("Test").put(ARef.from0(0, 0), CellValue.Text("hi"))
    val wb = Workbook(sheet)

    WriteCommands
      .freeze(wb, Some(sheet), "A3", outputPath, config)
      .unsafeRunSync()

    val pane = readPaneXml(outputPath)
    assert(pane.isDefined, "Expected <pane> element")
    // A3 means freeze rows 1-2, no column freeze
    assertEquals(pane.get.attribute("xSplit"), None, "No xSplit for column-less freeze")
    assertEquals(pane.get.attribute("ySplit").map(_.text), Some("2"))
    assertEquals(pane.get.attribute("topLeftCell").map(_.text), Some("A3"))
  }

  test("freeze C1: columns-only freeze produces xSplit without ySplit") {
    val sheet = Sheet("Test").put(ARef.from0(0, 0), CellValue.Text("hi"))
    val wb = Workbook(sheet)

    WriteCommands
      .freeze(wb, Some(sheet), "C1", outputPath, config)
      .unsafeRunSync()

    val pane = readPaneXml(outputPath)
    assert(pane.isDefined, "Expected <pane> element")
    assertEquals(pane.get.attribute("xSplit").map(_.text), Some("2"))
    assertEquals(pane.get.attribute("ySplit"), None, "No ySplit for row-less freeze")
    assertEquals(pane.get.attribute("topLeftCell").map(_.text), Some("C1"))
  }

  // =========================================================================
  // unfreeze
  // =========================================================================

  test("unfreeze: removes pane from sheet that had one") {
    // Step 1: freeze B2
    val sheet = Sheet("Test").put(ARef.from0(0, 0), CellValue.Text("hi"))
    WriteCommands
      .freeze(wb = Workbook(sheet), sheetOpt = Some(sheet), refStr = "B2", outputPath, config)
      .unsafeRunSync()
    assert(readPaneXml(outputPath).isDefined, "Setup: pane should exist before unfreeze")

    // Step 2: unfreeze
    val frozenWb = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    WriteCommands
      .unfreeze(frozenWb, Some(frozenWb.sheets.head), outputPath, config)
      .unsafeRunSync()

    assert(readPaneXml(outputPath).isEmpty, "Pane should be removed after unfreeze")
  }

  // =========================================================================
  // Domain model
  // =========================================================================

  test("FreezePane: Sheet.freezeAt sets At override") {
    val sheet = Sheet("Test").freezeAt(ARef.from0(1, 1))
    assertEquals(sheet.freezePane, Some(FreezePane.At(ARef.from0(1, 1))))
  }

  test("FreezePane: Sheet.unfreeze sets Remove override (not None)") {
    val sheet = Sheet("Test").freezeAt(ARef.from0(1, 1)).unfreeze
    assertEquals(
      sheet.freezePane,
      Some(FreezePane.Remove),
      "unfreeze sets Remove, distinct from None (preserve)"
    )
  }

  // =========================================================================
  // Qualified-ref support (freeze with `Sheet2!B2` syntax)
  // =========================================================================

  test("freeze Sheet2!B2: qualified ref targets the named sheet without -s") {
    val s1 = Sheet("Sheet1")
    val s2 = Sheet("Sheet2")
    val wb = Workbook(Vector(s1, s2))

    // No default sheet passed; qualified ref selects Sheet2.
    WriteCommands
      .freeze(wb, None, "Sheet2!B2", outputPath, config)
      .unsafeRunSync()

    // Verify per-worksheet XML: only sheet2.xml should carry a <pane> element.
    val panes = readAllPaneXmls(outputPath)
    assertEquals(panes.length, 2, "Should have two worksheets")
    assertEquals(panes(0)._2, None, "sheet1.xml must not have a pane")
    val sheet2Pane = panes(1)._2
    assert(sheet2Pane.isDefined, "sheet2.xml should carry the pane element")
    assertEquals(sheet2Pane.get.attribute("topLeftCell").map(_.text), Some("B2"))
  }

  test("freeze: range argument errors with clear message") {
    val sheet = Sheet("Test")
    val wb = Workbook(sheet)

    val err = WriteCommands
      .freeze(wb, Some(sheet), "A1:B2", outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(err.isLeft, "Expected range-arg rejection")
    val msg = err.swap.getOrElse(throw new Exception("expected left")).getMessage
    assert(msg.contains("single cell"), s"Expected 'single cell' in message: $msg")
  }
