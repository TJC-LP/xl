package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{CellRange, Sheet, Workbook}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.cli.commands.{DiffCommands, WriteCommands}
import com.tjclp.xl.cli.helpers.{BatchParser, CopyOps, ValueParser}
import com.tjclp.xl.cli.output.{Format, JsonRenderer, Markdown}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * GH-430 CLI semantics for formula records: putf rejects `TABLE(` (a record, not a function),
 * copies materialize/degrade like Excel paste, overwrites drop the record, and display renders
 * `{=...}` braces plus a JSON `formulaKind` field for non-Normal kinds.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class FormulaRecordCliSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  private val outputPath: Path = Files.createTempFile("xl-record-cli", ".xlsx")
  private val config: WriterConfig = WriterConfig.default

  override def afterEach(context: AfterEach): Unit =
    if Files.exists(outputPath) then Files.delete(outputPath)

  private def range(a1: String): CellRange = CellRange.parse(a1).fold(fail(_), identity)
  private def aref(a1: String): ARef = ARef.parse(a1).fold(fail(_), identity)
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  private val tableKind: FormulaKind.DataTable =
    FormulaKind.DataTable(
      range("F2:G2"),
      dt2D = true,
      dtr = false,
      r1 = Some(aref("A1")),
      r2 = Some(aref("A2"))
    )

  private def recordSheet: Sheet =
    Sheet("S")
      .put(aref("A1"), num(8))
      .put(aref("A2"), num(3))
      .put(aref("F2"), CellValue.dataTable(tableKind, Some(num(42))))
      .put(
        aref("C1"),
        CellValue.Formula(
          "SUM(A1:A2*10)",
          Some(num(110)),
          FormulaKind.ArrayFormula(range("C1:C1"))
        )
      )

  test("putf rejects a top-level TABLE( expression, steering to the GH-419 authoring surfaces") {
    val wb = Workbook(recordSheet)
    val error = intercept[Exception] {
      WriteCommands
        .putFormula(wb, Some(wb.sheets.head), "B9", List("=TABLE(A1,A2)"), outputPath, config)
        .unsafeRunSync()
    }
    assert(error.getMessage.contains("GH-419"), s"unexpected message: ${error.getMessage}")
    assert(
      error.getMessage.contains(
        """the data-table batch op ({"op":"data-table","ref":"D5:F6","rowInput":"B1","colInput":"B2"})"""
      ),
      s"steering must name the batch op: ${error.getMessage}"
    )
    assert(
      error.getMessage.contains("sheet.dataTable(interior, rowInput, colInput)"),
      s"steering must name the scripting API: ${error.getMessage}"
    )
    assert(!Files.exists(outputPath) || Files.size(outputPath) == 0L, "nothing must be written")
  }

  test("putf TABLE( rejection is case-insensitive and =-stripped; batch putf rejects too") {
    assert(ValueParser.dataTableFormulaError("table(A1,)").isDefined)
    assert(ValueParser.dataTableFormulaError("  =TaBlE(,B5)").isDefined)
    assert(ValueParser.dataTableFormulaError("=TABLEX(A1)").isEmpty)
    assert(ValueParser.dataTableFormulaError("=SUM(TABLE1)").isEmpty)
    val json = """[{"op":"putf","ref":"B9","formula":"=TABLE(A1,A2)"}]"""
    BatchParser.parseBatchJson(json) match
      case Left(err) =>
        assert(err.getMessage.contains("GH-419"), err.getMessage)
        assert(
          err.getMessage.contains("sheet.dataTable(interior, rowInput, colInput)"),
          err.getMessage
        )
      case Right(_) => fail("batch putf must reject TABLE( formulas")
  }

  test("putf of a real formula onto a record anchor replaces it with a Normal formula") {
    val wb = Workbook(recordSheet)
    WriteCommands
      .putFormula(wb, Some(wb.sheets.head), "F2", List("=A1*3"), outputPath, config)
      .unsafeRunSync()
    val reread = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    reread.sheets.head.cells.get(aref("F2")).map(_.value) match
      case Some(CellValue.Formula(expr, _, kind)) =>
        assertEquals(expr, "A1*3")
        assertEquals(kind, FormulaKind.Normal())
      case other => fail(s"expected Normal formula at F2, got $other")
  }

  test("put of a constant onto a record anchor drops the record from the output") {
    val wb = Workbook(recordSheet)
    WriteCommands
      .put(wb, Some(wb.sheets.head), "F2", List("7"), outputPath, config)
      .unsafeRunSync()
    val reread = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    assertEquals(reread.sheets.head.cells.get(aref("F2")).map(_.value), Some(num(7)))
  }

  test("copy of a dataTable cell materializes its cached constant (never a TABLE( paste)") {
    val sheet = recordSheet
    val wb = Workbook(sheet)
    val out = CopyOps.copyRange(
      wb,
      sheet,
      range("F2:F2"),
      sheet,
      range("D9:D9"),
      valuesOnly = false
    )
    val s2 = out.sheets.find(_.name == SheetName.unsafe("S")).getOrElse(fail("missing S"))
    assertEquals(s2(aref("D9")).value, num(42))
    // Source record untouched.
    assertEquals(s2(aref("F2")).value, CellValue.dataTable(tableKind, Some(num(42))))
  }

  test("fill of a dataTable cell materializes its cached constant at each destination") {
    val sheet = recordSheet
    val wb = Workbook(sheet)
    WriteCommands
      .fill(wb, Some(sheet), "F2", "F2:F4", FillDirection.Down, outputPath, config)
      .unsafeRunSync()

    val reread = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val cells = reread.sheets.headOption.map(_.cells).getOrElse(fail("missing output sheet"))
    assertEquals(
      cells.get(aref("F2")).map(_.value),
      Some(CellValue.dataTable(tableKind, Some(num(42))))
    )
    assertEquals(cells.get(aref("F3")).map(_.value), Some(num(42)))
    assertEquals(cells.get(aref("F4")).map(_.value), Some(num(42)))
  }

  test("sort materializes a moved dataTable cell instead of creating a TABLE formula") {
    val sheet = recordSheet
    val wb = Workbook(sheet)
    WriteCommands
      .sort(
        wb,
        Some(sheet),
        "A1:F2",
        List(SortKey("A", SortDirection.Ascending, SortMode.Numeric)),
        hasHeader = false,
        outputPath = outputPath,
        config = config
      )
      .unsafeRunSync()

    val reread = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val cells = reread.sheets.headOption.map(_.cells).getOrElse(fail("missing output sheet"))
    assertEquals(cells.get(aref("F1")).map(_.value), Some(num(42)))
  }

  test("copy of an array anchor pastes a plain (Normal) shifted formula") {
    val sheet = recordSheet
    val wb = Workbook(sheet)
    val out = CopyOps.copyRange(
      wb,
      sheet,
      range("C1:C1"),
      sheet,
      range("C5:C5"),
      valuesOnly = false
    )
    val s2 = out.sheets.find(_.name == SheetName.unsafe("S")).getOrElse(fail("missing S"))
    s2(aref("C5")).value match
      case CellValue.Formula(expr, _, kind) =>
        assertEquals(expr, "SUM(A5:A6*10)")
        assertEquals(kind, FormulaKind.Normal())
      case other => fail(s"expected pasted Normal formula, got $other")
  }

  test("diff compares formula record kinds and payloads while ignoring cached values") {
    val expression = "SUM(A1:A2)"
    val normal = CellValue.Formula(expression, Some(num(3)))
    val arrayC1 = CellValue.Formula(
      expression,
      Some(num(3)),
      FormulaKind.ArrayFormula(range("C1:C1"))
    )
    val arrayC1DifferentCache = CellValue.Formula(
      expression,
      Some(num(99)),
      FormulaKind.ArrayFormula(range("C1:C1"))
    )
    val arrayC1C2 = CellValue.Formula(
      expression,
      Some(num(3)),
      FormulaKind.ArrayFormula(range("C1:C2"))
    )

    def diff(left: CellValue, right: CellValue): DiffCommands.WorkbookDiff =
      DiffCommands
        .computeDiff(
          Workbook(Sheet("S").put(aref("C1"), left)),
          Workbook(Sheet("S").put(aref("C1"), right)),
          None
        )
        .fold(fail(_), identity)

    assert(!diff(normal, arrayC1).identical, "Normal and array records must differ")
    assert(!diff(arrayC1, arrayC1C2).identical, "array ref payloads must differ")
    assert(
      diff(arrayC1, arrayC1DifferentCache).identical,
      "derived caches must remain outside formula identity"
    )
  }

  test("view --formulas renders {=...} braces for record kinds") {
    val rendered = Markdown.renderRange(recordSheet, range("A1:G2"), showFormulas = true)
    assert(rendered.contains("{=TABLE(A1,A2)}"), rendered)
    assert(rendered.contains("{=SUM(A1:A2*10)}"), rendered)
  }

  test("JSON cell output carries an additive formulaKind field for record kinds") {
    val json = JsonRenderer.renderRange(recordSheet, range("A1:G2"))
    assert(json.contains(""""formulaKind": "dataTable""""), json)
    assert(json.contains(""""formulaKind": "array""""), json)
    // Normal formulas gain no kind field.
    val normalJson = JsonRenderer.renderRange(
      Sheet("N").put(aref("A1"), CellValue.Formula("1+1")),
      range("A1:A1")
    )
    assert(!normalJson.contains("formulaKind"), normalJson)
  }

  test("--stream batch sibling put passes record <f> elements through verbatim") {
    import com.tjclp.xl.cli.commands.StreamingWriteCommands
    import com.tjclp.xl.ooxml.XlsxWriter

    val sourcePath = Files.createTempFile("xl-record-stream-src", ".xlsx")
    val jsonPath = Files.createTempFile("xl-record-stream-ops", ".json")
    try
      XlsxWriter
        .write(Workbook(recordSheet), sourcePath)
        .fold(err => fail(err.message), identity)
      Files.writeString(jsonPath, """[{"op":"put","ref":"A5","value":"dirty"}]""")
      StreamingWriteCommands
        .batch(sourcePath, outputPath, Some("S"), jsonPath.toString)
        .unsafeRunSync()

      // The transform's untouched-cell path re-emits every attribute verbatim; empty elements
      // canonicalize to <f ...></f> (same as <dimension> on this path), which re-reads
      // identically. The lock is attr fidelity + the semantic re-read, not tag minimization.
      val sheetXml = zipEntryText(outputPath, "xl/worksheets/sheet1.xml")
      assert(
        sheetXml.contains("""<f t="dataTable" ref="F2:G2" dt2D="1" dtr="0" r1="A1" r2="A2">"""),
        s"streaming transform degraded the dataTable record: $sheetXml"
      )
      assert(
        sheetXml.contains("""<f t="array" ref="C1:C1">SUM(A1:A2*10)</f>"""),
        s"streaming transform degraded the array record: $sheetXml"
      )
      assert(sheetXml.contains("<v>42</v>"), sheetXml)

      val reread = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val cells = reread.sheets.headOption.map(_.cells).getOrElse(Map.empty)
      assertEquals(
        cells.get(aref("F2")).map(_.value),
        Some(CellValue.dataTable(tableKind, Some(num(42))))
      )
    finally
      Files.deleteIfExists(sourcePath)
      Files.deleteIfExists(jsonPath)
  }

  private def zipEntryText(path: Path, name: String): String =
    val zip = new java.util.zip.ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) =>
          new String(zip.getInputStream(e).readAllBytes(), java.nio.charset.StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found in $path")
    finally zip.close()

  test("cell detail shows the braced record expression") {
    val info = Format.cellInfo(
      aref("F2"),
      CellValue.dataTable(tableKind, Some(num(42))),
      formatted = "42",
      style = None,
      comment = None,
      hyperlink = None,
      dependencies = Vector.empty,
      dependents = Vector.empty
    )
    assert(info.contains("Formula: {=TABLE(A1,A2)}"), info)
  }
