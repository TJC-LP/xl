package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Workbook, Sheet}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.cli.contract.{CliError, CliException, ErrorCode, ExitCodes}
import com.tjclp.xl.cli.helpers.BatchParser
import com.tjclp.xl.cli.helpers.BatchParser.BatchOp
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.sheets.syntax.*
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Integration tests for batch put and fill pattern functionality.
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class BatchPutSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val outputPath: Path = Files.createTempFile("test", ".xlsx")
  val config: WriterConfig = WriterConfig.default

  override def afterEach(context: AfterEach): Unit =
    if Files.exists(outputPath) then Files.delete(outputPath)

  // ========== Put Command Mode Detection ==========

  test("put: single cell mode") {
    val wb = Workbook(Sheet("Test"))
    val result =
      WriteCommands
        .put(wb, Some(wb.sheets.head), "A1", List("100"), outputPath, config)
        .unsafeRunSync()

    assert(result.contains("Put: A1 = 100"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val cellValue = imported.sheets.head.cells.get(ARef.from0(0, 0)).map(_.value)
    assertEquals(cellValue, Some(CellValue.Number(BigDecimal("100"))))
  }

  test("put: ISO date detection applies Date number format") {
    val wb = Workbook(Sheet("Test"))
    WriteCommands
      .put(wb, Some(wb.sheets.head), "A1", List("2025-11-10"), outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    val date = sheet.cells.get(ARef.from0(0, 0)).map(_.value) match
      case Some(CellValue.DateTime(dateTime)) => dateTime.toLocalDate
      case Some(CellValue.Number(serial)) =>
        CellValue.excelSerialToDateTime(serial.toDouble).toLocalDate
      case other => fail(s"Expected detected DateTime, got $other")
    assertEquals(date.toString, "2025-11-10")
    assertEquals(sheet.getCellStyle(ARef.from0(0, 0)).map(_.numFmt), Some(NumFmt.Date))
  }

  test("put: detect=false preserves ISO date-like input as text") {
    val wb = Workbook(Sheet("Test"))
    WriteCommands
      .put(
        wb,
        Some(wb.sheets.head),
        "A1",
        List("2025-11-10"),
        outputPath,
        config,
        detect = false
      )
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    assertEquals(
      sheet.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Text("2025-11-10"))
    )
    assertEquals(sheet.getCellStyle(ARef.from0(0, 0)).map(_.numFmt), None)
  }

  test("put: fill pattern mode (1 value, range)") {
    val wb = Workbook(Sheet("Test"))
    val result =
      WriteCommands
        .put(wb, Some(wb.sheets.head), "A1:A5", List("TBD"), outputPath, config)
        .unsafeRunSync()

    assert(result.contains("Filled 5 cells"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    (0 to 4).foreach { row =>
      assertEquals(sheet.cells.get(ARef.from0(0, row)).map(_.value), Some(CellValue.Text("TBD")))
    }
  }

  test("put: batch values mode (N values, N-cell range)") {
    val wb = Workbook(Sheet("Test"))
    val result =
      WriteCommands
        .put(wb, Some(wb.sheets.head), "A1:A3", List("10", "20", "30"), outputPath, config)
        .unsafeRunSync()

    assert(result.contains("Put 3 values"))
    assert(result.contains("row-major"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    assertEquals(
      sheet.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Number(BigDecimal("10")))
    )
    assertEquals(
      sheet.cells.get(ARef.from0(0, 1)).map(_.value),
      Some(CellValue.Number(BigDecimal("20")))
    )
    assertEquals(
      sheet.cells.get(ARef.from0(0, 2)).map(_.value),
      Some(CellValue.Number(BigDecimal("30")))
    )
  }

  test("put: batch values 2D range (row-major)") {
    val wb = Workbook(Sheet("Test"))
    val result =
      WriteCommands
        .put(wb, Some(wb.sheets.head), "A1:B2", List("1", "2", "3", "4"), outputPath, config)
        .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    // Row-major: A1, B1, A2, B2
    assertEquals(
      sheet.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Number(BigDecimal("1")))
    ) // A1
    assertEquals(
      sheet.cells.get(ARef.from0(1, 0)).map(_.value),
      Some(CellValue.Number(BigDecimal("2")))
    ) // B1
    assertEquals(
      sheet.cells.get(ARef.from0(0, 1)).map(_.value),
      Some(CellValue.Number(BigDecimal("3")))
    ) // A2
    assertEquals(
      sheet.cells.get(ARef.from0(1, 1)).map(_.value),
      Some(CellValue.Number(BigDecimal("4")))
    ) // B2
  }

  test("put: error - count mismatch (too few values)") {
    val wb = Workbook(Sheet("Test"))
    val result = WriteCommands
      .put(wb, Some(wb.sheets.head), "A1:A5", List("1", "2", "3"), outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(result.isLeft)
    val error = result.swap.getOrElse(throw new Exception("Expected error"))
    assert(error.getMessage.contains("5 cells but 3 values"))
    assert(error.getMessage.contains("Hint"))
  }

  test("put: error - count mismatch (too many values)") {
    val wb = Workbook(Sheet("Test"))
    val result = WriteCommands
      .put(wb, Some(wb.sheets.head), "A1:A2", List("1", "2", "3", "4"), outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(result.isLeft)
    val error = result.swap.getOrElse(throw new Exception("Expected error"))
    assert(error.getMessage.contains("2 cells but 4 values"))
  }

  test("put: error - multiple values to single cell") {
    val wb = Workbook(Sheet("Test"))
    val result = WriteCommands
      .put(wb, Some(wb.sheets.head), "A1", List("1", "2", "3"), outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(result.isLeft)
    val error = result.swap.getOrElse(throw new Exception("Expected error"))
    assert(error.getMessage.contains("Cannot put 3 values to single cell"))
    assert(error.getMessage.contains("Use a range"))
  }

  // ========== Putf Command Mode Detection ==========

  test("putf: single cell mode") {
    val wb = Workbook(Sheet("Test").put(ARef.from0(0, 0), CellValue.Number(BigDecimal("10"))))
    val result = WriteCommands
      .putFormula(wb, Some(wb.sheets.head), "B1", List("=A1*2"), outputPath, config)
      .unsafeRunSync()

    assert(result.contains("Put: B1"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    sheet.cells.get(ARef.from0(1, 0)).map(_.value) match
      case Some(CellValue.Formula(formula, _, _)) =>
        assertEquals(formula, "A1*2")
      case other => fail(s"Expected Formula, got $other")
  }

  test("putf: formula dragging mode preserves $ anchors") {
    val wb = Workbook(
      Sheet("Test")
        .put(ARef.from0(0, 0), CellValue.Number(BigDecimal("10")))
        .put(ARef.from0(0, 1), CellValue.Number(BigDecimal("20")))
        .put(ARef.from0(0, 2), CellValue.Number(BigDecimal("30")))
    )

    val result = WriteCommands
      .putFormula(wb, Some(wb.sheets.head), "B1:B3", List("=SUM($A$1:A1)"), outputPath, config)
      .unsafeRunSync()

    assert(result.contains("with anchor-aware dragging"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    // Verify $ anchors work
    sheet.cells.get(ARef.from0(1, 0)).map(_.value) match
      case Some(CellValue.Formula(formula, _, _)) => assert(formula.contains("$A$1"))
      case other => fail(s"Expected Formula with $$A$$1, got $other")
  }

  test("GH-455: putf dragging keeps grouping parens around a product divisor") {
    val wb = Workbook(
      Sheet("Test")
        .put(ARef.from0(0, 0), CellValue.Number(BigDecimal("4")))
        .put(ARef.from0(1, 0), CellValue.Number(BigDecimal("2")))
        .put(ARef.from0(2, 0), CellValue.Number(BigDecimal("2")))
    )

    val result = WriteCommands
      .putFormula(wb, Some(wb.sheets.head), "D1:D3", List("=A1/(B1*C1)"), outputPath, config)
      .unsafeRunSync()

    assert(result.contains("with anchor-aware dragging"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    // Dragging reprints through FormulaPrinter: the divisor group must survive, or the
    // formula silently changes value (A2/B2*C2 is (A2/B2)*C2, not A2/(B2*C2)).
    sheet.cells.get(ARef.from0(3, 0)).map(_.value) match
      case Some(CellValue.Formula(formula, _, _)) => assertEquals(formula, "A1/(B1*C1)")
      case other => fail(s"Expected Formula, got $other")
    sheet.cells.get(ARef.from0(3, 1)).map(_.value) match
      case Some(CellValue.Formula(formula, _, _)) => assertEquals(formula, "A2/(B2*C2)")
      case other => fail(s"Expected Formula, got $other")
  }

  test("putf: batch formulas mode (no dragging)") {
    val wb = Workbook(
      Sheet("Test")
        .put(ARef.from0(0, 0), CellValue.Number(BigDecimal("1")))
        .put(ARef.from0(1, 0), CellValue.Number(BigDecimal("2")))
        .put(ARef.from0(0, 1), CellValue.Number(BigDecimal("3")))
        .put(ARef.from0(1, 1), CellValue.Number(BigDecimal("4")))
    )

    val result = WriteCommands
      .putFormula(
        wb,
        Some(wb.sheets.head),
        "C1:C3",
        List("=A1+B1", "=A2*B2", "=100"),
        outputPath,
        config
      )
      .unsafeRunSync()

    assert(result.contains("explicit, no dragging"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val sheet = imported.sheets.head
    // Verify formulas are as-is (no dragging)
    sheet.cells.get(ARef.from0(2, 0)).map(_.value) match
      case Some(CellValue.Formula(formula, _, _)) => assertEquals(formula, "A1+B1")
      case other => fail(s"Expected Formula A1+B1, got $other")
    sheet.cells.get(ARef.from0(2, 1)).map(_.value) match
      case Some(CellValue.Formula(formula, _, _)) => assertEquals(formula, "A2*B2")
      case other => fail(s"Expected Formula A2*B2, got $other")
    sheet.cells.get(ARef.from0(2, 2)).map(_.value) match
      case Some(CellValue.Formula(formula, _, _)) => assertEquals(formula, "100")
      case other => fail(s"Expected Formula 100, got $other")
  }

  test("putf: error - count mismatch") {
    val wb = Workbook(Sheet("Test"))
    val result = WriteCommands
      .putFormula(wb, Some(wb.sheets.head), "B1:B5", List("=X", "=Y"), outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(result.isLeft)
    val error = result.swap.getOrElse(throw new Exception("Expected error"))
    assert(error.getMessage.contains("5 cells but 2 formulas"))
    assert(error.getMessage.contains("Hint"))
  }

  // ========== Batch JSON: registry-backed parsing (ADR-017 §2.6) ==========

  private def parseError(json: String): CliError =
    BatchParser.parseBatchJson(json) match
      case Left(e: CliException) => e.error
      case Left(other) => fail(s"expected a CliException, got ${other.getClass.getName}: $other")
      case Right(r) => fail(s"expected a parse failure, got $r")

  private def parseOk(json: String): BatchParser.ParseResult =
    BatchParser.parseBatchJson(json) match
      case Right(r) => r
      case Left(e) => fail(s"unexpected parse failure: ${e.getMessage}")

  test("batch: an unknown op is BATCH_OP_UNKNOWN and suggests the nearest names") {
    val error = parseError("""[{"op":"putff","ref":"A1","value":"=1"}]""")
    assertEquals(error.code, ErrorCode.BATCH_OP_UNKNOWN)
    assert(
      error.message.startsWith("Object 1: Unknown operation 'putff'. Valid: put, putf, style"),
      error.message
    )
    assert(error.message.contains("Did you mean: putf, put"), error.message)
    assertEquals(error.candidates, Vector("putf", "put"))
    assertEquals(error.location.flatMap(_.opIndex), Some(1))
    assertEquals(error.exitCode, ExitCodes.usage)
  }

  test("batch: an unknown op with no near name keeps the historical text, no suggestion") {
    val error = parseError("""[{"op":"frobnicate","ref":"A1"}]""")
    assertEquals(error.code, ErrorCode.BATCH_OP_UNKNOWN)
    assert(!error.message.contains("Did you mean"), error.message)
    assert(error.message.endsWith("page-setup, header-footer, cf"), error.message)
    assertEquals(error.candidates, Vector.empty)
  }

  test("batch: op names are accepted in camelCase and kebab-case alike") {
    val result =
      parseOk("""[{"op":"removeComment","ref":"A1"},{"op":"remove-comment","ref":"A2"}]""")
    assertEquals(result.ops, Vector(BatchOp.RemoveComment("A1"), BatchOp.RemoveComment("A2")))
    assertEquals(result.warnings, Vector.empty)
  }

  test("batch: alias keys are honoured and raise no unknown-property warning") {
    val json = """[
      {"op":"put","ref":"A1","value":1,"numFormat":"percent"},
      {"op":"put","ref":"A2","value":1,"num-format":"percent"},
      {"op":"putf","ref":"B1:B2","value":"=A1","anchor":"B1"},
      {"op":"putf","ref":"B3","formula":"=A1"},
      {"op":"hyperlink","ref":"C1","url":"https://example.com"},
      {"op":"style","range":"A1","halign":"center","format":"percent","font-size":14,"border-top":"thin"},
      {"op":"page-setup","fit-to-height":0,"fit-to-width":1},
      {"op":"copy","source":"A1","target":"D1","values-only":true}
    ]"""
    val result = parseOk(json)
    assertEquals(result.warnings, Vector.empty, "aliases never warn")
    result.ops match
      case Vector(
            BatchOp.Put(_, _, Some(NumFmt.Percent)),
            BatchOp.Put(_, _, Some(NumFmt.Percent)),
            BatchOp.PutFormulaDragging("B1:B2", "=A1", "B1", None),
            BatchOp.PutFormula("B3", "=A1", None),
            BatchOp.Hyperlink("C1", Some("https://example.com")),
            BatchOp.Style("A1", props),
            BatchOp.SetPageSetup(None, None, Some(1), Some(0), None),
            BatchOp.CopyRange("A1", "D1", true)
          ) =>
        assertEquals(props.align, Some("center"))
        assertEquals(props.numFormat, Some("percent"))
        assertEquals(props.fontSize, Some(14.0))
        assertEquals(props.borderTop, Some("thin"))
      case other => fail(s"unexpected ops: $other")
  }

  test("batch: the canonical spelling wins when it and an alias are both present") {
    val result =
      parseOk("""[{"op":"put","ref":"A1","value":1,"format":"percent","numFormat":"currency"}]""")
    assertEquals(
      result.ops.map { case BatchOp.Put(_, _, f) => f; case _ => None },
      Vector(Some(NumFmt.Percent))
    )
    assertEquals(result.warnings, Vector.empty)
  }

  test("batch: an unknown key still warns, naming the op and index and the known properties") {
    val result = parseOk("""[{"op":"put","ref":"A1","value":1,"colour":"red"}]""")
    assertEquals(result.warnings.size, 1)
    val warning = result.warnings.head
    assert(
      warning.startsWith("Warning: Object 1 (put): unknown properties ignored: colour"),
      warning
    )
    assert(warning.contains("Known: "), warning)
    assert(
      warning.contains("format") && warning.contains("sheet") && warning.contains("detect"),
      warning
    )
  }

  test("batch: a non-array document is BATCH_JSON_INVALID (exit 2)") {
    BatchParser
      .parseBatchOperations("""{"op":"put","ref":"A1","value":1}""")
      .attempt
      .unsafeRunSync() match
      case Left(e: CliException) =>
        assertEquals(e.error.code, ErrorCode.BATCH_JSON_INVALID)
        assertEquals(e.getMessage, "Batch input must be a JSON array")
        assertEquals(e.error.exitCode, ExitCodes.usage)
      case other => fail(s"expected a CliException, got $other")
    assertEquals(parseError("""{"op":"put"}""").code, ErrorCode.BATCH_JSON_INVALID)
    val malformed = parseError("""[{"op": "put", "ref": "A1", value: unquoted}]""")
    assertEquals(malformed.code, ErrorCode.BATCH_JSON_INVALID)
    assert(malformed.message.startsWith("JSON parse error"), malformed.message)
  }

  test("batch: a shape error is BATCH_OP_INVALID carrying the op index") {
    val missingValue = parseError("""[{"op":"put","ref":"A1"}]""")
    assertEquals(missingValue.code, ErrorCode.BATCH_OP_INVALID)
    assertEquals(missingValue.message, "Object 1: Missing 'value' field")
    assertEquals(missingValue.location.flatMap(_.opIndex), Some(1))
    assertEquals(missingValue.exitCode, ExitCodes.usage)
    val notObject = parseError("""[{"op":"put","ref":"A1","value":1}, 42]""")
    assertEquals(notObject.code, ErrorCode.BATCH_OP_INVALID)
    assertEquals(notObject.location.flatMap(_.opIndex), Some(2))
    val noOp = parseError("""[{"ref":"A1"}]""")
    assertEquals(noOp.code, ErrorCode.BATCH_OP_INVALID)
    assertEquals(noOp.message, "Object 1: Missing or invalid 'op' field")
    val badRow = parseError("""[{"op":"rowheight","row":"two","height":9}]""")
    assertEquals(badRow.code, ErrorCode.BATCH_OP_INVALID)
    assert(badRow.message.contains("'row'"), badRow.message)
  }
