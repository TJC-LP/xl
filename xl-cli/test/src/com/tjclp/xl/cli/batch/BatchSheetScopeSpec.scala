package com.tjclp.xl.cli.batch

import munit.FunSuite

import cats.effect.unsafe
import com.tjclp.xl.{Sheet, Workbook}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.{CliException, ErrorCode, ExitCodes}
import com.tjclp.xl.cli.helpers.BatchParser
import com.tjclp.xl.cli.helpers.BatchParser.BatchOp
import com.tjclp.xl.sheets.syntax.*

/**
 * Every batch op accepts an optional `"sheet"` key (ADR-017 §2.6). The sheet for an op is: a
 * sheet-qualified ref > the op's `sheet` > the batch default; a qualified ref that disagrees with
 * `sheet` is a usage error naming the op; `rename-sheet` of the default retargets the ops after it;
 * every apply-time failure carries the op index.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class BatchSheetScopeSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  private val a1 = ARef.from0(0, 0)
  private val b1 = ARef.from0(1, 0)
  private val b2 = ARef.from0(1, 1)

  private def data: Sheet = Sheet("Data").put(a1, CellValue.Text("keep"))

  /** Content in A1/B1/B2 and a frozen pane, so every scoped op below changes something. */
  private def other: Sheet =
    Sheet("Other")
      .put(a1, CellValue.Text("x"))
      .put(b1, CellValue.Number(BigDecimal(1)))
      .put(b2, CellValue.Number(BigDecimal(2)))
      .freezeAt(b2)

  private def book: Workbook = Workbook(Vector(data, other))

  private def run(wb: Workbook, default: Option[Sheet], json: String): Either[Throwable, Workbook] =
    BatchParser
      .parseBatchOperations(json)
      .flatMap(result => BatchParser.applyScoped(wb, default, result.scoped))
      .attempt
      .unsafeRunSync()

  private def cliError(result: Either[Throwable, ?]): CliException = result match
    case Left(e: CliException) => e
    case Left(other) => fail(s"expected a CliException, got ${other.getClass.getName}: $other")
    case Right(_) => fail("expected a failure")

  private def sheet(wb: Workbook, name: String): Sheet =
    wb.sheets.find(_.name.value == name).getOrElse(fail(s"no sheet $name in ${wb.sheetNames}"))

  private val scopedOps: Vector[(String, String)] = Vector(
    "put" -> """{"op":"put","sheet":"Other","ref":"C1","value":7}""",
    "putf" -> """{"op":"putf","sheet":"Other","ref":"C2","value":"=B1+B2"}""",
    "putf-values" -> """{"op":"putf","sheet":"Other","ref":"C3:C4","values":["=1","=2"]}""",
    "putf-drag" -> """{"op":"putf","sheet":"Other","ref":"D1:D2","value":"=B1*2","from":"D1"}""",
    "put-values" -> """{"op":"put","sheet":"Other","ref":"E1:E2","values":[1,2]}""",
    "style" -> """{"op":"style","sheet":"Other","range":"A1","bold":true}""",
    "merge" -> """{"op":"merge","sheet":"Other","range":"F1:G1"}""",
    "colwidth" -> """{"op":"colwidth","sheet":"Other","col":"A","width":30}""",
    "rowheight" -> """{"op":"rowheight","sheet":"Other","row":1,"height":40}""",
    "comment" -> """{"op":"comment","sheet":"Other","ref":"A1","text":"n"}""",
    "hyperlink" -> """{"op":"hyperlink","sheet":"Other","ref":"A1","target":"https://x.test"}""",
    "clear" -> """{"op":"clear","sheet":"Other","range":"A1","all":true}""",
    "col-hide" -> """{"op":"col-hide","sheet":"Other","col":"B"}""",
    "col-show" -> """{"op":"col-show","sheet":"Other","col":"B"}""",
    "row-hide" -> """{"op":"row-hide","sheet":"Other","row":2}""",
    "row-show" -> """{"op":"row-show","sheet":"Other","row":2}""",
    "group-rows" -> """{"op":"group-rows","sheet":"Other","rows":"1:2","level":1}""",
    "group-cols" -> """{"op":"group-cols","sheet":"Other","cols":"A:B","level":1}""",
    "autofit" -> """{"op":"autofit","sheet":"Other","columns":"A:B"}""",
    "freeze" -> """{"op":"freeze","sheet":"Other","ref":"C3"}""",
    "unfreeze" -> """{"op":"unfreeze","sheet":"Other"}""",
    "copy" -> """{"op":"copy","sheet":"Other","source":"A1","target":"H1"}""",
    "sheet-view" -> """{"op":"sheet-view","sheet":"Other","gridlines":false}""",
    "tab-color" -> """{"op":"tab-color","sheet":"Other","color":"#FF0000"}""",
    "autofilter" -> """{"op":"autofilter","sheet":"Other","range":"A1:B2"}""",
    "page-setup" -> """{"op":"page-setup","sheet":"Other","orientation":"landscape"}""",
    "header-footer" -> """{"op":"header-footer","sheet":"Other","oddFooter":"&LConfidential"}""",
    "cf" -> """{"op":"cf","sheet":"Other","range":"B1:B2","rule":"cellIs:greaterThan:1","bold":true}""",
    "chart" -> """{"op":"chart","sheet":"Other","type":"column","data":"B1:B2","at":"J1:M8"}"""
  )

  scopedOps.foreach { (name, json) =>
    test(s"$name with sheet=Other lands on Other and leaves the default sheet untouched") {
      val wb = book
      val result = run(wb, Some(data), s"[$json]") match
        case Right(w) => w
        case Left(e) => fail(s"$name failed: ${e.getMessage}")
      assertEquals(sheet(result, "Data"), data, s"$name touched the default sheet")
      assertNotEquals(sheet(result, "Other"), other, s"$name changed nothing on Other")
    }
  }

  test("remove-comment with sheet=Other removes the comment there") {
    val wb = book
    val json =
      """[{"op":"comment","sheet":"Other","ref":"A1","text":"n"},""" +
        """{"op":"remove-comment","sheet":"Other","ref":"A1"}]"""
    val result = run(wb, Some(data), json).fold(e => fail(e.getMessage), identity)
    assertEquals(sheet(result, "Other").comments.get(a1), None)
    assertEquals(sheet(result, "Data"), data)
  }

  test("ungroup-rows/ungroup-cols with sheet=Other operate there") {
    val json =
      """[{"op":"group-rows","sheet":"Other","rows":"1:2"},{"op":"ungroup-rows","sheet":"Other","rows":"1:2"},""" +
        """{"op":"group-cols","sheet":"Other","cols":"A:B"},{"op":"ungroup-cols","sheet":"Other","cols":"A:B"}]"""
    val result = run(book, Some(data), json).fold(e => fail(e.getMessage), identity)
    assertEquals(sheet(result, "Data"), data)
  }

  test("the op's sheet applies without any batch default") {
    val json = """[{"op":"put","sheet":"Other","ref":"C1","value":"scoped"}]"""
    val result = run(book, None, json).fold(e => fail(e.getMessage), identity)
    assertEquals(
      sheet(result, "Other").cells.get(ARef.from0(2, 0)).map(_.value),
      Some(CellValue.Text("scoped"))
    )
  }

  test("a qualified ref names the sheet even when the batch default is another sheet") {
    val json = """[{"op":"put","ref":"Other!C1","value":"qualified"}]"""
    val result = run(book, Some(data), json).fold(e => fail(e.getMessage), identity)
    assertEquals(
      sheet(result, "Other").cells.get(ARef.from0(2, 0)).map(_.value),
      Some(CellValue.Text("qualified"))
    )
    assertEquals(sheet(result, "Data"), data)
  }

  test("a qualified ref that agrees with sheet is fine") {
    val json = """[{"op":"put","sheet":"Other","ref":"Other!C1","value":1}]"""
    assert(run(book, Some(data), json).isRight)
  }

  test("a qualified ref that disagrees with sheet is BATCH_OP_INVALID naming the index") {
    val json =
      """[{"op":"put","ref":"A2","value":1},{"op":"put","sheet":"Data","ref":"Other!A1","value":2}]"""
    val error = cliError(BatchParser.parseBatchOperations(json).attempt.unsafeRunSync()).error
    assertEquals(error.code, ErrorCode.BATCH_OP_INVALID)
    assert(error.message.startsWith("Object 2"), error.message)
    assert(error.message.contains("Other") && error.message.contains("Data"), error.message)
    assertEquals(error.location.flatMap(_.opIndex), Some(2))
    assertEquals(error.exitCode, ExitCodes.usage)
  }

  test("a disagreeing qualified range on a range op is BATCH_OP_INVALID too") {
    val json = """[{"op":"style","sheet":"Data","range":"Other!A1:B2","bold":true}]"""
    val error = cliError(BatchParser.parseBatchOperations(json).attempt.unsafeRunSync()).error
    assertEquals(error.code, ErrorCode.BATCH_OP_INVALID)
    assertEquals(error.location.flatMap(_.opIndex), Some(1))
  }

  test("copy may qualify its sides across sheets while sheet supplies the unqualified side") {
    val json = """[{"op":"copy","sheet":"Other","source":"Data!A1","target":"H1"}]"""
    val result = run(book, None, json).fold(e => fail(e.getMessage), identity)
    assertEquals(
      sheet(result, "Other").cells.get(ARef.from0(7, 0)).map(_.value),
      Some(CellValue.Text("keep"))
    )
  }

  test("a missing sheet is BATCH_OP_FAILED with the op index and did-you-mean candidates") {
    val json =
      """[{"op":"put","ref":"A1","value":1},{"op":"put","sheet":"Othr","ref":"A1","value":2}]"""
    val error = cliError(run(book, Some(data), json)).error
    assertEquals(error.code, ErrorCode.BATCH_OP_FAILED)
    assert(error.message.startsWith("Object 2 (put): "), error.message)
    assert(error.message.contains("Othr"), error.message)
    assert(error.candidates.contains("Other"), error.candidates.toString)
    assertEquals(error.location.flatMap(_.opIndex), Some(2))
    assertEquals(error.location.flatMap(_.sheet), Some("Othr"))
    assertEquals(error.exitCode, ExitCodes.failed)
  }

  test("an invalid sheet name is BATCH_OP_INVALID at parse time") {
    val json = """[{"op":"put","sheet":"","ref":"A1","value":1}]"""
    val error = cliError(BatchParser.parseBatchOperations(json).attempt.unsafeRunSync()).error
    assertEquals(error.code, ErrorCode.BATCH_OP_INVALID)
    assert(error.message.startsWith("Object 1"), error.message)
    assert(error.message.contains("sheet"), error.message)
  }

  test("a non-string sheet is BATCH_OP_INVALID at parse time") {
    val json = """[{"op":"put","sheet":7,"ref":"A1","value":1}]"""
    val error = cliError(BatchParser.parseBatchOperations(json).attempt.unsafeRunSync()).error
    assertEquals(error.code, ErrorCode.BATCH_OP_INVALID)
  }

  test("sheet on an op that is not sheet-scoped is an unknown property, not a scope") {
    val json = """[{"op":"add-sheet","sheet":"Data","name":"Third"}]"""
    val result = BatchParser.parseBatchOperations(json).unsafeRunSync()
    assertEquals(result.scoped.map(_.sheet), Vector(None))
    assert(
      result.warnings.exists(w => w.contains("unknown properties ignored") && w.contains("sheet")),
      result.warnings.toString
    )
  }

  test("rename-sheet of the default sheet retargets the ops that follow") {
    val json =
      """[{"op":"put","ref":"A2","value":"before"},""" +
        """{"op":"rename-sheet","from":"Data","to":"Renamed"},""" +
        """{"op":"put","ref":"A3","value":"after"},""" +
        """{"op":"style","range":"A3","bold":true}]"""
    val result = run(book, Some(data), json).fold(e => fail(e.getMessage), identity)
    assert(result.sheets.forall(_.name.value != "Data"), result.sheetNames.toString)
    val renamed = sheet(result, "Renamed")
    assertEquals(renamed.cells.get(ARef.from0(0, 1)).map(_.value), Some(CellValue.Text("before")))
    assertEquals(renamed.cells.get(ARef.from0(0, 2)).map(_.value), Some(CellValue.Text("after")))
    assert(renamed.getCellStyle(ARef.from0(0, 2)).exists(_.font.bold))
  }

  test("rename-sheet of another sheet leaves the default alone") {
    val json =
      """[{"op":"rename-sheet","from":"Other","to":"Elsewhere"},{"op":"put","ref":"A2","value":"still data"}]"""
    val result = run(book, Some(data), json).fold(e => fail(e.getMessage), identity)
    assertEquals(
      sheet(result, "Data").cells.get(ARef.from0(0, 1)).map(_.value),
      Some(CellValue.Text("still data"))
    )
    assert(result.sheets.exists(_.name.value == "Elsewhere"))
  }

  test("an op's own sheet naming the pre-rename default fails after the rename (explicit wins)") {
    val json =
      """[{"op":"rename-sheet","from":"Data","to":"Renamed"},{"op":"put","sheet":"Data","ref":"A3","value":1}]"""
    val error = cliError(run(book, Some(data), json)).error
    assertEquals(error.code, ErrorCode.BATCH_OP_FAILED)
    assertEquals(error.location.flatMap(_.opIndex), Some(2))
  }

  test("applyBatchOperations still applies plain ops against the default sheet") {
    val ops = Vector(BatchOp.Put("A9", CellValue.Text("plain"), None), BatchOp.Merge("B9:C9"))
    val result = BatchParser.applyBatchOperations(book, Some(data), ops).unsafeRunSync()
    assertEquals(
      sheet(result, "Data").cells.get(ARef.from0(0, 8)).map(_.value),
      Some(CellValue.Text("plain"))
    )
    assertEquals(sheet(result, "Other"), other)
  }

  test("applyBatchOperations wraps failures with the 1-based op index") {
    val ops = Vector(BatchOp.Put("A9", CellValue.Text("ok"), None), BatchOp.Merge("nonsense"))
    val error =
      cliError(BatchParser.applyBatchOperations(book, Some(data), ops).attempt.unsafeRunSync())
    assertEquals(error.error.code, ErrorCode.BATCH_OP_FAILED)
    assert(error.getMessage.startsWith("Object 2 (merge): "), error.getMessage)
    assertEquals(error.error.location.flatMap(_.opIndex), Some(2))
  }

  test(
    "an apply-time failure keeps the cause text as a suffix (existing contains-assertions hold)"
  ) {
    val json =
      """[{"op":"put","ref":"A1","value":1},{"op":"rowheight","ref":"A1","row":2,"height":9}]"""
    // rowheight has no default sheet: the historical message must survive inside the wrapper
    val error = cliError(run(book, None, json))
    assert(error.getMessage.startsWith("Object 1 (put): "), error.getMessage)
    assert(error.getMessage.contains("requires --sheet"), error.getMessage)
  }

  test("ParseResult.ops equals scoped.map(_.op) and indices are 1-based in order") {
    val json =
      """[{"op":"put","ref":"A1","value":1},{"op":"merge","sheet":"Other","range":"A1:B1"}]"""
    val result = BatchParser.parseBatchOperations(json).unsafeRunSync()
    assertEquals(result.scoped.map(_.op), result.ops)
    assertEquals(result.scoped.map(_.index), Vector(1, 2))
    assertEquals(result.scoped.map(_.sheet), Vector(None, SheetName("Other").toOption))
  }
