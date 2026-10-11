package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.ReadCommands
import com.tjclp.xl.cli.contract.{CliException, CliHarness, ErrorCode}
import com.tjclp.xl.error.XLException
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * Tests for eval command, particularly --with override propagation (TJC-698).
 *
 * The eval command with --with flag should propagate overrides through formula chains. For example,
 * if A1=100, B1=A1*2, C1=B1+50, then `eval "=C1" --with "A1=200"` should return 450 (not 250).
 */
// Test code uses .get/.head for brevity in assertions
@SuppressWarnings(
  Array(
    "org.wartremover.warts.OptionPartial",
    "org.wartremover.warts.IterableOps"
  )
)
class EvalCommandSpec extends CatsEffectSuite:

  // Create a test workbook with formula chain: A1=100, B1=A1*2, C1=B1+50
  private def createFormulaChainWorkbook: IO[Workbook] =
    IO {
      val sheet = Sheet("Test")
        .put(ref"A1", CellValue.Number(BigDecimal(100)))
        .put(ref"B1", CellValue.Formula("=A1*2"))
        .put(ref"C1", CellValue.Formula("=B1+50"))

      Workbook(Vector(sheet))
    }

  test("eval: formula without overrides returns correct value") {
    for
      wb <- createFormulaChainWorkbook
      sheet = wb.sheets.head
      result <- ReadCommands.eval(wb, Some(sheet), "=C1", None, Nil)
    yield
      // C1 = B1+50 = (A1*2)+50 = 100*2+50 = 250
      assert(result.contains("250"), s"Expected 250, got: $result")
  }

  test("eval: override propagates through single formula (TJC-698)") {
    for
      wb <- createFormulaChainWorkbook
      sheet = wb.sheets.head
      result <- ReadCommands.eval(wb, Some(sheet), "=B1", None, List("A1=200"))
    yield
      // B1 = A1*2 = 200*2 = 400
      assert(result.contains("400"), s"Expected 400 (A1=200 → B1=400), got: $result")
  }

  test("eval: override propagates through formula chain (TJC-698)") {
    for
      wb <- createFormulaChainWorkbook
      sheet = wb.sheets.head
      result <- ReadCommands.eval(wb, Some(sheet), "=C1", None, List("A1=200"))
    yield
      // C1 = B1+50 = (A1*2)+50 = 200*2+50 = 450
      // Bug before fix: returned 250 (B1 wasn't recalculated with new A1 value)
      assert(result.contains("450"), s"Expected 450 (A1=200 → B1=400 → C1=450), got: $result")
  }

  test("eval: multiple overrides propagate correctly") {
    for
      wb <- IO {
        val sheet = Sheet("Test")
          .put(ref"A1", CellValue.Number(BigDecimal(10)))
          .put(ref"A2", CellValue.Number(BigDecimal(20)))
          .put(ref"B1", CellValue.Formula("=A1+A2"))

        Workbook(Vector(sheet))
      }
      sheet = wb.sheets.head
      result <- ReadCommands.eval(wb, Some(sheet), "=B1", None, List("A1=100", "A2=200"))
    yield
      // B1 = A1+A2 = 100+200 = 300
      assert(result.contains("300"), s"Expected 300, got: $result")
  }

  test("eval: override with no dependent formulas works") {
    for
      wb <- createFormulaChainWorkbook
      sheet = wb.sheets.head
      result <- ReadCommands.eval(wb, Some(sheet), "=A1*10", None, List("A1=50"))
    yield
      // Direct formula using overridden value: 50*10 = 500
      assert(result.contains("500"), s"Expected 500, got: $result")
  }

  test("eval: constant formula without overrides doesn't need file") {
    for result <- ReadCommands.eval(Workbook.empty, None, "=1+1", None, Nil)
    yield assert(result.contains("2"), s"Expected 2, got: $result")
  }

  test("eval: an unparseable constant formula reads as the diagnostic, not a case class") {
    ReadCommands.eval(Workbook.empty, None, "=FOOBAR(1)", None, Nil).attempt.map {
      case Left(e) =>
        assert(e.getMessage.contains("Unknown function 'FOOBAR' at position 0"), e.getMessage)
        assert(e.getMessage.contains("Did you mean: FLOOR?"), e.getMessage)
        assert(!e.getMessage.contains("UnknownFunction("), e.getMessage)
      case Right(out) => fail(s"expected a parse failure, got $out")
    }
  }

  test("eval: only evaluates dependency closure (optimization)") {
    // Sheet with many formulas, but query only needs a few
    for
      wb <- IO {
        val sheet = Sheet("Test")
          .put(ref"A1", CellValue.Number(BigDecimal(100)))
          .put(ref"B1", CellValue.Formula("=A1*2")) // In closure
          .put(ref"C1", CellValue.Formula("=B1+50")) // In closure (target)
          .put(ref"D1", CellValue.Formula("=999")) // NOT in closure
          .put(ref"E1", CellValue.Formula("=D1*2")) // NOT in closure
        Workbook(Vector(sheet))
      }
      sheet = wb.sheets.head
      result <- ReadCommands.eval(wb, Some(sheet), "=C1", None, List("A1=200"))
    yield
      // C1 should correctly evaluate to 450 (A1=200 → B1=400 → C1=450)
      // D1 and E1 should NOT be evaluated (optimization - but result is same)
      assert(result.contains("450"), s"Expected 450, got: $result")
  }

  test("GH-344: eval renders an error VALUE as Excel shows it (no exception)") {
    for
      wb <- createFormulaChainWorkbook
      sheet = wb.sheets.head
      result <- ReadCommands.eval(wb, Some(sheet), "=1/0", None, Nil)
    yield assert(result.contains("Result: #DIV/0! (error)"), s"got: $result")
  }

  test("GH-344: eval over an error-valued precedent no longer aborts the closure walk") {
    for
      wb <- IO {
        val sheet = Sheet("Test")
          .put(ref"X1", CellValue.Formula("=1/0")) // precedent computes #DIV/0!
          .put(ref"Y1", CellValue.Formula("=X1+1")) // dependent reads it
        Workbook(Vector(sheet))
      }
      sheet = wb.sheets.head
      result <- ReadCommands.eval(wb, Some(sheet), "=Y1", None, Nil)
      caught <- ReadCommands.eval(wb, Some(sheet), "=IFERROR(Y1,42)", None, Nil)
    yield
      assert(result.contains("Result: #DIV/0! (error)"), s"got: $result")
      assert(caught.contains("42"), s"got: $caught")
  }

  // --- GH-715: eval --at <ref> evaluates the formula as the plain cell it would be at <ref> ------

  /** `sheet` with A1:A10 and B1:B10 both holding 1..10. */
  private def pairedColumns(sheet: Sheet): Sheet =
    (0 until 10).foldLeft(sheet) { (s, i) =>
      s.put(ARef.from0(0, i), CellValue.Number(BigDecimal(i + 1)))
        .put(ARef.from0(1, i), CellValue.Number(BigDecimal(i + 1)))
    }

  private val pairedBook: Workbook = Workbook(Vector(pairedColumns(Sheet("Test"))))

  private val sumProduct = "=SUM(A1:A10*B1:B10)"

  private def evalPaired(at: Option[String], overrides: List[String] = Nil): IO[String] =
    ReadCommands.eval(pairedBook, pairedBook.sheets.headOption, sumProduct, at, overrides)

  test("GH-715: eval without --at is positionless: the array sum") {
    for result <- evalPaired(None)
    yield
      assert(result.contains("Result: 385 (number)"), result)
      assert(!result.contains("At:"), result)
  }

  test("GH-715: eval --at D5 evaluates the plain cell at D5: A5*B5") {
    for result <- evalPaired(Some("D5"))
    yield
      assert(result.contains("Result: 25 (number)"), result)
      assert(result.contains("At: Test!D5"), result)
  }

  test("GH-715: eval --at a row the ranges do not span is #VALUE!, as the plain cell is") {
    for result <- evalPaired(Some("D20"))
    yield assert(result.contains("Result: #VALUE! (error)"), result)
  }

  test("GH-715: eval --at composes with --with") {
    for result <- evalPaired(Some("D5"), List("A5=100"))
    yield
      assert(result.contains("Result: 500 (number)"), result)
      assert(result.contains("With: A5=100"), result)
  }

  test("GH-715: eval --at threads the precedent closure before the positioned formula") {
    val sheet = pairedColumns(Sheet("Test")).put(ref"C5", CellValue.Formula("=A5*10"))
    val wb = Workbook(Vector(sheet))
    for result <- ReadCommands.eval(wb, Some(sheet), "=C1:C10+1", Some("E5"), List("A5=7"))
    yield assert(result.contains("Result: 71 (number)"), result)
  }

  test("GH-715: a qualified --at names the sheet the formula evaluates on, over -s") {
    val other = Sheet("Other").put(ref"A5", CellValue.Number(BigDecimal(42)))
    val wb = Workbook(Vector(pairedColumns(Sheet("Test")), other))
    for result <- ReadCommands.eval(wb, wb.sheets.headOption, "=A1:A10", Some("Other!D5"), Nil)
    yield
      assert(result.contains("Result: 42 (number)"), result)
      assert(result.contains("At: Other!D5"), result)
  }

  test("GH-715: an unqualified --at on a multi-sheet book with no sheet is SHEET_REQUIRED") {
    val wb = Workbook(Vector(pairedColumns(Sheet("Test")), Sheet("Other")))
    ReadCommands.eval(wb, None, sumProduct, Some("D5"), Nil).attempt.map {
      case Left(e: CliException) => assertEquals(e.error.code, "SHEET_REQUIRED")
      case other => fail(s"expected SHEET_REQUIRED, got $other")
    }
  }

  test("GH-715: --at must be one cell: a range or a malformed ref is INVALID_REFERENCE") {
    def refused(at: String): IO[Unit] =
      evalPaired(Some(at)).attempt.map {
        case Left(e: CliException) => assertEquals(e.error.code, "INVALID_REFERENCE")
        case other => fail(s"--at $at: expected INVALID_REFERENCE, got $other")
      }
    refused("D5:D6") >> refused("not a ref")
  }

  test("GH-715: eval --at needs no file for a formula without cell references") {
    val noFile = Workbook(Vector.empty) // what the CLI evaluates against without -f
    for
      row <- ReadCommands.eval(noFile, None, "=ROW()*100+COLUMN()", Some("D5"), Nil)
      qualified <- ReadCommands.eval(noFile, None, "=1", Some("Data!D5"), Nil).attempt
      reads <- ReadCommands.eval(noFile, None, "=A1", Some("D5"), Nil).attempt
    yield
      assert(row.contains("Result: 504 (number)"), row)
      assert(row.contains("At: D5"), row)
      qualified match
        case Left(e: CliException) => assertEquals(e.error.code, ErrorCode.USAGE)
        case other => fail(s"a qualified --at with no file: expected USAGE, got $other")
      reads match
        case Left(e: CliException) => assertEquals(e.error.code, ErrorCode.USAGE)
        case other => fail(s"a cell-reading formula with no file: expected USAGE, got $other")
  }

  test("GH-715: eval --json carries the position (null when positionless)") {
    def data(at: Option[String]): IO[ujson.Value] =
      ReadCommands
        .evalData(pairedBook, pairedBook.sheets.headOption, sumProduct, at, Nil)
        .map(ujson.read(_))
    for
      at <- data(Some("D5"))
      none <- data(None)
    yield
      assertEquals(at("at"), ujson.Str("Test!D5"))
      assertEquals(at("result")("value"), ujson.Num(25))
      assertEquals(none("at"), ujson.Null)
      assertEquals(none("result")("value"), ujson.Num(385))
  }

  /** `eval --at` on `sheet` must refuse, naming the cycle the plain cell would close. */
  private def circularAt(sheet: Sheet, formula: String, at: String, cycle: String): IO[Unit] =
    ReadCommands.eval(Workbook(Vector(sheet)), Some(sheet), formula, Some(at), Nil).attempt.map {
      case Left(e: XLException) =>
        assert(e.getMessage.contains(s"Circular reference detected: $cycle"), e.getMessage)
      case other => fail(s"$formula --at $at: expected a circular reference, got $other")
    }

  test("GH-715: eval --at a cell the formula reads is a circular reference, not its old value") {
    val sheet = pairedColumns(Sheet("Test"))
    circularAt(sheet, "=SUM(A1:A10)", "A5", "A5 → A5") >>
      circularAt(sheet, "=B3+1", "B3", "B3 → B3") >>
      circularAt(sheet, sumProduct, "B3", "B3 → B3")
  }

  test("GH-715: eval --at a cell its precedents read back is a circular reference") {
    val sheet = pairedColumns(Sheet("Test")).put(ref"C1", CellValue.Formula("=D5*2"))
    circularAt(sheet, "=C1+1", "D5", "D5 → C1 → D5")
  }

  test("GH-715: a cycle through a range is caught outside the used range too") {
    val sheet = pairedColumns(Sheet("Test")).put(ref"C1", CellValue.Formula("=SUM(D:D)"))
    circularAt(sheet, "=SUM(A:A)", "A20", "A20 → A20") >>
      circularAt(sheet, "=C1", "D20", "D20 → C1 → D20")
  }

  test("GH-715: eval --at a cell the formula does not read evaluates") {
    val sheet = pairedColumns(Sheet("Test")).put(ref"C1", CellValue.Formula("=D5*2"))
    val wb = Workbook(Vector(sheet))
    for
      below <- ReadCommands.eval(wb, Some(sheet), "=SUM(A1:A10)", Some("A11"), Nil)
      beside <- ReadCommands.eval(wb, Some(sheet), "=C1+1", Some("D6"), Nil)
    yield
      assert(below.contains("Result: 55 (number)"), below)
      assert(beside.contains("Result: 1 (number)"), beside)
  }

  // --- GH-715 end to end, through the real command tree ----------------------------------------

  private def written(wb: Workbook): IO[String] =
    for
      path <- IO.blocking(Files.createTempFile("xl-gh715-", ".xlsx"))
      _ <- ExcelIO.instance[IO].write(wb, path)
    yield path.toString

  private def warningCodes(stdout: String): Vector[String] =
    ujson.read(stdout)("warnings").arr.map(_("code").str).toVector

  test("GH-715 CLI: eval --at answers what the cell would show, read-only, with --with") {
    for
      book <- written(pairedBook)
      positionless <- CliHarness.run("-f", book, "eval", sumProduct)
      at <- CliHarness.run("-f", book, "eval", sumProduct, "--at", "D5")
      outside <- CliHarness.run("-f", book, "eval", sumProduct, "--at", "D20")
      overridden <- CliHarness.run("-f", book, "eval", sumProduct, "--at", "D5", "-w", "B5=3")
      reread <- ExcelIO.instance[IO].read(Path.of(book))
    yield
      assertEquals(positionless.exit, 0, positionless.stderr)
      assert(positionless.stdout.contains("Result: 385 (number)"), positionless.stdout)
      assertEquals(at.exit, 0, at.stderr)
      assert(at.stdout.contains("Result: 25 (number)"), at.stdout)
      assert(at.stdout.contains("At: Test!D5"), at.stdout)
      assert(outside.stdout.contains("Result: #VALUE! (error)"), outside.stdout)
      assert(overridden.stdout.contains("Result: 15 (number)"), overridden.stdout)
      assert(reread.sheets.forall(_.cells.get(ref"D5").isEmpty), "eval --at writes nothing")
  }

  test("GH-715 CLI: --json carries the position; a qualified --at needs no auto-select") {
    for
      book <- written(pairedBook)
      auto <- CliHarness.run("-f", book, "--json", "eval", sumProduct, "--at", "D5")
      qualified <- CliHarness.run("-f", book, "--json", "eval", sumProduct, "--at", "Test!D5")
    yield
      List(auto, qualified).foreach { run =>
        assertEquals(run.exit, 0, run.stderr)
        val data = ujson.read(run.stdout)("data")
        assertEquals(data("at"), ujson.Str("Test!D5"), run.stdout)
        assertEquals(data("result")("value"), ujson.Num(25), run.stdout)
      }
      assertEquals(warningCodes(auto.stdout), Vector("SHEET_AUTOSELECTED"), auto.stdout)
      assertEquals(warningCodes(qualified.stdout), Vector.empty, qualified.stdout)
  }

  test("GH-715 CLI: --at follows THE sheet rule on a multi-sheet book") {
    val other = (0 until 10).foldLeft(Sheet("Other")) { (s, i) =>
      s.put(ARef.from0(0, i), CellValue.Number(BigDecimal(100)))
        .put(ARef.from0(1, i), CellValue.Number(BigDecimal(2)))
    }
    for
      book <- written(Workbook(Vector(pairedColumns(Sheet("Test")), other)))
      required <- CliHarness.run("-f", book, "--json", "eval", sumProduct, "--at", "D5")
      flagged <- CliHarness.run("-f", book, "-s", "Test", "eval", sumProduct, "--at", "D5")
      qualified <- CliHarness.run("-f", book, "-s", "Test", "eval", sumProduct, "--at", "Other!D5")
    yield
      assertEquals(required.exit, 3, required.stdout)
      assertEquals(ujson.read(required.stdout)("error")("code"), ujson.Str("SHEET_REQUIRED"))
      assert(flagged.stdout.contains("Result: 25 (number)"), flagged.stdout)
      assert(qualified.stdout.contains("Result: 200 (number)"), qualified.stdout)
      assert(qualified.stdout.contains("At: Other!D5"), qualified.stdout)
  }

  test("GH-715 CLI: a bad --at is INVALID_REFERENCE; --stream refuses eval --at; help names it") {
    for
      book <- written(pairedBook)
      range <- CliHarness.run("-f", book, "--json", "eval", sumProduct, "--at", "D5:D6")
      garbage <- CliHarness.run("-f", book, "--json", "eval", sumProduct, "--at", "nope")
      streamed <- CliHarness.run("-f", book, "--stream", "eval", sumProduct, "--at", "D5")
      help <- CliHarness.run("eval", "--help")
      noFile <- CliHarness.run("eval", "=ROW()", "--at", "D5")
    yield
      assertEquals(noFile.exit, 0, noFile.stderr)
      assert(noFile.stdout.contains("Result: 5 (number)"), noFile.stdout)
      List(range, garbage).foreach { run =>
        assertEquals(run.exit, 3, run.stdout)
        assertEquals(ujson.read(run.stdout)("error")("code"), ujson.Str("INVALID_REFERENCE"))
      }
      assertEquals(streamed.exit, 2, streamed.stderr)
      assert(streamed.stderr.contains("eval is not supported with --stream"), streamed.stderr)
      assert(help.stdout.contains("--at"), help.stdout)
  }
