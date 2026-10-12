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
import com.tjclp.xl.workbooks.DefinedName

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

  // --- GH-715 review: same-sheet qualified reads, empty sheets, long cycles, columns ------------

  /** `eval --at` on `wb` (default sheet `sheet`) must refuse, naming the cycle. */
  private def circularIn(
    wb: Workbook,
    sheet: Option[Sheet],
    formula: String,
    at: String,
    cycle: String
  ): IO[Unit] =
    ReadCommands.eval(wb, sheet, formula, Some(at), Nil).attempt.map {
      case Left(e: XLException) =>
        assert(e.getMessage.contains(s"Circular reference detected: $cycle"), e.getMessage)
      case other => fail(s"$formula --at $at: expected a circular reference, got $other")
    }

  /** Test (A1:B10 paired, D5 holding a stale 7) and Other (D5 holding 9). */
  private val stale: Workbook = Workbook(
    Vector(
      pairedColumns(Sheet("Test")).put(ref"D5", CellValue.Number(BigDecimal(7))),
      Sheet("Other").put(ref"D5", CellValue.Number(BigDecimal(9)))
    )
  )

  private val staleTest: Sheet = stale.sheets.head
  private val staleOther: Sheet = stale.sheets(1)

  test("GH-715: a read qualified with the placement's own sheet is circular, not its old value") {
    val test = Some(staleTest)
    circularIn(stale, test, "=Test!D5", "Test!D5", "D5 → D5") >>
      circularIn(stale, test, "=Test!D5+1", "D5", "D5 → D5") >>
      circularIn(stale, Some(staleOther), "=Test!D5", "Test!D5", "D5 → D5") >>
      circularIn(stale, test, "='Test'!D5*2", "D5", "D5 → D5") >>
      circularIn(stale, test, "=test!D5", "D5", "D5 → D5") >>
      circularIn(stale, test, "=SUM(Test!D1:D9)", "Test!D5", "D5 → D5") >>
      circularIn(stale, test, "=SUM(Test!D:D)", "Test!D20", "D20 → D20")
  }

  test("GH-715: a precedent's read qualified with the placement's sheet closes the cycle too") {
    val direct = staleTest.put(ref"C1", CellValue.Formula("=Test!D5*2"))
    val chained = staleTest
      .put(ref"C1", CellValue.Formula("=Test!B20+1"))
      .put(ref"B20", CellValue.Formula("=D5*2"))
    circularIn(stale.put(direct), Some(direct), "=C1+1", "D5", "D5 → C1 → D5") >>
      circularIn(stale.put(chained), Some(chained), "=C1+1", "D5", "D5 → C1 → B20 → D5")
  }

  test("GH-715: another sheet's same address, and inspecting the own-sheet address, evaluate") {
    for
      other <- ReadCommands.eval(stale, Some(staleTest), "=Other!D5+1", Some("D5"), Nil)
      row <- ReadCommands.eval(stale, Some(staleTest), "=ROW(Test!D5)", Some("D5"), Nil)
    yield
      assert(other.contains("Result: 10 (number)"), other)
      assert(row.contains("Result: 5 (number)"), row)
  }

  /** `eval --at` must refuse: the run-time guard saw a read of the placement `cell`. */
  private def readsPlacement(wb: Workbook, formula: String, at: String, cell: String): IO[Unit] =
    ReadCommands.eval(wb, wb.sheets.headOption, formula, Some(at), Nil).attempt.map {
      case Left(e: XLException) =>
        val expected = s"Circular reference detected: the formula reads its own placement $cell"
        assert(e.getMessage.contains(expected), e.getMessage)
      case other => fail(s"$formula --at $at: expected a circular reference, got $other")
    }

  test("GH-715 (#729): a cycle through another sheet's uncached formula is circular") {
    // Other!A1 reads Test!D5. The static check follows no other sheet; the run-time guard sees
    // Other!A1's evaluation read the placement.
    val wb = stale.put(staleOther.put(ref"A1", CellValue.Formula("=Test!D5*2")))
    readsPlacement(wb, "=Other!A1", "D5", "D5")
  }

  test("GH-715 known gap (#729): another sheet's CACHED formula that reads the placement") {
    // A cached formula on another sheet is read as its cache, never re-evaluated, so the guard
    // never sees its read of Test!D5: the answer is that cache (here 14).
    val cached = CellValue.Formula("=Test!D5*2", Some(CellValue.Number(BigDecimal(14))))
    val wb = stale.put(staleOther.put(ref"A1", cached))
    for result <- ReadCommands.eval(wb, Some(staleTest), "=Other!A1", Some("D5"), Nil)
    yield assert(result.contains("Result: 14 (number)"), result)
  }

  test("GH-715: --at on an empty sheet does not enumerate the formula's whole-column ranges") {
    // Without a used range to bound them, A:XFC would expand to 17 billion cells.
    val wb = Workbook(Vector(pairedColumns(Sheet("Test")), Sheet("Other")))
    for
      at <- ReadCommands.eval(wb, wb.sheets.headOption, "=SUM(A:XFC)", Some("Other!XFD5"), Nil)
      positionless <- ReadCommands.eval(wb, wb.sheets.lift(1), "=SUM(A:XFC)", None, Nil)
    yield
      assert(at.contains("Result: 0 (number)"), at)
      assert(positionless.contains("Result: 0 (number)"), positionless)
  }

  /** Test with C1 = C2, C2 = C3, …, C`length` = D5. */
  private def chainTo(length: Int): Sheet =
    (1 to length).foldLeft(Sheet("Test")) { (s, row) =>
      val next = if row == length then "=D5" else s"=C${row + 1}"
      s.put(ARef.from0(2, row - 1), CellValue.Formula(next))
    }

  test("GH-715: a long cycle path is elided in the middle; a short one is printed whole") {
    val huge = chainTo(5000)
    for
      _ <- circularAt(chainTo(40), "=C1", "D5", "D5 → C1 → C2 → … (37 more) … → C40 → D5")
      _ <- circularAt(chainTo(7), "=C1", "D5", "D5 → C1 → C2 → C3 → C4 → C5 → C6 → C7 → D5")
      message <- ReadCommands.eval(Workbook(huge), Some(huge), "=C1", Some("D5"), Nil).attempt
    yield message match
      case Left(e: XLException) =>
        assert(e.getMessage.contains("… (4,997 more) …"), e.getMessage)
        assert(e.getMessage.length < 300, e.getMessage)
      case other => fail(s"expected a circular reference, got $other")
  }

  test("GH-715: a cycle elsewhere on the sheet does not block --at") {
    // Only the formula's own closure is ordered: F1 and F2 read each other but nothing reads them.
    val sheet = pairedColumns(Sheet("Test"))
      .put(ref"F1", CellValue.Formula("=F2"))
      .put(ref"F2", CellValue.Formula("=F1"))
      .put(ref"C1", CellValue.Formula("=A5*10"))
    for result <- ReadCommands.eval(Workbook(sheet), Some(sheet), "=C1+1", Some("D5"), Nil)
    yield assert(result.contains("Result: 51 (number)"), result)
  }

  test("GH-715: --at intersects a row range with the placement's column") {
    // A1:J1 holds 1..10 and A2:J2 holds 10..100; the plain cell at E5 reads E1 and E2.
    val rows = (0 until 10).foldLeft(Sheet("Test")) { (s, c) =>
      s.put(ARef.from0(c, 0), CellValue.Number(BigDecimal(c + 1)))
        .put(ARef.from0(c, 1), CellValue.Number(BigDecimal(10 * (c + 1))))
    }
    val wb = Workbook(Vector(rows))
    val formula = "=SUM(A1:J1*A2:J2)"
    for
      positionless <- ReadCommands.eval(wb, Some(rows), formula, None, Nil)
      at <- ReadCommands.eval(wb, Some(rows), formula, Some("E5"), Nil)
      outside <- ReadCommands.eval(wb, Some(rows), formula, Some("L5"), Nil)
    yield
      assert(positionless.contains("Result: 3850 (number)"), positionless)
      assert(at.contains("Result: 250 (number)"), at)
      assert(outside.contains("Result: #VALUE! (error)"), outside)
  }

  test("GH-715 CLI: a same-sheet qualified self-read is FORMULA_ERROR; --stream refuses --at") {
    for
      book <- written(stale)
      self <- CliHarness.run("-f", book, "-s", "Test", "--json", "eval", "=Test!D5", "--at", "D5")
      streamed <- CliHarness.run(
        "-f",
        book,
        "-s",
        "Test",
        "--json",
        "--stream",
        "eval",
        sumProduct,
        "--at",
        "D5"
      )
    yield
      assertEquals(self.exit, 3, self.stdout)
      val error = ujson.read(self.stdout)("error")
      assertEquals(error("code"), ujson.Str("FORMULA_ERROR"), self.stdout)
      assert(error("message").str.contains("Circular reference detected: D5 → D5"), self.stdout)
      assertEquals(streamed.exit, 2, streamed.stdout)
      assertEquals(ujson.read(streamed.stdout)("error")("code"), ujson.Str("UNSUPPORTED_IN_STREAM"))
  }

  test("GH-715: a formula already at the --at cell is replaced: neither a node nor evaluated") {
    // D5's own formula reads itself and the other one does not parse. Evaluating either, or
    // ordering it with the closure, would fail; the plain cell at D5 is the new formula alone.
    val base = pairedColumns(Sheet("Test")).put(ref"C1", CellValue.Formula("=A5*2"))
    val circular = base.put(ref"D5", CellValue.Formula("=D5+1"))
    val broken = base.put(ref"D5", CellValue.Formula("=SUM(1,"))
    for
      overCircular <- ReadCommands.eval(
        Workbook(circular),
        Some(circular),
        "=C1+1",
        Some("D5"),
        Nil
      )
      overBroken <- ReadCommands.eval(Workbook(broken), Some(broken), "=C1+1", Some("D5"), Nil)
    yield
      assert(overCircular.contains("Result: 11 (number)"), overCircular)
      assert(overBroken.contains("Result: 11 (number)"), overBroken)
  }

  test("GH-715 (#729): a reference computed at run time that lands on the placement is circular") {
    // The static check cannot see these reads; the run-time guard does, as Excel's would.
    val named = stale.copy(metadata =
      stale.metadata.copy(definedNames = Vector(DefinedName("placed", "Test!$D$5")))
    )
    val dynamicPrecedent = staleTest.put(ref"C1", CellValue.Formula("=INDIRECT(\"D5\")*2"))
    readsPlacement(stale, "=INDIRECT(\"D5\")+1", "D5", "D5") >>
      readsPlacement(stale, "=INDIRECT(\"Test!D5\")", "D5", "D5") >>
      readsPlacement(stale, "=OFFSET(A1,4,0)", "A5", "A5") >>
      readsPlacement(stale, "=SUM(OFFSET(A1,0,0,10,1))", "A5", "A5") >>
      readsPlacement(stale, "=IFERROR(INDIRECT(\"D5\"),0)", "D5", "D5") >>
      readsPlacement(named, "=placed+1", "D5", "D5") >>
      // a precedent whose run-time read lands on the placement fails while the closure evaluates
      readsPlacement(stale.put(dynamicPrecedent), "=C1+1", "D5", "D5")
  }

  test("GH-715: every reader family that reaches the placement trips the guard, under IFERROR") {
    // Match-style readers (COUNTIF, SUMIF, MATCH, the lookups) evaluate an uncached cell
    // themselves rather than through a cell reference; each must still see the guard. A range slot
    // takes no INDIRECT, so defined names hide the ranges from the static check: col is D1:D9 and
    // block D1:E9, both holding the placement D5.
    val named = stale.copy(metadata =
      stale.metadata.copy(definedNames =
        Vector(DefinedName("col", "Test!$D$1:$D$9"), DefinedName("block", "Test!$D$1:$E$9"))
      )
    )
    List(
      "=IFERROR(COUNTIF(col,1),0)",
      "=IFERROR(COUNTIFS(col,1),0)",
      "=IFERROR(SUMIF(col,\">0\"),0)",
      "=IFERROR(AVERAGEIF(col,\">0\"),0)",
      "=IFERROR(MATCH(7,col,0),0)",
      "=IFERROR(VLOOKUP(7,block,2,FALSE),0)",
      "=IFERROR(XLOOKUP(7,col,col),0)",
      "=IFERROR(SUMPRODUCT(col),0)",
      "=IFERROR(SUM(col),0)",
      "=IFERROR(COUNTA(col),0)",
      "=IFERROR(COUNTBLANK(col),0)",
      "=IFERROR(SUM(INDIRECT(\"D1:D9\")),0)"
    ).foldLeft(IO.unit)((acc, formula) => acc >> readsPlacement(named, formula, "D5", "D5"))
  }

  test("GH-715: run-time references elsewhere evaluate; INDEX over the placement is static") {
    val selector = staleTest.put(ref"C1", CellValue.Number(BigDecimal(5)))
    for
      beside <- ReadCommands.eval(stale, Some(staleTest), "=INDIRECT(\"A5\")+1", Some("D5"), Nil)
      offset <- ReadCommands.eval(stale, Some(staleTest), "=OFFSET(A1,4,0)*2", Some("D5"), Nil)
      // INDEX reads its whole array as values, so a computed selector over the placement is a
      // cycle the static check already names
      _ <- circularIn(stale.put(selector), Some(selector), "=INDEX(A1:A10,C1)", "A5", "A5 → A5")
    yield
      assert(beside.contains("Result: 6 (number)"), beside)
      assert(offset.contains("Result: 10 (number)"), offset)
  }

  test("GH-715: a LET body that returns its bound placement reads it") {
    circularAt(staleTest, "=LET(x,D5,x)", "D5", "D5 → D5") >>
      circularAt(staleTest, "=LET(x,Test!D5,x+0)", "D5", "D5 → D5")
  }

  test("GH-715: a --with qualifier names the evaluated sheet in any case, as Excel's names do") {
    // Intended: `test!D5` names Test, as sheet names compare in Excel. So the override of the
    // placement is the USAGE refusal, and a differently cased override elsewhere applies.
    for
      clash <- ReadCommands
        .eval(stale, Some(staleTest), "=A5+1", Some("Test!D5"), List("test!D5=3"))
        .attempt
      beside <- ReadCommands.eval(stale, Some(staleTest), "=A5+1", Some("D5"), List("TEST!A5=3"))
      positionless <- ReadCommands.eval(stale, Some(staleTest), "=A5+1", None, List("test!A5=3"))
    yield
      clash match
        case Left(e: CliException) => assertEquals(e.error.code, ErrorCode.USAGE)
        case other => fail(s"expected USAGE, got $other")
      assert(beside.contains("Result: 4 (number)"), beside)
      assert(positionless.contains("Result: 4 (number)"), positionless)
  }

  test("GH-715: --at far below the used range, read through a whole column, is circular") {
    // the used range ends at D10; D5000 lies far outside it, yet SUM(D:D) there reads it
    val column = (1 to 10).foldLeft(Sheet("Test")) { (s, row) =>
      s.put(ARef.from0(3, row - 1), CellValue.Number(BigDecimal(row)))
    }
    circularAt(column, "=SUM(D:D)", "D5000", "D5000 → D5000")
  }

  test("GH-715: --at evaluates a deep precedent chain, each step through the placement guard") {
    // C1 = C2+1, …, C2999 = C3000+1, C3000 = 1: a 3,000-formula closure the fold evaluates in
    // order, every read under the guard, with no memo held across the fold's steps
    val depth = 3000
    val chain = (1 to depth).foldLeft(Sheet("Test")) { (s, row) =>
      val formula = if row == depth then "=1" else s"=C${row + 1}+1"
      s.put(ARef.from0(2, row - 1), CellValue.Formula(formula))
    }
    for result <- ReadCommands.eval(Workbook(chain), Some(chain), "=C1", Some("D5"), Nil)
    yield assert(result.contains(s"Result: $depth (number)"), result)
  }

  test("GH-715: --at a cell on a sheet with no cells, read by the formula, is circular") {
    val empty = Sheet("Empty")
    circularIn(Workbook(Vector(empty)), Some(empty), "=D5", "D5", "D5 → D5") >>
      circularIn(Workbook(Vector(empty)), Some(empty), "=SUM(D:D)", "D5", "D5 → D5")
  }

  test(
    "GH-715: --with reaches another sheet's read-back alike with --at, without it, and in evala"
  ) {
    // Other!A1 reads Data!B1 back; the override of B1 must reach it in every mode.
    val data = Sheet("Data").put(ref"B1", CellValue.Number(BigDecimal(5)))
    val other = Sheet("Other").put(ref"A1", CellValue.Formula("=Data!B1*2"))
    val wb = Workbook(Vector(data, other))
    val overrides = List("B1=10")
    for
      at <- ReadCommands.eval(wb, Some(data), "=Other!A1", Some("D5"), overrides)
      positionless <- ReadCommands.eval(wb, Some(data), "=Other!A1+0*B1", None, overrides)
      qualifiedOnly <- ReadCommands.eval(wb, Some(data), "=Other!A1", None, overrides)
      array <- ReadCommands.evalArray(wb, Some(data), "=Other!A1+0*B1", None, overrides)
    yield
      assert(at.contains("Result: 20 (number)"), at)
      assert(positionless.contains("Result: 20 (number)"), positionless)
      assert(qualifiedOnly.contains("Result: 20 (number)"), qualifiedOnly)
      assert(array.contains("20"), array)
  }

  test("GH-715: --with on the --at cell itself is USAGE: the formula is that cell") {
    def refused(at: String, overrides: List[String]): IO[Unit] =
      ReadCommands.eval(stale, Some(staleTest), "=A5+1", Some(at), overrides).attempt.map {
        case Left(e: CliException) =>
          assertEquals(e.error.code, ErrorCode.USAGE)
          assert(e.error.message.contains("--with") && e.error.message.contains("--at"), e.error)
        case other => fail(s"--at $at --with $overrides: expected USAGE, got $other")
      }
    for
      _ <- refused("D5", List("D5=3"))
      _ <- refused("Test!D5", List("A1=2", "Test!D5=3"))
      _ <- refused("D5", List(" d5 = 3"))
      beside <- ReadCommands.eval(stale, Some(staleTest), "=A5+1", Some("D5"), List("A5=3"))
      book <- written(stale)
      cli <- CliHarness.run(
        "-f",
        book,
        "-s",
        "Test",
        "--json",
        "eval",
        "=A5",
        "--at",
        "D5",
        "-w",
        "D5=3"
      )
    yield
      assert(beside.contains("Result: 4 (number)"), beside)
      assertEquals(cli.exit, 2, cli.stdout)
      assertEquals(ujson.read(cli.stdout)("error")("code"), ujson.Str("USAGE"), cli.stdout)
  }
