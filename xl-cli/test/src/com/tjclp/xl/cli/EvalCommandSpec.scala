package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
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

  test("GH-715: a transitive precedent's range over a placement outside the used range is caught") {
    // The used range is A1:C10; D40 and F30 lie outside it. Neither the formula nor its direct
    // read reads them: that read's precedent does, through a bounded range or a whole row.
    val sheet = pairedColumns(Sheet("Test"))
      .put(ref"C1", CellValue.Formula("=C2*2"))
      .put(ref"C2", CellValue.Formula("=SUM(D1:D50)"))
      .put(ref"C3", CellValue.Formula("=C4+1"))
      .put(ref"C4", CellValue.Formula("=SUM(30:30)"))
    for
      _ <- circularAt(sheet, "=C1+1", "D40", "D40 → C1 → C2 → D40")
      _ <- circularAt(sheet, "=C3", "F30", "F30 → C3 → C4 → F30")
      // just past the bounded range: no cycle, the plain cell evaluates
      beyond <- ReadCommands.eval(Workbook(sheet), Some(sheet), "=C1+1", Some("D51"), Nil)
    yield assert(beyond.contains("Result: 1 (number)"), beyond)
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
        val expected =
          s"Circular reference detected: the formula reads its own placement $cell at run time"
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

  test(
    "GH-715: a lookup whose result cell is the placement reports the cycle, not an error value"
  ) {
    // C5 = 7 keys the table C1:D9 (a defined name, out of the static check's sight); the matched
    // result cell is D5 itself. A failing result cell shows its error, but the guard's read stays
    // the circular reference.
    val keyed = staleTest.put(ref"C5", CellValue.Number(BigDecimal(7)))
    val book = stale.put(keyed)
    val tabled = book.copy(metadata =
      book.metadata.copy(definedNames = Vector(DefinedName("tbl", "Test!$C$1:$D$9")))
    )
    readsPlacement(tabled, "=VLOOKUP(7,tbl,2,FALSE)", "D5", "D5") >>
      readsPlacement(tabled, "=IFERROR(VLOOKUP(7,tbl,2,FALSE),0)", "D5", "D5")
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

  test(
    "GH-715: --with reaches a cached formula however it is read, with --at, without it, in evala"
  ) {
    // Data!B1 = A1*2 cached as 10 (A1 = 5); with A1 = 200 it is 400 whichever way it is read:
    // unqualified, qualified with its own sheet, both at once, or back from another sheet
    val data = Sheet("Data")
      .put(ref"A1", CellValue.Number(BigDecimal(5)))
      .put(ref"B1", CellValue.Formula("=A1*2", Some(CellValue.Number(BigDecimal(10)))))
    val other = Sheet("Other").put(ref"A1", CellValue.Formula("=Data!B1*2"))
    val wb = Workbook(Vector(data, other))
    val overrides = List("A1=200")
    // `=INDIRECT("B1")` and `=Other!A1` reach B1 only at run time or through another sheet: the
    // shapes that make --with uncache stale readers sheet-wide, not just in the closure
    val cases = List(
      "=B1" -> 400,
      "=Data!B1" -> 400,
      "=B1+Data!B1" -> 800,
      "=Other!A1" -> 800,
      "=INDIRECT(\"B1\")" -> 400
    )
    val checks = for
      (formula, expected) <- cases
      at <- List(Some("D5"), None)
    yield ReadCommands.eval(wb, Some(data), formula, at, overrides).map { result =>
      assert(result.contains(s"Result: $expected (number)"), s"$formula at $at: $result")
    }
    checks.foldLeft(IO.unit)(_ >> _) >>
      ReadCommands.evalArray(wb, Some(data), "=Data!B1+Other!A1", None, overrides).map { array =>
        assert(array.contains("1200"), array)
      }
  }

  test("GH-715: --with over a running total finds stale readers without expanding their ranges") {
    // B{n} = SUM($A$1:A{n}) for 5,000 rows, each cached: every range is distinct, so indexing the
    // readers by cell would cost about 12.5 million entries. They are indexed by range.
    val rows = 5000
    val sheet = (1 to rows).foldLeft(Sheet("Test")) { (s, n) =>
      val cached = CellValue.Number(BigDecimal(n.toLong * (n + 1) / 2))
      s.put(ARef.from0(0, n - 1), CellValue.Number(BigDecimal(n)))
        .put(ARef.from0(1, n - 1), CellValue.Formula(s"=SUM($$A$$1:A$n)", Some(cached)))
    }
    val wb = Workbook(Vector(sheet))
    // A1 goes from 1 to 5: every running total rises by 4
    val expected = rows.toLong * (rows + 1) / 2 + 4
    for
      start <- IO.monotonic
      positionless <- ReadCommands.eval(wb, Some(sheet), s"=B$rows", None, List("A1=5"))
      at <- ReadCommands.eval(wb, Some(sheet), s"=INDIRECT(\"B$rows\")", Some("D1"), List("A1=5"))
      end <- IO.monotonic
    yield
      assert(positionless.contains(s"Result: $expected (number)"), positionless)
      assert(at.contains(s"Result: $expected (number)"), at)
      // generous: a cell-by-cell index took tens of seconds here; the range index takes well under
      assert((end - start).toSeconds < 20, s"took ${(end - start).toMillis} ms")
  }

  /** B{n} = SUM($A$1:A{n}) and C{n} = SUM($B$1:B{n}) over A{n} = n, every formula cached. */
  private def runningTotals(rows: Int): Sheet =
    (1 to rows).foldLeft(Sheet("Test")) { (s, n) =>
      val b = CellValue.Number(BigDecimal(n.toLong * (n + 1) / 2))
      val c = CellValue.Number(BigDecimal(n.toLong * (n + 1) * (n + 2) / 6))
      s.put(ARef.from0(0, n - 1), CellValue.Number(BigDecimal(n)))
        .put(ARef.from0(1, n - 1), CellValue.Formula(s"=SUM($$A$$1:A$n)", Some(b)))
        .put(ARef.from0(2, n - 1), CellValue.Formula(s"=SUM($$B$$1:B$n)", Some(c)))
    }

  test("GH-715: --with reaches a running total of running totals") {
    // A1 goes from 1 to 5: every B rises by 4, so C{n} rises by 4n
    val rows = 60
    val sheet = runningTotals(rows)
    val expected = rows.toLong * (rows + 1) * (rows + 2) / 6 + 4L * rows
    for result <- ReadCommands.eval(
        Workbook(sheet),
        Some(sheet),
        s"=INDIRECT(\"C$rows\")",
        Some("E1"),
        List("A1=5")
      )
    yield assert(result.contains(s"Result: $expected (number)"), result)
  }

  test("GH-715: --with over a running total of running totals stays near-linear to find") {
    // Every B goes stale, and each C range overlaps every B block up to its row, so asking each
    // stale B for its candidate readers afresh cost about rows²/2 checks (26 s here). A reader
    // leaves the index once stale, so each is checked a bounded number of times. The formula
    // reads one B and one C at run time, so the evaluation itself stays cheap.
    val rows = 20000
    val sheet = runningTotals(rows)
    // B{rows} rises by 4; C2 = B1 + B2 rises by 8, from 4
    val expected = rows.toLong * (rows + 1) / 2 + 4 + 12
    for
      start <- IO.monotonic
      result <- ReadCommands.eval(
        Workbook(sheet),
        Some(sheet),
        s"=INDIRECT(\"B$rows\")+INDIRECT(\"C2\")",
        Some("E1"),
        List("A1=5")
      )
      end <- IO.monotonic
    yield
      assert(result.contains(s"Result: $expected (number)"), result)
      assert((end - start).toSeconds < 20, s"took ${(end - start).toMillis} ms")
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

  test("GH-715: a range read through implicit intersection still counts every cell: deliberate") {
    // Pinned over-report: the plain cell =B:B+1 at D5 reads only B5, but the static check counts
    // the whole column as read, so B10 = D5 closes a cycle. A false cycle is safer than a stale
    // value; without B10's read-back, the same formula is B5 + 1.
    val base = pairedColumns(Sheet("Test"))
    val readBack = base.put(ref"B10", CellValue.Formula("=D5"))
    for
      plain <- ReadCommands.eval(Workbook(base), Some(base), "=B:B+1", Some("D5"), Nil)
      _ <- circularAt(readBack, "=B:B+1", "D5", "D5 → B10 → D5")
      _ <- circularAt(readBack, "=B1:B10+1", "D5", "D5 → B10 → D5")
    yield assert(plain.contains("Result: 6 (number)"), plain)
  }

  test("GH-715 (#742): XLOOKUP's return cell reaching the placement at run time") {
    // keys = A1:A10 and rets = C1:C10 (defined names, out of the static check's sight); the match
    // A5 = 5 returns C5 = INDIRECT("D5"), a run-time read of the placement. XLOOKUP returns a
    // reference (GH-713) whose cell is read as a reference reads it, so the guard sees the read,
    // as it does for VLOOKUP over the same cells.
    val returning = staleTest.put(ref"C5", CellValue.Formula("=INDIRECT(\"D5\")"))
    val book = stale.put(returning)
    val named = book.copy(metadata =
      book.metadata.copy(definedNames =
        Vector(
          DefinedName("keys", "Test!$A$1:$A$10"),
          DefinedName("rets", "Test!$C$1:$C$10"),
          DefinedName("tbl", "Test!$A$1:$C$10")
        )
      )
    )
    for
      _ <- readsPlacement(named, "=VLOOKUP(5,tbl,3,FALSE)", "D5", "D5")
      _ <- readsPlacement(named, "=XLOOKUP(5,keys,rets)", "D5", "D5")
    yield ()
  }

  test("GH-715: --at outside the used range with --with on a sheet with no cells") {
    // Other has no cells until --with puts A1 there, so its used range is A1 alone and Z100 lies
    // outside it: the placement still bounds the whole-column reads, and the override still applies
    val wb = Workbook(Vector(pairedColumns(Sheet("Test")), Sheet("Other")))
    val test = wb.sheets.headOption
    def at(formula: String): IO[String] =
      ReadCommands.eval(wb, test, formula, Some("Other!Z100"), List("A1=5"))
    def circular(formula: String, cycle: String): IO[Unit] =
      at(formula).attempt.map {
        case Left(e: XLException) =>
          assert(e.getMessage.contains(s"Circular reference detected: $cycle"), e.getMessage)
        case other => fail(s"$formula: expected a circular reference, got $other")
      }
    for
      scaled <- at("=A1*2")
      column <- at("=SUM(A:A)")
      _ <- circular("=Z100+A1", "Z100 → Z100")
      _ <- circular("=SUM(Z:Z)+A1", "Z100 → Z100")
    yield
      assert(scaled.contains("Result: 10 (number)"), scaled)
      assert(scaled.contains("At: Other!Z100"), scaled)
      assert(column.contains("Result: 5 (number)"), column)
  }

  /** A cached formula: `text` holding `cached`, as Excel saves every formula. */
  private def cachedFormula(text: String, cached: Int): CellValue =
    CellValue.Formula(text, Some(CellValue.Number(BigDecimal(cached))))

  /**
   * Data: A1 = 5, B1 = A1*2 (10), C1 = Other!A1+1 (21). Other: A1 = Data!B1*2 (20), A2 = Data!A2
   * holding a cache of 99 no override reaches, A3 = B1 (Other's own B1, empty) cached 7. Third: A1
   * = Other!A1+1 (21). Every formula holds its cache, as in a book Excel saved.
   */
  private val cachedReadBacks: Workbook = Workbook(
    Vector(
      Sheet("Data")
        .put(ref"A1", CellValue.Number(BigDecimal(5)))
        .put(ref"B1", cachedFormula("=A1*2", 10))
        .put(ref"C1", cachedFormula("=Other!A1+1", 21)),
      Sheet("Other")
        .put(ref"A1", cachedFormula("=Data!B1*2", 20))
        .put(ref"A2", cachedFormula("=Data!A2", 99))
        .put(ref"A3", cachedFormula("=B1", 7)),
      Sheet("Third").put(ref"A1", cachedFormula("=Other!A1+1", 21))
    )
  )

  test("GH-715: --with reaches a CACHED read-back on another sheet, transitively across sheets") {
    // A1 = 200: B1 = 400, Other!A1 = 800, and back on Data C1 = 801, as does Third!A1. Each read
    // is a cache saved before the override, so each must be dropped, not read.
    val data = cachedReadBacks.sheets.headOption
    val overrides = List("A1=200")
    val cases = List(
      "=Other!A1" -> 800,
      "=C1" -> 801,
      "=Third!A1" -> 801,
      "=INDIRECT(\"Other!A1\")" -> 800,
      // untouched: Other!A2 reads a cell no override sets, and Other!A3 reads Other's own B1
      "=Other!A2" -> 99,
      "=Other!A3" -> 7
    )
    val checks = for
      (formula, expected) <- cases
      at <- List(Some("D5"), None)
    yield ReadCommands.eval(cachedReadBacks, data, formula, at, overrides).map { result =>
      assert(result.contains(s"Result: $expected (number)"), s"$formula at $at: $result")
    }
    checks.foldLeft(IO.unit)(_ >> _) >>
      ReadCommands.evalArray(cachedReadBacks, data, "=Other!A1+Third!A1", None, overrides).map {
        array => assert(array.contains("1601"), array)
      }
  }

  /**
   * Data: A1 = 77, B1 = 10, C1 = rate+1 (11). Other: A1 = rate*2 (20), A2 = base*3 (30), A3 =
   * loc+100 (110), A4 = other (a stale 99 no override reaches), A5 = ubase (a stale 42: ubase is
   * `$B$1` unqualified, so from Other it reads Other!B1, empty). Names: rate = Data!$B$1, base =
   * rate, other = Data!$A$1, ubase = $B$1, and loc = Data!$B$1 scoped to Other.
   */
  private val namedReadBacks: Workbook =
    val book = Workbook(
      Vector(
        Sheet("Data")
          .put(ref"A1", CellValue.Number(BigDecimal(77)))
          .put(ref"B1", CellValue.Number(BigDecimal(10)))
          .put(ref"C1", cachedFormula("=rate+1", 11)),
        Sheet("Other")
          .put(ref"A1", cachedFormula("=rate*2", 20))
          .put(ref"A2", cachedFormula("=base*3", 30))
          .put(ref"A3", cachedFormula("=loc+100", 110))
          .put(ref"A4", cachedFormula("=other", 99))
          .put(ref"A5", cachedFormula("=ubase", 42))
      )
    )
    book.copy(metadata =
      book.metadata.copy(definedNames =
        Vector(
          DefinedName("rate", "Data!$B$1"),
          DefinedName("base", "rate"),
          DefinedName("other", "Data!$A$1"),
          DefinedName("ubase", "$B$1"),
          DefinedName("loc", "Data!$B$1", localSheetId = Some(1))
        )
      )
    )

  test("GH-715: --with reaches cached read-backs through defined names, on any sheet") {
    // B1 = 5: rate = 5, so Other!A1 = 10, A2 = 15 (through base), A3 = 105 (a name scoped to
    // Other), and Data!C1 = 6, read at run time so the closure fold does not compute it
    val data = namedReadBacks.sheets.headOption
    val overrides = List("B1=5")
    val cases = List(
      "=Other!A1" -> 10,
      "=Other!A2" -> 15,
      "=Other!A3" -> 105,
      "=INDIRECT(\"C1\")" -> 6,
      // untouched: other names Data!A1, and ubase resolves from Other to Other!B1
      "=Other!A4" -> 99,
      "=Other!A5" -> 42
    )
    val checks = for
      (formula, expected) <- cases
      at <- List(Some("D5"), None)
    yield ReadCommands.eval(namedReadBacks, data, formula, at, overrides).map { result =>
      assert(result.contains(s"Result: $expected (number)"), s"$formula at $at: $result")
    }
    checks.foldLeft(IO.unit)(_ >> _) >>
      ReadCommands.evalArray(namedReadBacks, data, "=Other!A1+Other!A2", None, overrides).map {
        array => assert(array.contains("25"), array)
      }
  }

  /**
   * GH-761: Data: B1 = 10, B2 = 1. Other: A2 = SUMIF(rng,">0") (11), A3 = COUNTIF(rng,">5") (1), A4 =
   * SUM(rng) (11), each cached. Names: rng = Data!$B$1:$B$2, read in a range slot (A2, A3) and as a
   * value (A4).
   */
  private val rangeSlotReadBacks: Workbook =
    val book = Workbook(
      Vector(
        Sheet("Data")
          .put(ref"B1", CellValue.Number(BigDecimal(10)))
          .put(ref"B2", CellValue.Number(BigDecimal(1))),
        Sheet("Other")
          .put(ref"A2", cachedFormula("=SUMIF(rng,\">0\")", 11))
          .put(ref"A3", cachedFormula("=COUNTIF(rng,\">5\")", 1))
          .put(ref"A4", cachedFormula("=SUM(rng)", 11))
      )
    )
    book.copy(metadata =
      book.metadata.copy(definedNames = Vector(DefinedName("rng", "Data!$B$1:$B$2")))
    )

  test("GH-761: --with reaches a cached reader through a defined name in a range slot") {
    // B1 = 0: SUMIF over rng is 1, COUNTIF(rng,">5") is 0, SUM(rng) is 1. Before the fix the
    // range-slot readers kept their caches (11 and 1) while the value-position SUM did not.
    val data = rangeSlotReadBacks.sheets.headOption
    val overrides = List("B1=0")
    val cases = List("=Other!A2" -> 1, "=Other!A3" -> 0, "=Other!A4" -> 1)
    val checks = for
      (formula, expected) <- cases
      at <- List(Some("D5"), None)
    yield ReadCommands.eval(rangeSlotReadBacks, data, formula, at, overrides).map { result =>
      assert(result.contains(s"Result: $expected (number)"), s"$formula at $at: $result")
    }
    checks.foldLeft(IO.unit)(_ >> _)
  }

  test("GH-761 CLI: eval --with sees past a range-slot name's cache in a saved file") {
    for
      book <- written(rangeSlotReadBacks)
      sumif <- CliHarness.run("-f", book, "-s", "Data", "eval", "=Other!A2", "--with", "B1=0")
      countif <- CliHarness.run("-f", book, "-s", "Data", "eval", "=Other!A3", "--with", "B1=0")
    yield
      assertEquals(sumif.exit, 0, sumif.stderr)
      assert(sumif.stdout.contains("Result: 1 (number)"), sumif.stdout)
      assertEquals(countif.exit, 0, countif.stderr)
      assert(countif.stdout.contains("Result: 0 (number)"), countif.stdout)
  }

  /**
   * GH-761: Data: B1 = 10, B2 = 1, C1 = 5, C2 = 7. Names: rng = Data!$B$1:$B$2, cost = Data!$C$1.
   * Other, each cached: A1 = SUMIF(rng,">0",Data!C1) (12), A2 = SUMIF(Data!B1:B2,">0",cost) (12),
   * A3 = AVERAGEIF(rng,">0",Data!C1) (6): a sum range sized by a name, or a name sized, to C1:C2.
   */
  private val sizedReadBacks: Workbook =
    val book = Workbook(
      Vector(
        Sheet("Data")
          .put(ref"B1", CellValue.Number(BigDecimal(10)))
          .put(ref"B2", CellValue.Number(BigDecimal(1)))
          .put(ref"C1", CellValue.Number(BigDecimal(5)))
          .put(ref"C2", CellValue.Number(BigDecimal(7))),
        Sheet("Other")
          .put(ref"A1", cachedFormula("=SUMIF(rng,\">0\",Data!C1)", 12))
          .put(ref"A2", cachedFormula("=SUMIF(Data!B1:B2,\">0\",cost)", 12))
          .put(ref"A3", cachedFormula("=AVERAGEIF(rng,\">0\",Data!C1)", 6))
      )
    )
    book.copy(metadata =
      book.metadata.copy(definedNames =
        Vector(DefinedName("rng", "Data!$B$1:$B$2"), DefinedName("cost", "Data!$C$1"))
      )
    )

  test("GH-761: --with reaches a cached reader whose sum range a defined name sizes") {
    // C2 = 0: each reader sums or averages C1:C2 (5 and 0), as Excel sizes the sum range to the
    // criteria range. Before, C2 was no read of theirs and the caches (12, 12, 6) survived
    val data = sizedReadBacks.sheets.headOption
    val cases = List("=Other!A1" -> "5", "=Other!A2" -> "5", "=Other!A3" -> "2.5")
    val checks = for
      (formula, expected) <- cases
      at <- List(Some("D5"), None)
    yield ReadCommands.eval(sizedReadBacks, data, formula, at, List("C2=0")).map { result =>
      assert(result.contains(s"Result: $expected (number)"), s"$formula at $at: $result")
    }
    checks.foldLeft(IO.unit)(_ >> _)
  }

  test("GH-715: evala --with sees another sheet's cached read-back without its cache") {
    // evala goes through the same stale-reader search as eval: Other!A1 is 800, not its 20
    val data = cachedReadBacks.sheets.headOption
    for array <- ReadCommands.evalArray(
        cachedReadBacks,
        data,
        "=Other!A1*{1,2}",
        None,
        List("A1=200")
      )
    yield
      assert(array.contains("800"), array)
      assert(array.contains("1600"), array)
      assert(!array.contains("20 "), array)
  }

  test("GH-715: --at inside another anchor's spill range answers as the plain cell there") {
    // Pinned: C1 holds SEQUENCE(3) spilling over C1:C3 (cached). Excel would show #SPILL! in C1
    // once C2 holds a formula; --at answers what C2 itself would show and does not model that.
    val spilled = Sheet("Data")
      .put(
        ref"C1",
        CellValue.Formula(
          "SEQUENCE(3)",
          Some(CellValue.Number(BigDecimal(1))),
          FormulaKind.ArrayFormula(CellRange(ref"C1", ref"C3"))
        )
      )
      .put(ref"C2", CellValue.Number(BigDecimal(2)))
      .put(ref"C3", CellValue.Number(BigDecimal(3)))
    val wb = Workbook(Vector(spilled))
    for
      inside <- ReadCommands.eval(wb, Some(spilled), "=7", Some("C2"), Nil)
      outside <- ReadCommands.eval(wb, Some(spilled), "=SUM(C1#)", Some("E5"), Nil)
      readsSpill <- ReadCommands.eval(wb, Some(spilled), "=SUM(C1#)", Some("C2"), Nil).attempt
    yield
      assert(inside.contains("Result: 7 (number)"), inside)
      assert(outside.contains("Result: 6 (number)"), outside)
      // reading the spill it sits in answers from C1's cached spill (Excel: #SPILL! and a cycle)
      readsSpill match
        case Right(result) => assert(result.contains("Result: 6 (number)"), result)
        case Left(e) => fail(s"pinned: =SUM(C1#) at C2 evaluates, got $e")
  }

  test("GH-715 CLI: --with reaches another sheet's cached read-back in a saved file") {
    for
      book <- written(cachedReadBacks)
      result <- CliHarness.run(
        "-f",
        book,
        "-s",
        "Data",
        "--json",
        "eval",
        "=Other!A1",
        "--with",
        "A1=200"
      )
    yield
      assertEquals(result.exit, 0, result.stdout)
      assertEquals(
        ujson.read(result.stdout)("data")("result")("value"),
        ujson.Num(800),
        result.stdout
      )
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

  // ===== #750: whole-axis arrays =====

  private val axisBook: Workbook =
    Workbook(
      Vector(
        Sheet("Test")
          .put(ref"A1", CellValue.Number(BigDecimal(1)))
          .put(ref"A2", CellValue.Text("x"))
      )
    )

  test("#750: eval of whole-axis arrays answers at once, with Excel's values") {
    val sheet = axisBook.sheets.head
    val started = System.nanoTime()
    for
      top <- ReadCommands.eval(axisBook, Some(sheet), "=R:KHI", None, Nil)
      whole <- ReadCommands.eval(axisBook, Some(sheet), "=A:XFD", None, Nil)
      blanks <- ReadCommands.eval(axisBook, Some(sheet), "=SUM((A:B=\"\")*1)", None, Nil)
      rows <- ReadCommands.eval(axisBook, Some(sheet), "=ROWS(A:A*1)", None, Nil)
      byColumn <- ReadCommands.eval(axisBook, Some(sheet), "=SUM((A:A=\"\")*{1,2,3})", None, Nil)
    yield
      val seconds = (System.nanoTime() - started) / 1e9
      assert(seconds < 30.0, s"took $seconds s")
      // a positionless whole-axis range is its top-left value, as `=A1:B3` is
      assert(top.contains("Result: 0 (number)"), top)
      assert(whole.contains("Result: 1 (number)"), whole)
      // 2 x 1,048,576 cells, two of them filled
      assert(blanks.contains("2097150"), blanks)
      assert(rows.contains("1048576"), rows)
      // a whole column times a row of constants: 1,048,574 blank rows × (1 + 2 + 3)
      assert(byColumn.contains("6291444"), byColumn)
  }

  test("#750: evala refuses a whole-axis spill past the grid or the cell limit, as #SPILL!") {
    val sheet = axisBook.sheets.head
    for
      // from the default anchor Z1000 a full column runs past row 1,048,576
      offGrid <- ReadCommands.evalArray(axisBook, Some(sheet), "=A:A*1", None, Nil).attempt
      // two full columns from A1 fit the grid but hold 2,097,152 cells
      tooLarge <- ReadCommands.evalArray(axisBook, Some(sheet), "=A:B*1", Some("A1"), Nil).attempt
    yield
      offGrid match
        case Left(e) =>
          assert(e.getMessage.contains("runs past the sheet's last cell"), e.getMessage)
        case Right(out) => fail(s"expected an off-grid refusal, got: $out")
      tooLarge match
        case Left(e) => assert(e.getMessage.contains("2,097,152 cells"), e.getMessage)
        case Right(out) => fail(s"expected a too-large refusal, got: $out")
  }

  // ===== #763: --with reaches readers whose references are computed at run time =====

  /**
   * #763: Data: B1 = 10, B2 = 1. Name: dyn = OFFSET(Data!$B$1,0,0,2,1). Other, each cached 11 (12
   * for A6): A1 = SUM(dyn), A2 = SUMIF(dyn,">0"), A3 = SUM(OFFSET(Data!B1,0,0,2,1)), A4 =
   * SUM(INDIRECT("Data!B1:B2")), A5 = SUM(Data!B1:B2) (the static control), A6 = A3+1.
   */
  private val dynamicReadBacks: Workbook =
    val book = Workbook(
      Vector(
        Sheet("Data")
          .put(ref"B1", CellValue.Number(BigDecimal(10)))
          .put(ref"B2", CellValue.Number(BigDecimal(1))),
        Sheet("Other")
          .put(ref"A1", cachedFormula("=SUM(dyn)", 11))
          .put(ref"A2", cachedFormula("=SUMIF(dyn,\">0\")", 11))
          .put(ref"A3", cachedFormula("=SUM(OFFSET(Data!B1,0,0,2,1))", 11))
          .put(ref"A4", cachedFormula("=SUM(INDIRECT(\"Data!B1:B2\"))", 11))
          .put(ref"A5", cachedFormula("=SUM(Data!B1:B2)", 11))
          .put(ref"A6", cachedFormula("=A3+1", 12))
      )
    )
    book.copy(metadata =
      book.metadata.copy(definedNames = Vector(DefinedName("dyn", "OFFSET(Data!$B$1,0,0,2,1)")))
    )

  test("#763: --with reaches a cached reader through OFFSET or INDIRECT, direct or by a name") {
    // B2 = 0: each reader sums B1:B2 to 10, and A6 = A3+1 is 11. Before, only the static control
    // dropped its cache; the computed references read B1 alone to the static search
    val data = dynamicReadBacks.sheets.headOption
    val cases = List(
      "=Other!A1" -> 10,
      "=Other!A2" -> 10,
      "=Other!A3" -> 10,
      "=Other!A4" -> 10,
      "=Other!A5" -> 10,
      "=Other!A6" -> 11
    )
    val checks = for
      (formula, expected) <- cases
      at <- List(Some("D5"), None)
    yield ReadCommands.eval(dynamicReadBacks, data, formula, at, List("B2=0")).map { result =>
      assert(result.contains(s"Result: $expected (number)"), s"$formula at $at: $result")
    }
    checks.foldLeft(IO.unit)(_ >> _) >>
      ReadCommands
        .evalArray(dynamicReadBacks, data, "=Other!A1+Other!A4", None, List("B2=0"))
        .map(array => assert(array.contains("20"), array))
  }

  test("#763 CLI: eval --with sees past a dynamic reader's cache in a saved file") {
    for
      book <- written(dynamicReadBacks)
      named <- CliHarness.run("-f", book, "-s", "Data", "eval", "=Other!A1", "--with", "B2=0")
      indirect <- CliHarness.run("-f", book, "-s", "Data", "eval", "=Other!A4", "--with", "B2=0")
    yield
      assertEquals(named.exit, 0, named.stderr)
      assert(named.stdout.contains("Result: 10 (number)"), named.stdout)
      assertEquals(indirect.exit, 0, indirect.stderr)
      assert(indirect.stdout.contains("Result: 10 (number)"), indirect.stdout)
  }

  // ===== #727: evala --at validates its anchor =====

  test("#727: evala --at a bad anchor is INVALID_REFERENCE, not the Z1000 fallback") {
    def refused(at: String): IO[Unit] =
      ReadCommands
        .evalArray(pairedBook, pairedBook.sheets.headOption, "=A1:A3*2", Some(at), Nil)
        .attempt
        .map {
          case Left(e: CliException) => assertEquals(e.error.code, "INVALID_REFERENCE", at)
          case other => fail(s"evala --at $at: expected INVALID_REFERENCE, got $other")
        }
    // `$C$2`: a ref argument is written without `$` anchors, for --at as for every verb
    refused("ZZZZ9") >> refused("not a ref") >> refused("D5:D6") >> refused("$C$2") >>
      refused("'My Sheet'!$C$2")
  }

  test("#727: a qualified evala --at anchors on, and evaluates against, the sheet it names") {
    val other = Sheet("Other")
      .put(ref"A1", CellValue.Number(BigDecimal(7)))
      .put(ref"A2", CellValue.Number(BigDecimal(8)))
    val wb = Workbook(Vector(pairedColumns(Sheet("Test")), other))
    for
      qualified <- ReadCommands.evalArrayData(
        wb,
        wb.sheets.headOption,
        "=A1:A2*2",
        Some("Other!C2"),
        Nil
      )
      plain <- ReadCommands.evalArrayData(wb, wb.sheets.headOption, "=A1:A2*2", Some("C2"), Nil)
    yield
      val q = ujson.read(qualified)
      assertEquals(q("spillRange"), ujson.Str("C2:C3"), qualified)
      assertEquals(q("result")("sheet"), ujson.Str("Other"), qualified)
      assert(qualified.contains("14") && qualified.contains("16"), qualified)
      val p = ujson.read(plain)
      assertEquals(p("result")("sheet"), ujson.Str("Test"), plain)
  }

  test("#727 CLI: evala --at ZZZZ9 exits 3; a qualified anchor is honoured") {
    for
      book <- written(pairedBook)
      bad <- CliHarness.run("-f", book, "--json", "evala", "=A1:A3*2", "--at", "ZZZZ9")
      good <- CliHarness.run("-f", book, "evala", "=A1:A3*2", "--at", "Test!C2")
    yield
      assertEquals(bad.exit, 3, bad.stdout)
      assertEquals(ujson.read(bad.stdout)("error")("code"), ujson.Str("INVALID_REFERENCE"))
      assertEquals(good.exit, 0, good.stderr)
      assert(good.stdout.contains("C2:C4"), good.stdout)
  }
