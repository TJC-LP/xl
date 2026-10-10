package com.tjclp.xl.formula

import com.tjclp.xl.XLResult
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.formula.eval.{IterativeCalc, IterativeMode, RecalcOptions}
import com.tjclp.xl.formula.eval.SheetEvaluator.*
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-655: `x#` — Excel 365's spill reference, stored as `_xlfn.ANCHORARRAY(x)`. The value is the
 * whole array the anchor cell's formula spills: the file's `<f t="array" ref>` span read through
 * its cached cells when Excel wrote one, else the anchor's formula evaluated as an array. A cell
 * that is not a spill anchor — a constant, a scalar formula, a spilled non-anchor value — is
 * `#REF!`. The parser accepts `x#` and `ANCHORARRAY(x)`; the printer emits `x#`.
 */
class SpillReferenceSpec extends FunSuite:

  private val n = (v: Int) => CellValue.Number(BigDecimal(v))

  /**
   * A1 spills SORT(B1:B3) without a cache (an xl-authored dynamic formula); C1 carries the file's
   * array record for SEQUENCE(3) with cached spill cells; D1 the same for a function xl does not
   * evaluate (SORTBY), so only the cached spill can answer; E1 is a constant, E2 a scalar formula.
   */
  private val sheet = Sheet("Data")
    .put(ref"B1", n(3))
    .put(ref"B2", n(1))
    .put(ref"B3", n(2))
    .put(ref"A1", CellValue.Formula("SORT(B1:B3)"))
    .put(
      ref"C1",
      CellValue.Formula(
        "SEQUENCE(3)",
        Some(n(1)),
        FormulaKind.ArrayFormula(CellRange(ref"C1", ref"C3"))
      )
    )
    .put(ref"C2", n(2))
    .put(ref"C3", n(3))
    .put(
      ref"D1",
      CellValue.Formula(
        "SORTBY(B1:B3,B1:B3)",
        Some(n(10)),
        FormulaKind.ArrayFormula(CellRange(ref"D1", ref"D3"))
      )
    )
    .put(ref"D2", n(20))
    .put(ref"D3", n(30))
    .put(ref"E1", n(5))
    .put(ref"E2", CellValue.Formula("B1*2"))

  private val wb = Workbook(sheet, Sheet("Other").put(ref"A1", n(1)))
    .withDefinedName("spill", "Data!$A$1")

  private def value(formula: String): XLResult[CellValue] =
    sheet.evaluateFormula(formula, Clock.system, Some(wb), Some(ref"H1"))

  /** `formula` as an array formula at H1 (a dynamic-array anchor), its anchor value. */
  private def arrayNumber(formula: String): BigDecimal =
    val placed = sheet.put(
      ref"H1",
      CellValue.Formula(formula, None, FormulaKind.ArrayFormula(CellRange(ref"H1", ref"H1")))
    )
    placed.evaluateCell(ref"H1", Clock.system, Some(wb.put(placed))) match
      case Right(CellValue.Number(v)) => v
      case other => fail(s"$formula: expected a number, got $other")

  private def number(formula: String): BigDecimal =
    value(formula) match
      case Right(CellValue.Number(v)) => v
      case other => fail(s"$formula: expected a number, got $other")

  private def isRef(result: XLResult[CellValue]): Boolean =
    result match
      case Left(err) => err.toString.contains("#REF!") || err.toString.contains("Ref")
      case Right(CellValue.Error(CellError.Ref)) => true
      case _ => false

  test("an uncached dynamic formula's spill is its array evaluation at the anchor") {
    assertEquals(number("=SUM(A1#)"), BigDecimal(6))
    assertEquals(number("=ROWS(A1#)"), BigDecimal(3))
    assertEquals(arrayNumber("=SUM(A1#*10)"), BigDecimal(60), "the spill enters array arithmetic")
    // a plain cell reads an array value under an operator at its top-left, as a legacy formula
    assertEquals(number("=SUM(A1#*10)"), BigDecimal(10))
    assertEquals(number("=A1#"), BigDecimal(1), "a scalar position collapses to the top-left cell")
  }

  test("standalone, x# spills the anchor's array") {
    val (s, range) = sheet.evaluateArrayFormula("=A1#", ref"G1").toOption.get
    assertEquals(range, CellRange(ref"G1", ref"G3"))
    assertEquals(s(ref"G1").value, n(1))
    assertEquals(s(ref"G2").value, n(2))
    assertEquals(s(ref"G3").value, n(3))
  }

  test("the file's array record is Excel's extent, read through the cached spill cells") {
    assertEquals(number("=SUM(C1#)"), BigDecimal(6))
    // a spill from a function xl does not evaluate still reads — the caches carry it
    assertEquals(number("=SUM(D1#)"), BigDecimal(60))
    assertEquals(number("=ROWS(D1#)"), BigDecimal(3))
  }

  test("a constant, a scalar formula and a spilled non-anchor cell are #REF!") {
    assert(isRef(value("=SUM(E1#)")), value("=SUM(E1#)").toString)
    assert(isRef(value("=SUM(E2#)")), value("=SUM(E2#)").toString)
    assert(isRef(value("=SUM(C2#)")), value("=SUM(C2#)").toString)
    assert(isRef(value("=SUM(Z9#)")), value("=SUM(Z9#)").toString)
  }

  test("a sheet-qualified anchor and a defined name bound to the anchor spill the same array") {
    val other = wb.sheets.find(_.name.value == "Other").getOrElse(fail("no Other sheet"))
    assertEquals(
      other.evaluateFormula("=SUM(Data!A1#)", Clock.system, Some(wb), Some(ref"A2")),
      Right(n(6))
    )
    assertEquals(number("=SUM(spill#)"), BigDecimal(6))
    assertEquals(number("=SUM(Data!spill#)"), BigDecimal(6))
  }

  test("the stored spelling ANCHORARRAY(x) evaluates and prints as x#") {
    assertEquals(number("=SUM(ANCHORARRAY(A1))"), BigDecimal(6))
    assertEquals(number("=SUM(_xlfn.ANCHORARRAY(A1))"), BigDecimal(6))
    def printed(f: String): String =
      FormulaParser.parse(f).map(FormulaPrinter.print(_)).fold(e => fail(e.toString), identity)
    assertEquals(printed("=ANCHORARRAY(A1)"), "=A1#")
    assertEquals(printed("=SUM(A1#)"), "=SUM(A1#)")
    assertEquals(printed("=Data!A1#"), "=Data!A1#")
    assertEquals(printed("='My Sheet'!A1#"), "='My Sheet'!A1#")
    assertEquals(printed("=$A$1#"), "=$A$1#")
    assertEquals(printed("=spill#"), "=spill#")
    assertEquals(printed("=A1#*2"), "=A1#*2")
    assertEquals(printed("=@A1#"), "=@A1#")
    assertEquals(printed("=[1]Book!A1#"), "=[1]Book!A1#")
    // a call whose argument is not a shape `#` can follow keeps the call spelling
    assertEquals(printed("=ANCHORARRAY(A1:A3)"), "=ANCHORARRAY(A1:A3)")
    val fileForm = FormulaParser.parse("=SUM(A1#)").fold(e => fail(e.toString), identity)
    assertEquals(FormulaPrinter.printFileForm(fileForm), "SUM(A1#)")
  }

  test("parse ∘ print = id over the spill-reference spellings") {
    List(
      "=SUM(A1#)",
      "=Data!A1#*2",
      "='My Sheet'!$A$1#",
      "=spill#",
      "=@A1#",
      "=A1#%",
      "=INDEX(A1:A3,1)+A1#",
      "=ANCHORARRAY(A1:A3)"
    ).foreach { f =>
      val parsed = FormulaParser.parse(f)
      assert(parsed.isRight, s"$f: $parsed")
      val reprinted = parsed.map(FormulaPrinter.print(_))
      assertEquals(reprinted.flatMap(FormulaParser.parse), parsed, s"$f via $reprinted")
    }
  }

  test("# applies to one cell or name only: a range, a value or a call before it is an error") {
    assert(FormulaParser.parse("=A1:A3#").isLeft)
    assert(FormulaParser.parse("=5#").isLeft)
    assert(FormulaParser.parse("=SUM(A1)#").isLeft)
    assert(FormulaParser.parse("=\"x\"#").isLeft)
    assert(FormulaParser.parse("=#").isLeft)
  }

  test("the anchor's spill is not statically knowable: ANCHORARRAY is a dynamic-dependency call") {
    assert(FunctionSpecs.anchorArray.flags.dynamicDeps)
  }

  test("a spill reference that reaches its own anchor stops at the depth guard, never overflows") {
    val circular = Sheet("Loop").put(ref"A1", CellValue.Formula("ROWS(A1#)"))
    circular.evaluateFormula("=SUM(A1#)", Clock.system, None, Some(ref"B1")) match
      case Left(_) => ()
      case Right(v) => fail(s"a circular spill reference should not evaluate, got $v")
  }

  /**
   * GH-695: Excel's shape — B1 a cached dynamic-array anchor spilling SORT(A1:A5) into B2:B5, C1
   * `SUM(B1#)` cached as 15. A threaded evaluation fold writes B1's computed value back before C1
   * reads it; that write must keep B1 an anchor, else `B1#` sees a constant and is `#REF!`.
   */
  private def excelSpillBook(c1: String): Workbook =
    Workbook(
      Sheet("Sheet1")
        .put(ref"A1", n(5))
        .put(ref"A2", n(3))
        .put(ref"A3", n(1))
        .put(ref"A4", n(4))
        .put(ref"A5", n(2))
        .put(
          ref"B1",
          CellValue.Formula(
            "SORT(A1:A5)",
            Some(n(1)),
            FormulaKind.ArrayFormula(CellRange(ref"B1", ref"B5"))
          )
        )
        .put(ref"B2", n(2))
        .put(ref"B3", n(3))
        .put(ref"B4", n(4))
        .put(ref"B5", n(5))
        .put(ref"C1", CellValue.Formula(c1, Some(n(15))))
        .put(ref"D1", n(0))
    )

  private val spillReaders = List("SUM(B1#)", "SUM(Sheet1!B1#)", "SUM(ANCHORARRAY(B1))")

  private def c1(wb: Workbook): CellValue =
    wb.sheets.headOption.getOrElse(fail("no sheet"))(ref"C1").value

  test("GH-695: a whole-workbook recalculation keeps the spill reader's value") {
    spillReaders.foreach { reader =>
      val book = excelSpillBook(reader)
      val sequential = book.recalculate(Clock.system)
      assertEquals(sequential.errors, Vector.empty, reader)
      assertEquals(c1(sequential.workbook), CellValue.Formula(reader, Some(n(15))), reader)
      val parallel = book.recalculate(RecalcOptions(parallelism = 4))
      assertEquals(c1(parallel.workbook), CellValue.Formula(reader, Some(n(15))), s"$reader ∥")
    }
  }

  test("GH-695: a targeted recalculation that re-evaluates the anchor keeps the reader's value") {
    spillReaders.foreach { reader =>
      // A1 rewritten with its own value: B1 (its dependent) and the reader both re-evaluate
      val result = excelSpillBook(reader)
        .recalculateAfterEdit(SheetName.unsafe("Sheet1"), Set(ref"A1"), RecalcOptions())
      assertEquals(c1(result.workbook), CellValue.Formula(reader, Some(n(15))), reader)
    }
  }

  test("GH-695: the sheet-level dependency folds read the anchor's spill") {
    spillReaders.foreach { reader =>
      val book = excelSpillBook(reader)
      val sheet = book.sheets.head
      assertEquals(
        sheet.evaluateWithDependencyCheck(Clock.system, Some(book)).map(_.get(ref"C1")),
        Right(Some(n(15))),
        reader
      )
      assertEquals(
        sheet.evaluateForRange(CellRange(ref"C1", ref"C1"), Clock.system, Some(book)),
        Right(Map(ref"C1" -> n(15))),
        reader
      )
      assertEquals(
        sheet.evaluateForRangePerCell(CellRange(ref"C1", ref"C1"), Clock.system, Some(book)).values,
        Map(ref"C1" -> n(15)),
        reader
      )
    }
  }

  test("GH-695: a changed input re-spills the anchor for its readers, not its stale spill cells") {
    // A1 5 → 0: SORT gives {0;1;2;3;4}, but B2:B5 still hold the old 2,3,4,5 (xl writes no spill)
    spillReaders.foreach { reader =>
      val book = excelSpillBook(reader)
      val edited = book.put(
        book.sheets.head.put(ref"A1", n(0)).put(ref"D1", CellValue.Formula("B1+0"))
      )
      val full = edited.recalculate(Clock.system).workbook
      assertEquals(c1(full), CellValue.Formula(reader, Some(n(10))), reader)
      assertEquals(
        full.sheets.head(ref"D1").value,
        CellValue.Formula("B1+0", Some(n(0))),
        s"$reader: a plain read of the anchor is its new top-left"
      )
      val targeted =
        edited.recalculateAfterEdit(SheetName.unsafe("Sheet1"), Set(ref"A1"), RecalcOptions())
      assertEquals(c1(targeted.workbook), CellValue.Formula(reader, Some(n(10))), s"$reader →")
    }
  }

  test("GH-695: an uncached array record with no spill cells evaluates its whole array") {
    val book = Workbook(
      Sheet("Sheet1")
        .put(
          ref"A1",
          CellValue
            .Formula("SEQUENCE(3)", None, FormulaKind.ArrayFormula(CellRange(ref"A1", ref"A3")))
        )
        .put(ref"C1", CellValue.Formula("SUM(A1#)"))
    )
    val sum = CellValue.Formula("SUM(A1#)", Some(n(6)))
    assertEquals(c1(book.recalculate(Clock.system).workbook), sum)
    assertEquals(c1(book.recalculate(RecalcOptions(parallelism = 4)).workbook), sum, "∥")
    val sheet = book.sheets.head
    assertEquals(
      sheet.evaluateWithDependencyCheck(Clock.system, Some(book)).map(_.get(ref"C1")),
      Right(Some(n(6)))
    )
  }

  private def arrayRecord(at: ARef, formula: String, cache: Option[CellValue] = None): CellValue =
    CellValue.Formula(formula, cache, FormulaKind.ArrayFormula(CellRange(at, at)))

  test("GH-695: plain and spill readers reuse the anchor's one computation in the generation") {
    val book = Workbook(
      Sheet("Sheet1")
        .put(
          ref"A1",
          CellValue.Formula(
            "SEQUENCE(2,1,RAND(),0)",
            None,
            FormulaKind.ArrayFormula(CellRange(ref"A1", ref"A2"))
          )
        )
        .put(ref"B1", CellValue.Formula("A1+0"))
        .put(ref"C1", CellValue.Formula("A1+0"))
        .put(ref"D1", CellValue.Formula("SUM(A1#)"))
    )
    val sheet = book.recalculate(RecalcOptions(rng = Rng.seeded(42))).workbook.sheets.head
    def cached(at: ARef): BigDecimal = sheet(at).value match
      case CellValue.Formula(_, Some(CellValue.Number(v)), _) => v
      case other => fail(s"${at.toA1}: $other")
    assertEquals(cached(ref"B1"), cached(ref"A1"))
    assertEquals(cached(ref"C1"), cached(ref"A1"))
    assertEquals(cached(ref"D1"), cached(ref"A1") * 2, "the spill is the same draw")
  }

  test("GH-695: a long acyclic chain of array records evaluates in order, never recursively") {
    val chain = (1 to 105).foldLeft(Sheet("Sheet1")) { (s, i) =>
      val at = ARef.from0(0, i - 1)
      s.put(at, arrayRecord(at, if i == 1 then "1" else s"A${i - 1}+1"))
    }
    val result = Workbook(chain).recalculate(Clock.system)
    assertEquals(result.errors, Vector.empty)
    assertEquals(
      result.workbook.sheets.head(ARef.from0(0, 104)).value,
      arrayRecord(ARef.from0(0, 104), "A104+1", Some(n(105)))
    )
  }

  test("GH-695: a pinned external array record keeps its cache for its readers") {
    val book = Workbook(
      Sheet("Sheet1")
        .put(ref"A1", arrayRecord(ref"A1", "[1]Book1!A1", Some(n(7))))
        .put(ref"B1", CellValue.Formula("A1+1"))
    )
    val result = book.recalculate(Clock.system)
    assertEquals(result.errors, Vector.empty)
    assertEquals(c1Of(result.workbook, ref"B1"), CellValue.Formula("A1+1", Some(n(8))))
  }

  test("GH-695: an iterative array member's settled value is what its dependents read") {
    val book = Workbook(
      Sheet("Sheet1")
        .put(ref"A1", arrayRecord(ref"A1", "A1*0.5+1"))
        .put(ref"B1", CellValue.Formula("A1+0"))
    )
    val result = book.recalculate(
      RecalcOptions(iterative = IterativeMode.Force(IterativeCalc(3, BigDecimal("0.001"))))
    )
    assertEquals(result.errors, Vector.empty)
    val settled = c1Of(result.workbook, ref"A1") match
      case CellValue.Formula(_, Some(v), _) => v
      case other => fail(s"A1: $other")
    assertEquals(c1Of(result.workbook, ref"B1"), CellValue.Formula("A1+0", Some(settled)))
  }

  private def c1Of(wb: Workbook, at: ARef): CellValue =
    wb.sheets.headOption.getOrElse(fail("no sheet"))(at).value

  test("GH-695: a function returning a referenced anchor or cached formula stores its value") {
    // IFERROR/IFNA/lookups hand back the referenced cell itself: a threaded anchor's record, or
    // (targeted) a cached precedent the pass did not re-evaluate — the cell stores the value
    List("IFERROR(B1,0)", "IFNA(B1,0)", "VLOOKUP(7,B1:B1,1,FALSE)", "XLOOKUP(7,B1:B1,B1:B1)")
      .foreach { f =>
        val anchored = Workbook(
          Sheet("Sheet1")
            .put(
              ref"B1",
              CellValue.Formula(
                "SEQUENCE(2,1,7)",
                None,
                FormulaKind.ArrayFormula(CellRange(ref"B1", ref"B2"))
              )
            )
            .put(ref"C1", CellValue.Formula(f))
            .put(ref"D1", CellValue.Formula("C1+1"))
        ).recalculate(Clock.system)
        assertEquals(anchored.errors, Vector.empty, f)
        assertEquals(c1Of(anchored.workbook, ref"C1"), CellValue.Formula(f, Some(n(7))), f)
        assertEquals(c1Of(anchored.workbook, ref"D1"), CellValue.Formula("C1+1", Some(n(8))), f)

        val targeted = Workbook(
          Sheet("Sheet1")
            .put(ref"B1", CellValue.Formula("7", Some(n(7))))
            .put(ref"C1", CellValue.Formula(f))
        ).recalculateAfterEdit(SheetName.unsafe("Sheet1"), Set(ref"C1"), RecalcOptions())
        assertEquals(c1Of(targeted.workbook, ref"C1"), CellValue.Formula(f, Some(n(7))), s"$f →")
      }
  }

  test("GH-695: nested IFERROR/IFNA consumers read selected values in every recalculation fold") {
    val cases = List(
      ("1", n(1), "TEXT(%s,\"0\")", CellValue.Text("1")),
      ("TRUE", CellValue.Bool(true), "SUM(%s,1)", n(2)),
      ("\"7\"", CellValue.Text("7"), "SUM(%s,1)", n(8)),
      ("\"bad\"", CellValue.Text("bad"), "SUM(%s,1)", CellValue.Error(CellError.Value)),
      ("\"\"", CellValue.Text(""), "TEXT(%s,\"0\")", CellValue.Text(""))
    )
    for
      guard <- List("IFERROR", "IFNA")
      selected <- List(s"$guard(A1,0)", s"$guard(B1,A1)")
      (source, sourceValue, consumer, expected) <- cases
    do
      val formula = consumer.format(selected)
      val label = s"$source -> $formula"
      val sheet = Sheet("Sheet1")
        .put(ref"A1", CellValue.Formula(source, Some(sourceValue)))
        .put(ref"B1", CellValue.Formula("NA()", Some(CellValue.Error(CellError.NA))))
        .put(ref"C1", CellValue.Formula(formula))
      val book = Workbook(sheet)
      val recalculated = List(
        "full" -> book.recalculate(RecalcOptions()),
        "parallel" -> book.recalculate(RecalcOptions(parallelism = 4)),
        "edited precedent" -> book.recalculateAfterEdit(sheet.name, Set(ref"A1"), RecalcOptions()),
        "cached precedent" -> book.recalculateAfterEdit(sheet.name, Set(ref"C1"), RecalcOptions())
      )
      recalculated.foreach { (mode, result) =>
        assertEquals(result.errors, Vector.empty, s"$label ($mode)")
        assertEquals(
          c1(result.workbook),
          CellValue.Formula(formula, Some(expected)),
          s"$label ($mode)"
        )
      }
      assertEquals(
        sheet.evaluateWithDependencyCheck(Clock.system, Some(book)).map(_.get(ref"C1")),
        Right(Some(expected)),
        s"$label (sheet fold)"
      )
      // The direct scalar entry point can keep a coercion refusal on the Left channel; the
      // dependency folds above promote that nonnumeric SUM case to its cached #VALUE!.
      if expected != CellValue.Error(CellError.Value) then
        assertEquals(sheet.evaluateFormula(s"=$formula"), Right(expected), s"$label (direct cache)")
        val uncached = sheet
          .put(ref"A1", CellValue.Formula(source))
          .put(ref"B1", CellValue.Formula("NA()"))
        assertEquals(
          uncached.evaluateFormula(s"=$formula"),
          Right(expected),
          s"$label (direct uncached)"
        )
  }

  test("GH-695: nested value and reference selectors do not expose threaded formula records") {
    val selectors = List(
      "IFERROR(A1,0)",
      "IFNA(A1,0)",
      "IFERROR(IFNA(A1,0),0)",
      "IF(TRUE,A1,0)",
      "IFS(TRUE,A1)",
      "CHOOSE(1,A1,0)",
      "SWITCH(1,1,A1,0)",
      "INDEX(A1:A2,1)",
      "VLOOKUP(1,A1:A2,1,FALSE)",
      "HLOOKUP(1,A1:B1,1,FALSE)",
      "XLOOKUP(1,A1:A2,A1:A2)",
      "XLOOKUP(3,A1:A2,A1:A2,A1)"
    )
    selectors.foreach { selector =>
      val formula = s"TEXT($selector,\"0\")"
      val book = Workbook(
        Sheet("Sheet1")
          .put(ref"A1", CellValue.Formula("1"))
          .put(ref"A2", n(2))
          .put(ref"B1", n(2))
          .put(ref"C1", CellValue.Formula(formula))
      )
      val result = book.recalculate(RecalcOptions())
      assertEquals(result.errors, Vector.empty, formula)
      assertEquals(
        c1(result.workbook),
        CellValue.Formula(formula, Some(CellValue.Text("1"))),
        formula
      )
    }
  }

  test("GH-695: lookup values feed aggregates without retaining their formula records") {
    // VLOOKUP and HLOOKUP return the value (TRUE, counted as 1); XLOOKUP returns a reference
    // (GH-713), whose logical SUM skips exactly as `SUM(B1,1)` and `SUM(INDEX(B1:B1,1),1)` do in
    // Excel — its if_not_found is a value again
    val selectors = List(
      ("VLOOKUP(1,A1:B1,2,FALSE)", 2),
      ("HLOOKUP(1,A1:A2,2,FALSE)", 2),
      ("XLOOKUP(1,A1:A1,B1:B1)", 1),
      ("INDEX(B1:B1,1)", 1),
      ("B1", 1),
      ("XLOOKUP(2,A1:A1,B1:B1,B1)", 2)
    )
    selectors.foreach { (selector, expected) =>
      val formula = s"SUM($selector,1)"
      val sheet = Sheet("Sheet1")
        .put(ref"A1", n(1))
        .put(ref"A2", CellValue.Formula("TRUE", Some(CellValue.Bool(true))))
        .put(ref"B1", CellValue.Formula("TRUE", Some(CellValue.Bool(true))))
        .put(ref"C1", CellValue.Formula(formula))
      val book = Workbook(sheet)
      val results = List(
        book.recalculate(RecalcOptions()),
        book.recalculateAfterEdit(sheet.name, Set(ref"C1"), RecalcOptions())
      )
      results.foreach { result =>
        assertEquals(result.errors, Vector.empty, formula)
        assertEquals(c1(result.workbook), CellValue.Formula(formula, Some(n(expected))), formula)
      }
    }
  }

  test("GH-695: a Normal-kind anchor stays an anchor in every evaluation fold") {
    // an xl-authored dynamic formula is a plain <f>; `x#` evaluates its formula as an array
    val sheet = Sheet("Sheet1")
      .put(ref"B1", n(3))
      .put(ref"B2", n(1))
      .put(ref"B3", n(2))
      .put(ref"A1", CellValue.Formula("SORT(B1:B3)"))
      .put(ref"C1", CellValue.Formula("SUM(A1#)"))
    val book = Workbook(sheet)
    val sum = CellValue.Formula("SUM(A1#)", Some(n(6)))
    assertEquals(c1(book.recalculate(Clock.system).workbook), sum)
    assertEquals(c1(book.recalculate(RecalcOptions(parallelism = 4)).workbook), sum, "∥")
    val targeted =
      book.recalculateAfterEdit(SheetName.unsafe("Sheet1"), Set(ref"B1"), RecalcOptions())
    assertEquals(c1(targeted.workbook), sum, "targeted")
    assertEquals(
      sheet.evaluateWithDependencyCheck(Clock.system, Some(book)).map(_.get(ref"C1")),
      Right(Some(n(6)))
    )
  }

  test("GH-695: a fold reads a scalar or erroring anchor as x# does when it evaluates it itself") {
    // a scalar result is not a spill anchor (#REF!); an error result is that error
    List("1+1" -> CellError.Ref, "1/0" -> CellError.Div0).foreach { (formula, expected) =>
      List("SUM(A1#)", "ROWS(A1#)").foreach { reader =>
        val sheet = Sheet("Sheet1")
          .put(ref"A1", arrayRecord(ref"A1", formula))
          .put(ref"C1", CellValue.Formula(reader))
        val book = Workbook(sheet)
        val direct = sheet.evaluateFormula(s"=$reader", Clock.system, Some(book), Some(ref"C1"))
        assert(
          direct.isLeft || direct == Right(CellValue.Error(expected)),
          s"$formula $reader: $direct"
        )
        assertEquals(
          c1(book.recalculate(Clock.system).workbook),
          CellValue.Formula(reader, Some(CellValue.Error(expected))),
          s"$formula $reader"
        )
      }
    }
  }

  test("GH-695: a bare reference to a formula cached with an error is that error") {
    // folds now thread formula records cached with their value, so `=Y1` over an erroring Y1
    // reads the record — it must read as the error cell it replaced, not as 0
    val sheet = Sheet("Sheet1")
      .put(ref"Y1", CellValue.Formula("X1+1", Some(CellValue.Error(CellError.Div0))))
    assertEquals(sheet.evaluateFormula("=Y1"), Right(CellValue.Error(CellError.Div0)))
  }
