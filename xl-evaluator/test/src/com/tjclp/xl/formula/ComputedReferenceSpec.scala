package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-713: a reference that INDEX, OFFSET, INDIRECT or XLOOKUP computes, read as a reference.
 *
 *   - As an endpoint of the range operator `:`: a scalar INDEX position keeps its reference in
 *     array mode, and an endpoint that fails carries its own error code (`#REF!` from INDIRECT or
 *     OFFSET, the `#N/A` an XLOOKUP fallback names), never a generic `#VALUE!`.
 *   - Under SUMPRODUCT: a computed whole column or row takes the extent shared with the written
 *     ranges beside it — leading blanks, an empty sheet, a range on another sheet — so its values
 *     line up element by element; without written ranges it keeps its own materialization.
 */
class ComputedReferenceSpec extends FunSuite:
  // the pinned edge-range gap below reads a written A2:A1048576 whole (a million cells), which
  // can pass munit's 30 s default on a loaded machine
  override val munitTimeout = scala.concurrent.duration.Duration(120, "s")

  private def number(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private val sheet = (1 to 3).foldLeft(Sheet(SheetName.unsafe("Data"))) { (s, i) =>
    s.put(ARef.from0(0, i - 1), number(i))
      .put(ARef.from0(1, i - 1), number(i * 10))
  }

  // Both helpers take the formula as a cell stores it, without the leading `=`; `evaluateFormula`
  // reads formula-bar text, so `array` adds it (the parser accepts the text either way).

  /** `formula` evaluated as an array formula (`evaluateFormula`, evala's mode). */
  private def array(formula: String): XLResult[CellValue] =
    sheet.evaluateFormula(s"=$formula", Clock.system, Some(Workbook(sheet)), None)

  /** `formula` as a plain formula cell at H9, read through the implicit intersection. */
  private def plain(formula: String): XLResult[CellValue] =
    val placed = sheet.put(ref"H9", CellValue.Formula(formula, None))
    placed.evaluateCell(ref"H9", Clock.system, Some(Workbook(placed)))

  test("computed scalar INDEX position retains its reference in array mode") {
    assertEquals(array("SUM(A1:INDEX(A1:A3,COUNTA(A1:A3)))"), Right(number(6)))
  }

  test("a computed range propagates INDIRECT's reference error") {
    assertEquals(plain("SUM(A1:INDIRECT(\"nowhere!!\"))"), Right(CellValue.Error(CellError.Ref)))
  }

  test("a computed range propagates OFFSET's reference error") {
    assertEquals(plain("SUM(A1:OFFSET(A1,-1,0))"), Right(CellValue.Error(CellError.Ref)))
  }

  test("a computed range propagates an explicit XLOOKUP fallback error") {
    assertEquals(plain("SUM(A1:XLOOKUP(9,A1:A3,B1:B3,#N/A))"), Right(CellValue.Error(CellError.NA)))
  }

  test("SUMPRODUCT INDEX whole-column result aligns with a written column after leading blanks") {
    val offset = Sheet(SheetName.unsafe("Offset"))
      .put(ref"A2", number(2))
      .put(ref"A3", number(3))
    assertEquals(
      offset.evaluateFormula("=SUMPRODUCT(INDEX(A:A,0),A:A)"),
      Right(number(13))
    )
  }

  test("SUMPRODUCT INDEX whole-column result agrees with an empty written column") {
    assertEquals(
      Sheet(SheetName.unsafe("Empty")).evaluateFormula("=SUMPRODUCT(INDEX(A:A,0),A:A)"),
      Right(number(0))
    )
  }

  test("computed whole axes share both ends of the cross-sheet SUMPRODUCT extent") {
    val offset = Sheet(SheetName.unsafe("Offset")).put(ref"A3", number(3))
    val other = Sheet(SheetName.unsafe("Other"))
      .put(ref"B2", number(10))
      .put(ref"B3", number(20))
      .put(ref"B5", number(30))
    assertEquals(
      offset.evaluateFormula(
        "=SUMPRODUCT(A:A,INDEX(Other!B:B,0))",
        Clock.system,
        Some(Workbook(offset, other))
      ),
      Right(number(60))
    )
    val row = Sheet(SheetName.unsafe("Row")).put(ref"B1", number(2)).put(ref"C1", number(3))
    assertEquals(row.evaluateFormula("=SUMPRODUCT(INDEX(1:1,0),1:1)"), Right(number(13)))
    assertEquals(
      Sheet(SheetName.unsafe("Empty")).evaluateFormula("=SUMPRODUCT(INDEX(1:1,0),1:1)"),
      Right(number(0))
    )
  }

  test("a computed whole axis without static range inputs keeps its materialization semantics") {
    val book = Workbook(sheet).withDefinedName("fullCol", "Data!$A:$A")
    assertEquals(
      sheet.evaluateFormula("=SUMPRODUCT(INDEX(fullCol,0))", Clock.system, Some(book)),
      Right(number(6))
    )
  }

  test("the shared extent reaches a nested call inside a SUMPRODUCT argument (pinned)") {
    // S!B1:B2 = 1,1 and T!A1:A5 = 1..5, the formula on S; `tcol` = T!$A:$A. SUMPRODUCT's shared
    // extent spans the used rows of the sheets its written whole columns name, and on this branch
    // it rides into every call inside the argument (see EvalContext.Operands).
    //   - Written, T!A:A brings T's rows 1-5 into the extent, so nothing is cut: 30, as
    //     LibreOffice 24.2.7 answers (SUM reads the whole column, 15 × (1+1)).
    //   - Through a name, T is not in the extent (S's rows 1-2 only), so the inner
    //     SUM(INDEX(tcol,0)) is clipped to T!A1:A2 and the answer is 3 × 2 = 6, where
    //     LibreOffice answers 30. This is the pinned divergence.
    // PR #719 stops the extent at the entry to a call that neither lifts nor enters it, which
    // should turn the name case into LibreOffice's 30; update this test when the branches meet.
    val s = Sheet(SheetName.unsafe("S")).put(ref"B1", number(1)).put(ref"B2", number(1))
    val t = (1 to 5).foldLeft(Sheet(SheetName.unsafe("T")))((acc, i) =>
      acc.put(ARef.from0(0, i - 1), number(i))
    )
    def at(formula: String): XLResult[CellValue] =
      val placed = s.put(ref"H1", CellValue.Formula(formula, None))
      val book = Workbook(Vector(placed, t)).withDefinedName("tcol", "T!$A:$A")
      placed.evaluateCell(ref"H1", Clock.system, Some(book))
    assertEquals(at("SUMPRODUCT(SUM(INDEX(T!A:A,0))*B:B)"), Right(number(30)))
    assertEquals(at("SUMPRODUCT(SUM(T!A:A)*B:B)"), Right(number(30)))
    assertEquals(at("SUMPRODUCT(SUM(INDEX(tcol,0))*B:B)"), Right(number(6)))
    // the computed column itself, aligned with B:B: 3 (1×1 + 2×1) either way, as LibreOffice
    assertEquals(at("SUMPRODUCT(INDEX(T!A:A,0)*B:B)"), Right(number(3)))
    assertEquals(at("SUMPRODUCT(INDEX(tcol,0)*B:B)"), Right(number(3)))
  }

  test("a computed edge range that starts below row 1 gets no shared extent (pinned, #746)") {
    // A1:A10 = 1, B1:B10 = 1..10. `B5:INDEX(B:B,1048576)` computes B5:B1048576, which is not a
    // whole column, so SUMPRODUCT gives it no shared extent; it stops at B's used rows, unlike a
    // written A5:A1048576 (all its rows) or a shared-bounded A:A — the lengths differ. Bounding it
    // like a whole column would not help (the written range beside it keeps its height); the fix
    // bounds edge ranges, written and computed, the same way: #746. LibreOffice 24.2.7 answers
    // 45, 45, 55 and 45.
    val f = (1 to 10).foldLeft(Sheet(SheetName.unsafe("F"))) { (acc, r) =>
      acc.put(ARef.from1(1, r), number(1)).put(ARef.from1(2, r), number(r))
    }
    def at(formula: String): XLResult[CellValue] =
      val placed = f.put(ref"H1", CellValue.Formula(formula, None))
      placed.evaluateCell(ref"H1", Clock.system, Some(Workbook(Vector(placed))))
    assertEquals(
      at("SUMPRODUCT((A:A>0)*(B5:INDEX(B:B,1048576)))"),
      Right(CellValue.Error(CellError.NA))
    )
    assertEquals(
      at("SUMPRODUCT((A5:A1048576>0)*(B5:INDEX(B:B,1048576)))"),
      Right(CellValue.Error(CellError.NA))
    )
    // a computed whole column takes the extent; alone, the edge range is just its data
    assertEquals(at("SUMPRODUCT((A:A>0)*(B1:INDEX(B:B,1048576)))"), Right(number(55)))
    assertEquals(at("SUMPRODUCT(B5:INDEX(B:B,1048576))"), Right(number(45)))
  }

  test("an edge-reaching range that starts below row 1 gets no shared extent (pinned gap)") {
    // A2:A4 = 1,2,3 and B2:B4 = 10,20,30. A range reaching row 1048576 but not starting at row 1
    // is not a whole column, so SUMPRODUCT gives it no shared extent: written, A2:A1048576 keeps
    // all its rows, while INDEX over B2:B1048576 stops at the used rows — the shapes differ.
    // LibreOffice 24.2.7 answers 140 for both. This predates GH-713 (0.23.1 gives the same
    // #N/A and #VALUE!); docs/LIMITATIONS.md lists it, and #746 tracks the fix.
    val e = List((2, 1, 10), (3, 2, 20), (4, 3, 30)).foldLeft(Sheet(SheetName.unsafe("E"))) {
      case (acc, (r, a, b)) => acc.put(ARef.from1(1, r), number(a)).put(ARef.from1(2, r), number(b))
    }
    def at(formula: String): XLResult[CellValue] =
      val placed = e.put(ref"H1", CellValue.Formula(formula, None))
      placed.evaluateCell(ref"H1", Clock.system, Some(Workbook(Vector(placed))))
    assertEquals(
      at("SUMPRODUCT(A2:A1048576*INDEX(B2:B1048576,0))"),
      Right(CellValue.Error(CellError.NA))
    )
    assertEquals(
      at("SUMPRODUCT(A2:A1048576,INDEX(B2:B1048576,0))"),
      Right(CellValue.Error(CellError.Value))
    )
    // what does line up: whole columns, or INDEX on both sides (two written edge ranges do too,
    // at a million cells each — left out to keep this suite fast)
    assertEquals(at("SUMPRODUCT(A:A*INDEX(B:B,0))"), Right(number(140)))
    assertEquals(at("SUMPRODUCT(INDEX(A2:A1048576,0)*INDEX(B2:B1048576,0))"), Right(number(140)))
  }
