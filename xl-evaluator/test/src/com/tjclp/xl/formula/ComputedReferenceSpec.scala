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
