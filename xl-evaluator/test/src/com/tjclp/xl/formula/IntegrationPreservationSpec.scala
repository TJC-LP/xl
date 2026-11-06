package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import munit.FunSuite

/** Preserve main's behavior when GH-741 moves evaluation from the CLI into the library. */
class IntegrationPreservationSpec extends FunSuite:
  private def n(i: Int): CellValue = CellValue.Number(BigDecimal(i))

  test("GH-763/GH-741: overrides invalidate cached dynamic cross-sheet readers") {
    val data = Sheet("Data").put(ref"B1", n(10)).put(ref"B2", n(1))
    val other = Sheet("Other")
      .put(ref"A1", CellValue.Formula("SUM(OFFSET(Data!B1,0,0,2,1))", Some(n(11))))
      .put(ref"A2", CellValue.Formula("SUM(INDIRECT(\"Data!B1:B2\"))", Some(n(11))))
    val wb = Workbook(data, other)
    List("Other!A1", "Other!A2").foreach { formula =>
      assertEquals(
        SheetEvaluator.evaluateWithClosure(wb, data, formula, Map(ref"B2" -> n(0))),
        Right(n(10))
      )
    }
  }

  test("GH-802/GH-741: position-dependent INDEX cannot hide a placement cycle") {
    val data = Sheet("Data").put(ref"A1", n(1)).put(ref"A2", n(2)).put(ref"A3", n(3))
    val result = SheetEvaluator.evaluateAt(Workbook(data), data, "INDEX(A1:A3,ROW())", ref"A3")
    assert(result.left.exists(_.cycle.nonEmpty), result.toString)
  }

  test("GH-802/GH-741: a positioned INDEX reading the previous row is not circular") {
    val data = Sheet("Data").put(ref"A1", n(1)).put(ref"A2", n(2)).put(ref"A3", n(3))
    assertEquals(
      SheetEvaluator.evaluateAt(Workbook(data), data, "INDEX(A1:A3,ROW()-1)", ref"A3"),
      Right(n(2))
    )
  }

  test("GH-741: a supplied sheet absent from the original workbook supports spills") {
    val source = Sheet("New")
      .put(ref"A1", CellValue.Formula("SEQUENCE(3)"))
      .put(ref"B1", CellValue.Formula("SUM(A1#)"))
    assertEquals(
      SheetEvaluator.evaluateWithClosure(Workbook(Sheet("Old")), source, "B1"),
      Right(n(6))
    )
  }

  test("GH-741: a typed override cannot replace the placement cell") {
    val source = Sheet("Data").put(ref"A1", n(1))
    val result =
      SheetEvaluator.evaluateAt(Workbook(source), source, "A1", ref"D5", Map(ref"D5" -> n(2)))
    assert(
      result.left.exists(e =>
        e.cycle.isEmpty && e.error.toString.contains("override sets the placement cell D5")
      ),
      result.toString
    )
  }
