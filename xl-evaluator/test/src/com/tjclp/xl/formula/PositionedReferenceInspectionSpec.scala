package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.FormulaEvaluationError

class PositionedReferenceInspectionSpec extends FunSuite:
  private def evaluate(book: Workbook, sheet: Sheet, formula: String, at: String) =
    ARef
      .parse(at)
      .left
      .map(e => FormulaEvaluationError(com.tjclp.xl.error.XLError.Other(e.toString)))
      .flatMap(cell => SheetEvaluator.evaluateAt(book, sheet, cell, formula))

  private val sheet = Sheet("S").put(ref"D5", CellValue.Number(BigDecimal(99)))
  private val wb = Workbook(sheet)

  List(
    "ROW(D5)" -> 5,
    "COLUMN(D5)" -> 4,
    "ROWS(D5:D6)" -> 2,
    "CELL(\"row\",D5)" -> 5,
    "CELL(\"col\",D5)" -> 4
  ).foreach { (formula, expected) =>
    test(s"positioned $formula can inspect its own address without reading its value") {
      assertEquals(
        evaluate(wb, sheet, formula, "D5"),
        Right(CellValue.Number(BigDecimal(expected)))
      )
    }
  }

  // Static and run-time agreement: in a reference-only position (ROWS, COLUMNS, AREAS) each
  // ValueDependencies.referencePassing entry passes its reference through unread, so the static
  // check lets the placement through; the run-time guard must agree and read no value either.
  // Excel gives no cycle for any of these.
  private val lookup = sheet
    .put(ref"A1", CellValue.Number(BigDecimal(9)))
    .put(ref"A2", CellValue.Number(BigDecimal(1)))
    .put(ref"D1", CellValue.Number(BigDecimal(7)))

  List(
    ("ROWS(XLOOKUP(9,A1:A2,D1:D2))", "D1", 1),
    ("ROWS(OFFSET(D5,0,0,2,1))", "D5", 2),
    ("ROWS(IF(TRUE,D5:D6,A1))", "D5", 2),
    ("ROWS(CHOOSE(1,D5:D6))", "D5", 2),
    ("ROWS(IFS(TRUE,D5:D6))", "D5", 2),
    ("ROWS(SWITCH(1,1,D5:D6,A1))", "D5", 2),
    ("ROWS(SWITCH(2,1,A1,D5:D6))", "D5", 2),
    ("ROWS(INDEX(D5:D6,0))", "D5", 2),
    ("ROWS(INDEX(D5:D6,A2))", "D5", 1),
    ("ROWS(D4:D6 D5:E9)", "D5", 2),
    ("AREAS((D5,A1))", "D5", 2)
  ).foreach { (formula, at, expected) =>
    test(s"positioned $formula at $at passes its reference through without reading it") {
      assertEquals(
        evaluate(Workbook(lookup), lookup, formula, at),
        Right(CellValue.Number(BigDecimal(expected)))
      )
    }
  }

  test("positioned ordinary scalar reference still refuses a real circular read") {
    assert(evaluate(wb, sheet, "D5+1", "D5").left.exists(_.cycle.nonEmpty))
  }

  test("positioned ROW without a reference composes with array semantics changes") {
    assertEquals(
      evaluate(wb, sheet, "ROW()+COLUMN()", "D5"),
      Right(CellValue.Number(BigDecimal(9.0)))
    )
  }

  test("inspection nested in arithmetic does not create a circular value dependency") {
    assertEquals(
      evaluate(wb, sheet, "ROW(D5)+ROWS(D5:D6)", "D5"),
      Right(CellValue.Number(BigDecimal(7.0)))
    )
  }

  test("an INDEX selector that reads the placed cell is still circular") {
    assert(evaluate(wb, sheet, "ROW(INDEX(A1:A10,D5))", "D5").left.exists(_.cycle.nonEmpty))
  }

  test("a precedent may inspect the placement address without reading its old value") {
    val source = sheet.put(ref"A1", CellValue.Formula("ROW(D5)", None))
    assertEquals(
      evaluate(Workbook(source), source, "A1+1", "D5"),
      Right(CellValue.Number(BigDecimal(6.0)))
    )
  }

  test("an existing self-inspection formula does not poison the closure order") {
    val source = sheet.put(ref"A1", CellValue.Formula("ROW(A1)", None))
    assertEquals(
      evaluate(Workbook(source), source, "A1+1", "D5"),
      Right(CellValue.Number(BigDecimal(2.0)))
    )
  }

  List(
    "LET(chosen,D5,ROWS(chosen))" -> 1,
    "LET(chosen,INDEX(D5:D6,1),ROWS(chosen))" -> 1,
    "LET(outer,D5,LET(inner,outer,COLUMNS(inner)))" -> 1
  ).foreach { (formula, expected) =>
    test(s"positioned $formula follows LET references without a circular value read") {
      assertEquals(
        evaluate(wb, sheet, formula, "D5"),
        Right(CellValue.Number(BigDecimal(expected)))
      )
    }
  }

  List(
    "LET(chosen,D5,chosen+1)",
    "LET(chosen,INDEX(A1:A10,D5),ROWS(chosen))",
    "LET(unused,D5+1,ROW(A1))"
  ).foreach { formula =>
    test(s"positioned $formula retains actual LET value reads") {
      assert(evaluate(wb, sheet, formula, "D5").left.exists(_.cycle.nonEmpty))
    }
  }
