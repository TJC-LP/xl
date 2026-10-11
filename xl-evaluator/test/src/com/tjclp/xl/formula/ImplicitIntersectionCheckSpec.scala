package com.tjclp.xl.formula

import java.time.{LocalDate, LocalDateTime}

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.formula.eval.ImplicitIntersection
import com.tjclp.xl.formula.functions.FunctionRegistry
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/** GH-714: `ImplicitIntersection.check` — a plain cell's value against its array value. */
class ImplicitIntersectionCheckSpec extends FunSuite:

  private val S = SheetName.unsafe("S")
  private val clock = Clock.fixed(LocalDate.of(2026, 1, 2), LocalDateTime.of(2026, 1, 2, 3, 4, 5))
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  /** A1:A10 = 1..10 and B1:B10 = 1..10, plus `extra` formulas. */
  private def book(extra: (ARef, String)*): Workbook =
    val base = (1 to 10).foldLeft(Sheet(S)) { (s, i) =>
      s.put(ARef.from0(0, i - 1), num(i)).put(ARef.from0(1, i - 1), num(i))
    }
    Workbook(Vector(extra.foldLeft(base) { case (s, (r, f)) => s.put(r, CellValue.Formula(f)) }))

  private def check(extra: (ARef, String)*) =
    ImplicitIntersection.check(book(extra*), extra.map((r, _) => S -> r).toVector, clock).hits

  test("a product summed in row 5: plain 25, array 385") {
    val hits = check(ref"C5" -> "SUM(A1:A10*B1:B10)")
    assertEquals(hits.size, 1)
    val d = hits.headOption.getOrElse(fail("no hit"))
    assertEquals((d.plain, d.arrayTopLeft, d.rows, d.cols), (num(25), num(385), 1, 1))
  }

  test("shape divergences: a range operation, outside its rows, SORT, SEQUENCE") {
    val shape = check(ref"C1" -> "A1:A10*2").headOption.getOrElse(fail("no hit"))
    assertEquals((shape.plain, shape.rows, shape.cols), (num(2), 10, 1))
    val outside = check(ref"C20" -> "A1:A10*2").headOption.getOrElse(fail("no hit"))
    assertEquals(outside.plain, CellValue.Error(CellError.Value))
    assert(check(ref"D1" -> "SORT(A1:A3)").nonEmpty)
    assert(check(ref"D1" -> "SEQUENCE(3)").nonEmpty)
  }

  test("silent: scalar formulas, SUMPRODUCT, @, a whole-range aggregate, volatile scalars") {
    List("A1*2", "SUMPRODUCT(A1:A10*B1:B10)", "@A1:A10*2", "SUM(A1:A10)", "RAND()", "NOW()")
      .foreach(f => assertEquals(check(ref"C5" -> f), Vector.empty, f))
  }

  test("only plain formulas are checked; a host failure is no finding") {
    val wb = book().put(
      book().sheets.headOption
        .getOrElse(fail("no sheet"))
        .put(
          ref"C5",
          CellValue.Formula(
            "SUM(A1:A10*B1:B10)",
            None,
            FormulaKind.dynamicArray(CellRange(ref"C5", ref"C5"))
          )
        )
    )
    assertEquals(ImplicitIntersection.check(wb, Vector(S -> ref"C5"), clock).hits, Vector.empty)
    assertEquals(check(ref"C5" -> "SUM(NoSuchName*2)"), Vector.empty)
  }

  test("hits come back in sheet, then row-major, order and deterministically") {
    val targets = Vector(ref"C9", ref"C1", ref"D1").map(r => r -> "A1:A10*2")
    val wb = book(targets*)
    val asked = targets.reverse.map((r, _) => S -> r)
    val hits = ImplicitIntersection.check(wb, asked, clock).hits
    assertEquals(hits.map(_.cell), Vector(ref"C1", ref"D1", ref"C9"))
    assertEquals(ImplicitIntersection.check(wb, asked, clock).hits, hits)
  }

  private def mayDiverge(formula: String): Boolean =
    FormulaParser
      .parse(formula)
      .fold(e => fail(s"$formula: $e"), ImplicitIntersection.mayDiverge)

  test("a range taken whole by its slot is never a candidate: aggregates, criteria, lookups") {
    List(
      "=SUM($A$1:A1)",
      "=VLOOKUP(C1,$A$1:$B$10,2,FALSE)",
      "=COUNTIF(A1:A10,\">3\")",
      "=SUMPRODUCT(A1:A10,B1:B10)",
      "=SUMIF(A:A,\">3\",B:B)",
      "=MATCH(C1,A1:A10,0)+AVERAGE(B1:B10)",
      "=IF(A1>0,SUM(A1:A10),0)",
      "=A1*2",
      "=ROUND(A1/3,2)&\" units\"",
      "=LET(x,A1*2,x+1)",
      "=TODAY()+1"
    ).foreach(f => assert(!mayDiverge(f), f))
  }

  test("an array source in a value position is a candidate") {
    List(
      "=SUM(A1:A10*B1:B10)",
      "=A1:A10*2",
      "=A1:A10",
      "=A:A",
      "=IF(A1:A10>5,1,0)",
      "=SORT(A1:A3)",
      "=SEQUENCE(3)",
      "=TRANSPOSE(A1:B2)",
      "=INDEX(A1:B10,0,1)",
      "=ABS(A1:A3)",
      "=SUM({1,2,3}*A1)",
      "=Rates*2",
      "=LET(r,A1:A10,SUM(r))",
      "=ROW(A1:A3)"
    ).foreach(f => assert(mayDiverge(f), f))
  }

  test("array-valued functions come from the specs' declared result type") {
    val names = FunctionRegistry.arrayResultNames
    List("SEQUENCE", "SORT", "TRANSPOSE", "FILTER", "UNIQUE", "INDEX", "OFFSET", "INDIRECT")
      .foreach(n => assert(names.contains(n), n))
    List("SUM", "VLOOKUP", "COUNTIF", "IF", "SUMPRODUCT").foreach(n =>
      assert(!names.contains(n), n)
    )
  }

  test("a running-total drag is never evaluated twice") {
    val drag = (1 to 200).map(i => ARef.from0(2, i - 1) -> s"SUM($$A$$1:A$i)")
    val report = ImplicitIntersection.check(book(drag*), drag.map((r, _) => S -> r).toVector, clock)
    assertEquals((report.hits, report.checked, report.candidates), (Vector.empty, 0, 0))
  }

  test("the two-way evaluation stops at the budget and says it sampled") {
    val drag = (1 to 12).map(i => ARef.from0(2, i - 1) -> "A$1:A$10*2")
    val targets = drag.map((r, _) => S -> r).toVector
    val report = ImplicitIntersection.check(book(drag*), targets, clock, budget = 5)
    assertEquals((report.hits.size, report.checked, report.candidates), (5, 5, 12))
    assert(report.sampled)
    assertEquals(report.hits.map(_.cell), (1 to 5).map(i => ARef.from0(2, i - 1)).toVector)
    val full = ImplicitIntersection.check(book(drag*), targets, clock)
    assertEquals((full.hits.size, full.sampled), (12, false))
  }
