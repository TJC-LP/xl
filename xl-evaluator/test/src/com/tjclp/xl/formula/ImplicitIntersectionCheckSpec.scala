package com.tjclp.xl.formula

import java.time.{LocalDate, LocalDateTime}

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.formula.eval.ImplicitIntersection
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
    ImplicitIntersection.check(book(extra*), extra.map((r, _) => S -> r).toVector, clock)

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
    assertEquals(ImplicitIntersection.check(wb, Vector(S -> ref"C5"), clock), Vector.empty)
    assertEquals(check(ref"C5" -> "SUM(NoSuchName*2)"), Vector.empty)
  }

  test("hits come back in sheet, then row-major, order and deterministically") {
    val targets = Vector(ref"C9", ref"C1", ref"D1").map(r => r -> "A1:A10*2")
    val wb = book(targets*)
    val asked = targets.reverse.map((r, _) => S -> r)
    val hits = ImplicitIntersection.check(wb, asked, clock)
    assertEquals(hits.map(_.cell), Vector(ref"C1", ref"D1", ref"C9"))
    assertEquals(ImplicitIntersection.check(wb, asked, clock), hits)
  }
