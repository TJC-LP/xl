package com.tjclp.xl.formula

import java.time.{LocalDate, LocalDateTime}

import com.tjclp.xl.{*, given}
import com.tjclp.xl.formula.eval.DependentRecalculation

import munit.FunSuite

class AfterEditRecalculationSpec extends FunSuite:
  private def num(value: Int): CellValue = CellValue.Number(BigDecimal(value))
  private val data = SheetName.unsafe("Data")

  private def cached(wb: Workbook, name: String, ref: ARef): Option[CellValue] =
    wb(SheetName.unsafe(name)).fold(error => fail(error.message), identity)(ref).value match
      case CellValue.Formula(_, value, _) => value
      case other => fail(s"Expected formula, got $other")

  private final class CountingClock extends Clock:
    var calls = 0
    def today(): LocalDate =
      calls += 1
      LocalDate.of(2026, 9, 4)
    def now(): LocalDateTime =
      calls += 1
      LocalDateTime.of(2026, 9, 4, 12, 0)

  test("GH-504: targeted recalculation does not evaluate unrelated volatile formulas") {
    val wb = Workbook(
      Sheet("Data").put(ref"A1", num(2)).put(ref"B1", CellValue.Formula("A1*2", Some(num(2)))),
      Sheet("Other").put(ref"A1", CellValue.Formula("TODAY()", Some(num(123))))
    )
    val clock = new CountingClock
    val result = DependentRecalculation.recalculateAfterEdit(wb, data, Set(ref"A1"), clock)
    assertEquals(cached(result.workbook, "Data", ref"B1"), Some(num(4)))
    assertEquals(cached(result.workbook, "Other", ref"A1"), Some(num(123)))
    assertEquals(result.evaluated(SheetName.unsafe("Other")), Map.empty[ARef, CellValue])
    assertEquals(clock.calls, 0)
  }

  test("GH-504: an authored formula cannot trust a cache from an existing upstream cycle") {
    val wb = Workbook(
      Sheet("Data")
        .put(ref"B1", CellValue.Formula("B2", Some(num(1))))
        .put(ref"B2", CellValue.Formula("B1", Some(num(1))))
        .put(ref"C1", CellValue.Formula("B1+1", Some(num(2))))
    )
    val result = DependentRecalculation.recalculateAfterEdit(wb, data, Set(ref"C1"), Clock.system)
    assertEquals(cached(result.workbook, "Data", ref"C1"), None)
    assert(result.errors.exists(_.ref == ref"C1"))
    assert(!result.certified)
  }

  test("GH-504: dynamic reads also reject cached values from an existing upstream cycle") {
    val wb = Workbook(
      Sheet("Data")
        .put(ref"B1", CellValue.Formula("B2", Some(num(1))))
        .put(ref"B2", CellValue.Formula("B1", Some(num(1))))
        .put(ref"C1", CellValue.Formula("INDIRECT(\"B1\")+1", Some(num(2))))
    )
    val result = DependentRecalculation.recalculateAfterEdit(wb, data, Set(ref"C1"), Clock.system)
    assertEquals(cached(result.workbook, "Data", ref"C1"), None)
    assert(result.errors.exists(_.ref == ref"C1"))
  }

  test("GH-504: healthy targeted results agree with full recalculation and are stable") {
    val wb = Workbook(
      Sheet("Data")
        .put(ref"A1", num(2))
        .put(ref"B1", CellValue.Formula("A1*2", Some(num(2))))
        .put(ref"C1", CellValue.Formula("B1+1", Some(num(3))))
    )
    val result = DependentRecalculation.recalculateAfterEdit(wb, data, Set(ref"A1"), Clock.system)
    assertEquals(result, wb.recalculate())
    assertEquals(
      DependentRecalculation
        .recalculateAfterEdit(result.workbook, data, Set(ref"A1"), Clock.system),
      result
    )
  }

  test("GH-504: authored formulas retain their current-cell context") {
    val wb = Workbook(Sheet("Data").put(ref"A5", CellValue.Formula("ROW()")))
    val result = DependentRecalculation.recalculateAfterEdit(wb, data, Set(ref"A5"), Clock.system)
    assertEquals(cached(result.workbook, "Data", ref"A5"), Some(num(5)))
  }

  test("GH-508: unresolved named readers are evaluated and invalidated after a single-cell edit") {
    val wb = Workbook(
      Sheet("Data").put(ref"A1", num(2)),
      Sheet("Other").put(ref"B2", CellValue.Formula("SUM(Multi)", Some(num(33))))
    )
      .withDefinedName("Multi", "Other!$A$1:$A$3,Other!$A$8:$A$10")
    val result = DependentRecalculation.recalculateAfterEdit(wb, data, Set(ref"A1"), Clock.system)
    assertEquals(cached(result.workbook, "Other", ref"B2"), None)
    assert(result.errors.exists(error => error.sheet.value == "Other" && error.ref == ref"B2"))
  }

  test("GH-504: an empty edit preserves the workbook without reading the clock") {
    val wb = Workbook(Sheet("Data").put(ref"A1", CellValue.Formula("TODAY()", Some(num(123)))))
    val clock = new CountingClock
    val result = DependentRecalculation.recalculateAfterEdit(wb, data, Set.empty, clock)
    assertEquals(result.workbook, wb)
    assertEquals(result.errors, Vector.empty)
    assertEquals(clock.calls, 0)
  }

  test("GH-504: a fixed lookup selector excludes unused table cells from cycle detection") {
    List("3", "$H$1").foreach { selector =>
      val wb = Workbook(
        Sheet("Data")
          .put(ref"D1", CellValue.Formula("1"))
          .put(ref"E1", CellValue.Formula(s"VLOOKUP(1,Data!D1:F1,$selector,FALSE)"))
          .put(ref"F1", num(42))
          .put(ref"H1", num(3))
      )
      val result = DependentRecalculation.recalculateAfterEdit(
        wb,
        data,
        Set(ref"D1", ref"E1", ref"F1"),
        Clock.system
      )
      assertEquals(result.errors, Vector.empty)
      assertEquals(cached(result.workbook, "Data", ref"E1"), Some(num(42)))
    }
    val horizontal = Workbook(
      Sheet("Data")
        .put(ref"A1", num(1))
        .put(ref"A2", CellValue.Formula("HLOOKUP(1,Data!A1:A3,3,FALSE)"))
        .put(ref"A3", num(42))
    )
    val result =
      DependentRecalculation.recalculateAfterEdit(horizontal, data, Set(ref"A2"), Clock.system)
    assertEquals(cached(result.workbook, "Data", ref"A2"), Some(num(42)))
  }

  test("GH-504: a lookup that selects its own result cell is still circular") {
    val wb = Workbook(
      Sheet("Data")
        .put(ref"D1", num(1))
        .put(ref"E1", CellValue.Formula("VLOOKUP(1,Data!D1:F1,2,FALSE)", Some(num(42))))
        .put(ref"F1", num(42))
    )
    val result = DependentRecalculation.recalculateAfterEdit(wb, data, Set(ref"E1"), Clock.system)
    assert(result.errors.nonEmpty)
    assertEquals(cached(result.workbook, "Data", ref"E1"), None)
  }
