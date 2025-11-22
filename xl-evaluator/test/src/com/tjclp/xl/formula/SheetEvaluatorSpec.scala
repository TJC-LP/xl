package com.tjclp.xl.formula

import com.tjclp.xl.*
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.SheetEvaluator.* // Extension methods
import com.tjclp.xl.sheets.Sheet
import munit.FunSuite
import java.time.{LocalDate, LocalDateTime}

/**
 * Tests for Sheet evaluation API (WI-09b Phase 2).
 *
 * Tests evaluateFormula, evaluateCell, and evaluateAllFormulas extension methods.
 */
class SheetEvaluatorSpec extends FunSuite:
  val emptySheet = new Sheet(name = SheetName.unsafe("Test"))

  def sheetWith(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(emptySheet) { case (s, (ref, value)) =>
      s.put(ref, value)
    }

  // ===== evaluateFormula Tests =====

  test("evaluateFormula: simple arithmetic") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Number(BigDecimal(20))
    )
    val result = sheet.evaluateFormula("=A1+B1")
    assertEquals(result, Right(CellValue.Number(BigDecimal(30))))
  }

  test("evaluateFormula: SUM function") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(1)),
      ref"A2" -> CellValue.Number(BigDecimal(2)),
      ref"A3" -> CellValue.Number(BigDecimal(3))
    )
    val result = sheet.evaluateFormula("=SUM(A1:A3)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(6))))
  }

  test("evaluateFormula: with leading equals") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(42)))
    val result = sheet.evaluateFormula("=A1*2")
    assertEquals(result, Right(CellValue.Number(BigDecimal(84))))
  }

  test("evaluateFormula: without leading equals") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(42)))
    val result = sheet.evaluateFormula("A1*2")
    assertEquals(result, Right(CellValue.Number(BigDecimal(84))))
  }

  test("evaluateFormula: text result") {
    val result = emptySheet.evaluateFormula("=UPPER(\"hello\")")
    assertEquals(result, Right(CellValue.Text("HELLO")))
  }

  test("evaluateFormula: boolean result") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(10)))
    val result = sheet.evaluateFormula("=A1>5")
    assertEquals(result, Right(CellValue.Bool(true)))
  }

  test("evaluateFormula: date result with TODAY") {
    val clock = Clock.fixedDate(LocalDate.of(2025, 11, 21))
    val result = emptySheet.evaluateFormula("=TODAY()", clock)
    assertEquals(result, Right(CellValue.DateTime(LocalDate.of(2025, 11, 21).atStartOfDay())))
  }

  test("evaluateFormula: parse error") {
    val result = emptySheet.evaluateFormula("=INVALID(")
    assert(result.isLeft)
  }

  test("evaluateFormula: eval error (division by zero)") {
    val result = emptySheet.evaluateFormula("=10/0")
    assert(result.isLeft)
  }

  test("evaluateFormula: MIN function") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(5)),
      ref"A2" -> CellValue.Number(BigDecimal(10)),
      ref"A3" -> CellValue.Number(BigDecimal(3))
    )
    val result = sheet.evaluateFormula("=MIN(A1:A3)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(3))))
  }

  test("evaluateFormula: MAX function") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(5)),
      ref"A2" -> CellValue.Number(BigDecimal(10)),
      ref"A3" -> CellValue.Number(BigDecimal(3))
    )
    val result = sheet.evaluateFormula("=MAX(A1:A3)")
    assertEquals(result, Right(CellValue.Number(BigDecimal(10))))
  }

  test("evaluateFormula: CONCATENATE function") {
    val result = emptySheet.evaluateFormula("=CONCATENATE(\"Hello\", \" \", \"World\")")
    assertEquals(result, Right(CellValue.Text("Hello World")))
  }

  test("evaluateFormula: LEFT function") {
    val result = emptySheet.evaluateFormula("=LEFT(\"Hello\", 3)")
    assertEquals(result, Right(CellValue.Text("Hel")))
  }

  test("evaluateFormula: DATE function") {
    val result = emptySheet.evaluateFormula("=DATE(2025, 11, 21)")
    assertEquals(result, Right(CellValue.DateTime(LocalDate.of(2025, 11, 21).atStartOfDay())))
  }

  // ===== evaluateCell Tests =====

  test("evaluateCell: formula cell") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1*2")
    )
    val result = sheet.evaluateCell(ref"B1")
    assertEquals(result, Right(CellValue.Number(BigDecimal(20))))
  }

  test("evaluateCell: non-formula cell returns unchanged") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(42)))
    val result = sheet.evaluateCell(ref"A1")
    assertEquals(result, Right(CellValue.Number(BigDecimal(42))))
  }

  test("evaluateCell: empty cell returns Empty") {
    val result = emptySheet.evaluateCell(ref"Z99")
    assertEquals(result, Right(CellValue.Empty))
  }

  test("evaluateCell: text cell returns unchanged") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("Hello"))
    val result = sheet.evaluateCell(ref"A1")
    assertEquals(result, Right(CellValue.Text("Hello")))
  }

  test("evaluateCell: formula with parse error") {
    val sheet = sheetWith(ref"A1" -> CellValue.Formula("=INVALID("))
    val result = sheet.evaluateCell(ref"A1")
    assert(result.isLeft)
  }

  test("evaluateCell: formula with eval error") {
    val sheet = sheetWith(ref"A1" -> CellValue.Formula("=10/0"))
    val result = sheet.evaluateCell(ref"A1")
    assert(result.isLeft)
  }

  // ===== evaluateAllFormulas Tests =====

  test("evaluateAllFormulas: empty sheet") {
    val result = emptySheet.evaluateAllFormulas()
    assertEquals(result, Right(Map.empty[ARef, CellValue]))
  }

  test("evaluateAllFormulas: sheet with no formulas") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Text("Hello")
    )
    val result = sheet.evaluateAllFormulas()
    assertEquals(result, Right(Map.empty[ARef, CellValue]))
  }

  test("evaluateAllFormulas: single formula") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1*2")
    )
    val result = sheet.evaluateAllFormulas()
    assert(result.isRight)
    result.foreach { results =>
      assertEquals(results.size, 1)
      assertEquals(results.get(ref"B1"), Some(CellValue.Number(BigDecimal(20))))
    }
  }

  test("evaluateAllFormulas: multiple formulas") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1*2"),
      ref"C1" -> CellValue.Formula("=A1+5"),
      ref"D1" -> CellValue.Text("Not a formula")
    )
    val result = sheet.evaluateAllFormulas()
    assert(result.isRight)
    result.foreach { results =>
      assertEquals(results.size, 2)
      assertEquals(results.get(ref"B1"), Some(CellValue.Number(BigDecimal(20))))
      assertEquals(results.get(ref"C1"), Some(CellValue.Number(BigDecimal(15))))
    }
  }

  test("evaluateAllFormulas: stops on first error") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Formula("=10/0"),
      ref"B1" -> CellValue.Formula("=20+2")
    )
    val result = sheet.evaluateAllFormulas()
    assert(result.isLeft)
  }

  test("evaluateAllFormulas: text functions with literals") {
    // Note: Cell references in text functions require type-aware parsing (future enhancement)
    // For now, test with literal text
    val sheet = sheetWith(
      ref"B1" -> CellValue.Formula("=UPPER(\"hello\")"),
      ref"C1" -> CellValue.Formula("=CONCATENATE(\"Hello\", \" \", \"World\")")
    )
    val result = sheet.evaluateAllFormulas()
    result match
      case Right(results) =>
        assertEquals(results.size, 2)
        assertEquals(results.get(ref"B1"), Some(CellValue.Text("HELLO")))
        assertEquals(results.get(ref"C1"), Some(CellValue.Text("Hello World")))
      case Left(error) =>
        fail(s"Expected Right but got Left($error)")
  }

  test("evaluateAllFormulas: date functions with clock") {
    val clock = Clock.fixedDate(LocalDate.of(2025, 11, 21))
    val sheet = sheetWith(
      ref"A1" -> CellValue.Formula("=TODAY()"),
      ref"B1" -> CellValue.Formula("=YEAR(TODAY())")
    )
    val result = sheet.evaluateAllFormulas(clock)
    assert(result.isRight)
    result.foreach { results =>
      assertEquals(results.size, 2)
      assertEquals(results.get(ref"A1"), Some(CellValue.DateTime(LocalDate.of(2025, 11, 21).atStartOfDay())))
      assertEquals(results.get(ref"B1"), Some(CellValue.Number(BigDecimal(2025))))
    }
  }

  // ===== evaluateWithDependencyCheck Tests (WI-09d Integration) =====

  test("evaluateWithDependencyCheck: no cycles, simple chain") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1+5"),
      ref"C1" -> CellValue.Formula("=B1*2")
    )
    val result = sheet.evaluateWithDependencyCheck()
    assert(result.isRight)
    result.foreach { results =>
      assertEquals(results.size, 2) // B1 and C1 (constants not included)
      assertEquals(results.get(ref"B1"), Some(CellValue.Number(BigDecimal(15))))
      assertEquals(results.get(ref"C1"), Some(CellValue.Number(BigDecimal(30))))
    }
  }

  test("evaluateWithDependencyCheck: circular reference detected") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Formula("=B1"),
      ref"B1" -> CellValue.Formula("=A1")
    )
    val result = sheet.evaluateWithDependencyCheck()
    assert(result.isLeft)
    result match
      case scala.util.Left(error) =>
        // Should be a FormulaError with CircularRef
        assert(error.toString.contains("Circular") || error.toString.contains("circular"))
      case _ => fail("Expected Left with circular reference error")
  }

  test("evaluateWithDependencyCheck: self-reference detected") {
    val sheet = sheetWith(ref"A1" -> CellValue.Formula("=A1"))
    val result = sheet.evaluateWithDependencyCheck()
    assert(result.isLeft)
  }

  test("evaluateWithDependencyCheck: diamond dependency") {
    // A1 = 10, B1 = A1+5, C1 = A1*2, D1 = B1+C1
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"B1" -> CellValue.Formula("=A1+5"),
      ref"C1" -> CellValue.Formula("=A1*2"),
      ref"D1" -> CellValue.Formula("=B1+C1")
    )
    val result = sheet.evaluateWithDependencyCheck()
    assert(result.isRight)
    result.foreach { results =>
      assertEquals(results.get(ref"B1"), Some(CellValue.Number(BigDecimal(15))))
      assertEquals(results.get(ref"C1"), Some(CellValue.Number(BigDecimal(20))))
      assertEquals(results.get(ref"D1"), Some(CellValue.Number(BigDecimal(35)))) // 15 + 20
    }
  }

  test("evaluateWithDependencyCheck: formula referencing another formula") {
    // Test that formulas can reference other formulas and get evaluated values
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(5)),
      ref"A2" -> CellValue.Formula("=A1*2"),    // 10
      ref"A3" -> CellValue.Formula("=A2+A1"),   // 10 + 5 = 15
      ref"A4" -> CellValue.Formula("=A3*A2")    // 15 * 10 = 150
    )
    val result = sheet.evaluateWithDependencyCheck()
    assert(result.isRight)
    result.foreach { results =>
      assertEquals(results.get(ref"A2"), Some(CellValue.Number(BigDecimal(10))))
      assertEquals(results.get(ref"A3"), Some(CellValue.Number(BigDecimal(15))))
      assertEquals(results.get(ref"A4"), Some(CellValue.Number(BigDecimal(150))))
    }
  }

  test("evaluateWithDependencyCheck: multiple disconnected components") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Formula("=A1+5"),
      ref"B1" -> CellValue.Number(BigDecimal(20)),
      ref"B2" -> CellValue.Formula("=B1*2")
    )
    val result = sheet.evaluateWithDependencyCheck()
    assert(result.isRight)
    result.foreach { results =>
      assertEquals(results.get(ref"A2"), Some(CellValue.Number(BigDecimal(15))))
      assertEquals(results.get(ref"B2"), Some(CellValue.Number(BigDecimal(40))))
    }
  }

  test("evaluateWithDependencyCheck: range aggregation functions") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(1)),
      ref"A2" -> CellValue.Number(BigDecimal(2)),
      ref"A3" -> CellValue.Number(BigDecimal(3)),
      ref"B1" -> CellValue.Formula("=SUM(A1:A3)"),
      ref"C1" -> CellValue.Formula("=B1*2")
    )
    val result = sheet.evaluateWithDependencyCheck()
    assert(result.isRight)
    result.foreach { results =>
      assertEquals(results.get(ref"B1"), Some(CellValue.Number(BigDecimal(6))))
      assertEquals(results.get(ref"C1"), Some(CellValue.Number(BigDecimal(12))))
    }
  }

  test("evaluateWithDependencyCheck: empty sheet") {
    val result = emptySheet.evaluateWithDependencyCheck()
    assert(result.isRight)
    result.foreach { results =>
      assertEquals(results.size, 0)
    }
  }
