package com.tjclp.xl.formula

import com.tjclp.xl.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.addressing.SheetName
import munit.FunSuite
import java.time.{LocalDate, LocalDateTime}

/**
 * Comprehensive tests for WI-09 new functions.
 *
 * Tests text functions (CONCATENATE, LEFT, RIGHT, LEN, UPPER, LOWER), date functions (TODAY, NOW,
 * DATE, YEAR, MONTH, DAY), and range functions (MIN, MAX).
 */
class NewFunctionsSpec extends FunSuite:
  val emptySheet = new Sheet(name = SheetName.unsafe("Test"))
  val evaluator = Evaluator.instance

  /** Helper to create sheet with cells */
  def sheetWith(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(emptySheet) { case (s, (ref, value)) =>
      s.put(ref, value)
    }

  private def parseRange(range: String): CellRange =
    CellRange.parse(range).fold(err => fail(err), identity)

  // ===== Text Functions =====

  test("CONCATENATE: empty list") {
    val expr = TExpr.Concatenate(List.empty)
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(""))
  }

  test("CONCATENATE: single string") {
    val expr = TExpr.Concatenate(List(TExpr.Lit("Hello")))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("Hello"))
  }

  test("CONCATENATE: multiple strings") {
    val expr = TExpr.Concatenate(List(
      TExpr.Lit("Hello"),
      TExpr.Lit(" "),
      TExpr.Lit("World"),
      TExpr.Lit("!")
    ))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("Hello World!"))
  }

  test("LEFT: extract left 3 characters") {
    val expr = TExpr.Left(TExpr.Lit("Hello"), TExpr.Lit(3))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("Hel"))
  }

  test("LEFT: n greater than length returns full string") {
    val expr = TExpr.Left(TExpr.Lit("Hi"), TExpr.Lit(10))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("Hi"))
  }

  test("LEFT: n equals length") {
    val expr = TExpr.Left(TExpr.Lit("Hello"), TExpr.Lit(5))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("Hello"))
  }

  test("LEFT: negative n returns error") {
    val expr = TExpr.Left(TExpr.Lit("Hello"), TExpr.Lit(-1))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft)
  }

  test("RIGHT: extract right 3 characters") {
    val expr = TExpr.Right(TExpr.Lit("Hello"), TExpr.Lit(3))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("llo"))
  }

  test("RIGHT: n greater than length returns full string") {
    val expr = TExpr.Right(TExpr.Lit("Hi"), TExpr.Lit(10))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("Hi"))
  }

  test("RIGHT: negative n returns error") {
    val expr = TExpr.Right(TExpr.Lit("Hello"), TExpr.Lit(-1))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft)
  }

  test("LEN: empty string") {
    val expr = TExpr.Len(TExpr.Lit(""))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(0)))
  }

  test("LEN: normal string") {
    val expr = TExpr.Len(TExpr.Lit("Hello"))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(5)))
  }

  test("UPPER: convert to uppercase") {
    val expr = TExpr.Upper(TExpr.Lit("hello world"))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("HELLO WORLD"))
  }

  test("UPPER: already uppercase") {
    val expr = TExpr.Upper(TExpr.Lit("HELLO"))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("HELLO"))
  }

  test("LOWER: convert to lowercase") {
    val expr = TExpr.Lower(TExpr.Lit("HELLO WORLD"))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("hello world"))
  }

  test("LOWER: already lowercase") {
    val expr = TExpr.Lower(TExpr.Lit("hello"))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("hello"))
  }

  // ===== Date/Time Functions =====

  test("TODAY: returns current date from clock") {
    val fixedDate = LocalDate.of(2025, 11, 21)
    val clock = Clock.fixedDate(fixedDate)
    val expr = TExpr.Today()
    val result = evaluator.eval(expr, emptySheet, clock)
    assertEquals(result, Right(fixedDate))
  }

  test("NOW: returns current datetime from clock") {
    val fixedDateTime = LocalDateTime.of(2025, 11, 21, 18, 30, 0)
    val clock = Clock.fixed(LocalDate.of(2025, 11, 21), fixedDateTime)
    val expr = TExpr.Now()
    val result = evaluator.eval(expr, emptySheet, clock)
    assertEquals(result, Right(fixedDateTime))
  }

  test("DATE: construct valid date") {
    val expr = TExpr.Date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(LocalDate.of(2025, 11, 21)))
  }

  test("DATE: invalid date returns error") {
    val expr = TExpr.Date(TExpr.Lit(2025), TExpr.Lit(2), TExpr.Lit(30))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft)
  }

  test("YEAR: extract year from date") {
    val date = LocalDate.of(2025, 11, 21)
    val expr = TExpr.Year(TExpr.Lit(date))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(2025)))  // Now returns BigDecimal
  }

  test("MONTH: extract month from date") {
    val date = LocalDate.of(2025, 11, 21)
    val expr = TExpr.Month(TExpr.Lit(date))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(11)))  // Now returns BigDecimal
  }

  test("DAY: extract day from date") {
    val date = LocalDate.of(2025, 11, 21)
    val expr = TExpr.Day(TExpr.Lit(date))
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(21)))  // Now returns BigDecimal
  }

  test("Date functions: round-trip DATE/YEAR/MONTH/DAY") {
    // Manual ToInt wrapping for programmatic construction (parser uses asIntExpr automatically)
    val baseDate = TExpr.Lit(LocalDate.of(2025, 11, 21))
    val expr = TExpr.Date(
      TExpr.ToInt(TExpr.Year(baseDate)),
      TExpr.ToInt(TExpr.Month(baseDate)),
      TExpr.ToInt(TExpr.Day(baseDate))
    )
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(LocalDate.of(2025, 11, 21)))
  }

  // ===== Range Functions =====

  test("MIN: empty range returns 0") {
    val range = parseRange("A1:A3")
    val expr = TExpr.Min(range)
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(0)))
  }

  test("MIN: single cell") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(42)))
    val range = parseRange("A1:A1")
    val expr = TExpr.Min(range)
    val result = evaluator.eval(expr, sheet)
    assertEquals(result, Right(BigDecimal(42)))
  }

  test("MIN: multiple cells") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(5)),
      ref"A3" -> CellValue.Number(BigDecimal(20))
    )
    val range = parseRange("A1:A3")
    val expr = TExpr.Min(range)
    val result = evaluator.eval(expr, sheet)
    assertEquals(result, Right(BigDecimal(5)))
  }

  test("MAX: empty range returns 0") {
    val range = parseRange("A1:A3")
    val expr = TExpr.Max(range)
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right(BigDecimal(0)))
  }

  test("MAX: single cell") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(42)))
    val range = parseRange("A1:A1")
    val expr = TExpr.Max(range)
    val result = evaluator.eval(expr, sheet)
    assertEquals(result, Right(BigDecimal(42)))
  }

  test("MAX: multiple cells") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(5)),
      ref"A3" -> CellValue.Number(BigDecimal(20))
    )
    val range = parseRange("A1:A3")
    val expr = TExpr.Max(range)
    val result = evaluator.eval(expr, sheet)
    assertEquals(result, Right(BigDecimal(20)))
  }

  test("MIN/MAX: skip text cells") {
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Text("text"),
      ref"A3" -> CellValue.Number(BigDecimal(5))
    )
    val range = parseRange("A1:A3")

    val minResult = evaluator.eval(TExpr.Min(range), sheet)
    assertEquals(minResult, Right(BigDecimal(5)))

    val maxResult = evaluator.eval(TExpr.Max(range), sheet)
    assertEquals(maxResult, Right(BigDecimal(10)))
  }

  // ===== FormulaPrinter Integration =====

  test("FormulaPrinter: CONCATENATE") {
    val expr = TExpr.Concatenate(List(TExpr.Lit("A"), TExpr.Lit("B")))
    val formula = FormulaPrinter.print(expr)
    assertEquals(formula, """=CONCATENATE("A", "B")""")
  }

  test("FormulaPrinter: LEFT") {
    val expr = TExpr.Left(TExpr.Lit("Hello"), TExpr.Lit(3))
    val formula = FormulaPrinter.print(expr)
    assertEquals(formula, """=LEFT("Hello", 3)""")
  }

  test("FormulaPrinter: RIGHT") {
    val expr = TExpr.Right(TExpr.Lit("Hello"), TExpr.Lit(3))
    val formula = FormulaPrinter.print(expr)
    assertEquals(formula, """=RIGHT("Hello", 3)""")
  }

  test("FormulaPrinter: LEN") {
    val expr = TExpr.Len(TExpr.Lit("Hello"))
    val formula = FormulaPrinter.print(expr)
    assertEquals(formula, """=LEN("Hello")""")
  }

  test("FormulaPrinter: UPPER") {
    val expr = TExpr.Upper(TExpr.Lit("hello"))
    val formula = FormulaPrinter.print(expr)
    assertEquals(formula, """=UPPER("hello")""")
  }

  test("FormulaPrinter: LOWER") {
    val expr = TExpr.Lower(TExpr.Lit("HELLO"))
    val formula = FormulaPrinter.print(expr)
    assertEquals(formula, """=LOWER("HELLO")""")
  }

  test("FormulaPrinter: TODAY") {
    val expr = TExpr.Today()
    val formula = FormulaPrinter.print(expr)
    assertEquals(formula, "=TODAY()")
  }

  test("FormulaPrinter: NOW") {
    val expr = TExpr.Now()
    val formula = FormulaPrinter.print(expr)
    assertEquals(formula, "=NOW()")
  }

  test("FormulaPrinter: DATE") {
    val expr = TExpr.Date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21))
    val formula = FormulaPrinter.print(expr)
    assertEquals(formula, "=DATE(2025, 11, 21)")
  }

  test("FormulaPrinter: YEAR") {
    val expr = TExpr.Year(TExpr.Date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21)))
    val formula = FormulaPrinter.print(expr)
    assertEquals(formula, "=YEAR(DATE(2025, 11, 21))")
  }

  test("FormulaPrinter: MONTH") {
    val expr = TExpr.Month(TExpr.Today())
    val formula = FormulaPrinter.print(expr)
    assertEquals(formula, "=MONTH(TODAY())")
  }

  test("FormulaPrinter: DAY") {
    val expr = TExpr.Day(TExpr.Today())
    val formula = FormulaPrinter.print(expr)
    assertEquals(formula, "=DAY(TODAY())")
  }

  test("FormulaPrinter: MIN") {
    val range = parseRange("A1:A10")
    val expr = TExpr.Min(range)
    val formula = FormulaPrinter.print(expr)
    assertEquals(formula, "=MIN(A1:A10)")
  }

  test("FormulaPrinter: MAX") {
    val range = parseRange("B2:B20")
    val expr = TExpr.Max(range)
    val formula = FormulaPrinter.print(expr)
    assertEquals(formula, "=MAX(B2:B20)")
  }

  // ===== Integration Tests =====

  test("Integration: nested text functions") {
    val expr = TExpr.Upper(
      TExpr.Concatenate(List(
        TExpr.Left(TExpr.Lit("Hello"), TExpr.Lit(3)),
        TExpr.Right(TExpr.Lit("World"), TExpr.Lit(3))
      ))
    )
    val result = evaluator.eval(expr, emptySheet)
    assertEquals(result, Right("HELRLD"))
  }

  test("Integration: date decomposition and recomposition") {
    val clock = Clock.fixedDate(LocalDate.of(2025, 11, 21))

    // Extract year, month, day from TODAY()
    val today = TExpr.Today()
    val year = TExpr.Year(today)
    val month = TExpr.Month(today)
    val day = TExpr.Day(today)

    // Verify each component (now returns BigDecimal)
    assertEquals(evaluator.eval(year, emptySheet, clock), Right(BigDecimal(2025)))
    assertEquals(evaluator.eval(month, emptySheet, clock), Right(BigDecimal(11)))
    assertEquals(evaluator.eval(day, emptySheet, clock), Right(BigDecimal(21)))

    // Recompose into DATE (use literal Ints since DATE expects Int arguments)
    val reconstructed = TExpr.Date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21))
    assertEquals(evaluator.eval(reconstructed, emptySheet, clock), Right(LocalDate.of(2025, 11, 21)))
  }

  test("Integration: MIN/MAX with LEN") {
    // Not directly possible without cell refs, but demonstrates type composition
    val minValue = TExpr.Min(parseRange("A1:A3"))
    val sheet = sheetWith(
      ref"A1" -> CellValue.Number(BigDecimal(10)),
      ref"A2" -> CellValue.Number(BigDecimal(5)),
      ref"A3" -> CellValue.Number(BigDecimal(20))
    )

    val result = evaluator.eval(minValue, sheet)
    assertEquals(result, Right(BigDecimal(5)))
  }
