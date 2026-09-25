package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-476: the parity gaps that fail LOUD — host errors that abort whole-book recalc on ordinary
 * banker files, as distinct from the silent-wrong-value class (#467/#488).
 *
 *   - SEARCH absent (FIND exists but is case-sensitive; SEARCH is the banker idiom)
 *   - N() and HYPERLINK() absent → UnknownFunction
 *   - `=IF(1=1,+S2!G1,0)`: unary plus on a cross-sheet ref inside a function argument escaped
 *     PolyRef resolution and hit the "Unresolved SheetPolyRef" programming-error arm
 *   - VALUE("None") raised a host error where Excel caches #VALUE!
 */
class LoudParityGapsSpec extends FunSuite:

  private def sheetWith(name: String, cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(Sheet(SheetName.unsafe(name))) { case (s, (ref, v)) => s.put(ref, v) }

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  // A1 = "Widget Corporation": o at 9 and 12, so start_num is observable
  private val text = Sheet("T")
    .put(ARef.from0(0, 0), CellValue.Text("Widget Corporation"))
    .put(ARef.from0(0, 1), CellValue.Number(BigDecimal("42.5")))
    .put(ARef.from0(0, 2), CellValue.Bool(true))

  // ========== SEARCH ==========

  test("GH-476: SEARCH is case-insensitive (where FIND is not)") {
    assertEquals(text.evaluateFormula("""=SEARCH("widget", A1)"""), Right(num(1)))
    assertEquals(text.evaluateFormula("""=SEARCH("CORP", A1)"""), Right(num(8)))
  }

  test("GH-476: FIND stays case-sensitive") {
    text.evaluateFormula("""=FIND("widget", A1)""") match
      case Right(CellValue.Error(CellError.Value)) => ()
      case Left(_) => ()
      case other => fail(s"FIND must not match 'widget' in 'Widget Corp', got $other")
  }

  test("GH-476: SEARCH honours start_num") {
    assertEquals(text.evaluateFormula("""=SEARCH("o", A1)"""), Right(num(9)))
    assertEquals(text.evaluateFormula("""=SEARCH("o", A1, 10)"""), Right(num(12)))
  }

  test("GH-476: SEARCH start_num past the end is a cached #VALUE!") {
    assertEquals(
      text.evaluateFormula("""=SEARCH("o", A1, 99)"""),
      Right(CellValue.Error(CellError.Value))
    )
  }

  test("GH-476: SEARCH supports Excel wildcards") {
    assertEquals(text.evaluateFormula("""=SEARCH("W?dget", A1)"""), Right(num(1)))
    assertEquals(text.evaluateFormula("""=SEARCH("Wid*rp", A1)"""), Right(num(1)))
  }

  test("GH-476: SEARCH miss is a cached #VALUE!, not a host error") {
    assertEquals(
      text.evaluateFormula("""=SEARCH("zzz", A1)"""),
      Right(CellValue.Error(CellError.Value))
    )
  }

  test("GH-476: SEARCH composes with the IFERROR/MID idiom") {
    assertEquals(
      text.evaluateFormula("""=IFERROR(SEARCH("zzz", A1), 0)"""),
      Right(num(0))
    )
  }

  // ========== N ==========

  test("GH-476: N() coerces per Excel's table") {
    assertEquals(text.evaluateFormula("=N(5)"), Right(num(5)))
    assertEquals(text.evaluateFormula("=N(TRUE)"), Right(num(1)))
    assertEquals(text.evaluateFormula("=N(FALSE)"), Right(num(0)))
    assertEquals(text.evaluateFormula("""=N("hello")"""), Right(num(0)))
    assertEquals(text.evaluateFormula("=N(A1)"), Right(num(0)))
    assertEquals(text.evaluateFormula("=N(A2)"), Right(CellValue.Number(BigDecimal("42.5"))))
    assertEquals(text.evaluateFormula("=N(A3)"), Right(num(1)))
    assertEquals(text.evaluateFormula("=N(Z9)"), Right(num(0)))
  }

  test("GH-476: N() of a date is its Excel serial") {
    assertEquals(text.evaluateFormula("=N(DATE(2026,1,1))"), Right(num(46023)))
  }

  test("GH-476: N() propagates direct and referenced Excel errors") {
    val errors = Sheet("Errors")
      .put(ref"A1", CellValue.Error(CellError.Div0))
      .put(ref"A2", CellValue.Error(CellError.NA))

    assertEquals(errors.evaluateFormula("=N(1/0)"), Right(CellValue.Error(CellError.Div0)))
    assertEquals(errors.evaluateFormula("=N(A1)"), Right(CellValue.Error(CellError.Div0)))
    assertEquals(errors.evaluateFormula("=N(A2)"), Right(CellValue.Error(CellError.NA)))
  }

  test("GH-476: N() propagates an error cached by a referenced formula") {
    val errors = Sheet("Errors")
      .put(
        ref"A1",
        CellValue.Formula("=1/0", Some(CellValue.Error(CellError.Div0)))
      )

    assertEquals(errors.evaluateFormula("=N(A1)"), Right(CellValue.Error(CellError.Div0)))
  }

  // ========== IFNA / NA ==========

  test("GH-511: IFNA catches #N/A on the value channel and the Left channel alike") {
    val errors = Sheet("Errors")
      .put(ref"A1", CellValue.Error(CellError.NA))
      .put(ref"A2", CellValue.Formula("=NA()", Some(CellValue.Error(CellError.NA))))

    assertEquals(errors.evaluateFormula("=IFNA(A1,42)"), Right(num(42)))
    assertEquals(errors.evaluateFormula("=IFNA(A2,42)"), Right(num(42)))
    assertEquals(errors.evaluateFormula("=IFNA(NA(),42)"), Right(num(42)))
  }

  test("GH-511: IFNA does NOT swallow other errors — that is the point of it") {
    // The whole reason a banker book guards a lookup with IFNA rather than IFERROR: a missing key
    // is expected, a division by zero is a bug, and one guard must not hide both.
    val errors = Sheet("Errors").put(ref"A1", CellValue.Error(CellError.Div0))

    assertEquals(errors.evaluateFormula("=IFNA(A1,42)"), Right(CellValue.Error(CellError.Div0)))
    assertEquals(errors.evaluateFormula("=IFNA(1/0,42)"), Right(CellValue.Error(CellError.Div0)))
    assertEquals(errors.evaluateFormula("=IFERROR(A1,42)"), Right(num(42)), "IFERROR still catches")
  }

  test("GH-511: IFNA passes a healthy value straight through") {
    val healthy = Sheet("H").put(ref"A1", num(7))
    assertEquals(healthy.evaluateFormula("=IFNA(A1,42)"), Right(num(7)))
    assertEquals(healthy.evaluateFormula("=IFNA(A1*2,42)"), Right(num(14)))
  }

  test("GH-512: the error guards see an error CACHED inside a formula cell") {
    // The shape every recalculated book is made of: B1 is a formula whose cached value is the
    // error. Matching only a bare CellValue.Error made ISERROR answer FALSE and IFERROR hand back
    // the formula cell instead of the fallback — a silently wrong guard on the most ordinary input
    // there is. All five now share `ArrayArithmetic.carriedError`, the matcher the aggregate
    // guards already used.
    val cached = Sheet("C")
      .put(ref"A1", CellValue.Formula("=1/0", Some(CellValue.Error(CellError.Div0))))
      .put(ref"A2", CellValue.Formula("=NA()", Some(CellValue.Error(CellError.NA))))
      .put(ref"A3", CellValue.Formula("=1+1", Some(num(2))))

    assertEquals(cached.evaluateFormula("=ISERROR(A1)"), Right(CellValue.Bool(true)))
    assertEquals(cached.evaluateFormula("=IFERROR(A1,42)"), Right(num(42)))
    assertEquals(cached.evaluateFormula("=ISERR(A1)"), Right(CellValue.Bool(true)))
    assertEquals(
      cached.evaluateFormula("=ISERR(A2)"),
      Right(CellValue.Bool(false)),
      "#N/A excluded"
    )
    assertEquals(cached.evaluateFormula("=ISNA(A2)"), Right(CellValue.Bool(true)))
    assertEquals(cached.evaluateFormula("=IFNA(A2,42)"), Right(num(42)))
    // A healthy cached formula is not an error under any of them.
    assertEquals(cached.evaluateFormula("=ISERROR(A3)"), Right(CellValue.Bool(false)))
    assertEquals(cached.evaluateFormula("=ISNA(A3)"), Right(CellValue.Bool(false)))
  }

  test("GH-511: NA() authors the #N/A literal") {
    assertEquals(text.evaluateFormula("=NA()"), Right(CellValue.Error(CellError.NA)))
    assertEquals(text.evaluateFormula("=ISNA(NA())"), Right(CellValue.Bool(true)))
    assertEquals(text.evaluateFormula("=ISERR(NA())"), Right(CellValue.Bool(false)))
    assertEquals(text.evaluateFormula("=ISERROR(NA())"), Right(CellValue.Bool(true)))
  }

  // ========== HYPERLINK ==========

  test("GH-476: HYPERLINK displays the friendly name when given") {
    assertEquals(
      text.evaluateFormula("""=HYPERLINK("https://x.test","Open")"""),
      Right(CellValue.Text("Open"))
    )
  }

  test("GH-476: HYPERLINK with no friendly name displays the link") {
    assertEquals(
      text.evaluateFormula("""=HYPERLINK("https://x.test")"""),
      Right(CellValue.Text("https://x.test"))
    )
  }

  // ========== unary plus on a cross-sheet ref inside a function argument ==========

  private val book: Workbook =
    Workbook(
      Vector(
        sheetWith("Main", ARef.from0(0, 0) -> num(1)),
        sheetWith("S2", ARef.from0(6, 0) -> num(11)) // G1 = 11
      )
    )

  private def main: Sheet = book.sheets.headOption.getOrElse(fail("no Main sheet"))

  test("GH-476: +Sheet!Ref inside a function argument resolves (was Unresolved SheetPolyRef)") {
    assertEquals(
      main.evaluateFormula("=IF(1=1,+S2!G1,0)", workbook = Some(book)),
      Right(num(11))
    )
  }

  test("GH-476: bare +Sheet!Ref and arithmetic over it keep working") {
    assertEquals(main.evaluateFormula("=+S2!G1", workbook = Some(book)), Right(num(11)))
    assertEquals(main.evaluateFormula("=+S2!G1*2", workbook = Some(book)), Right(num(22)))
  }

  test("GH-476: +Ref (same sheet) inside a function argument resolves too") {
    val s = Sheet("Main").put(ARef.from0(0, 0), num(7))
    assertEquals(s.evaluateFormula("=IF(1=1,+A1,0)"), Right(num(7)))
    assertEquals(s.evaluateFormula("=SUM(+A1,1)"), Right(num(8)))
  }

  // ========== VALUE error parity ==========

  test("GH-476: VALUE of unparseable text is a cached #VALUE!, not a host error") {
    assertEquals(
      text.evaluateFormula("""=VALUE("None")"""),
      Right(CellValue.Error(CellError.Value))
    )
  }

  test("GH-476: VALUE's #VALUE! is catchable by IFERROR") {
    assertEquals(text.evaluateFormula("""=IFERROR(VALUE("None"), 0)"""), Right(num(0)))
  }

  test("GH-476: VALUE still parses the Excel-numeric-text forms") {
    assertEquals(text.evaluateFormula("""=VALUE("$1,234")"""), Right(num(1234)))
    assertEquals(
      text.evaluateFormula("""=VALUE("(500)")"""),
      Right(CellValue.Number(BigDecimal(-500)))
    )
  }

  // ========== GH-662: legacy lookup misses are the typed #N/A ==========

  /**
   * A3:A7 = 2021..2025 keyed to B3:B7 = 190..238 (VLOOKUP/MATCH); D1:H1 = 2021..2025 over D2:H2 =
   * 190..238 (HLOOKUP); C3:C7 = 2025..2021, the descending keys MATCH -1 wants.
   */
  private val lookups: Sheet =
    val years = Vector(2021, 2022, 2023, 2024, 2025)
    val values = Vector(190, 202, 214, 226, 238)
    years.indices.foldLeft(Sheet("L")) { (s, i) =>
      s.put(ARef.from0(0, 2 + i), num(years(i)))
        .put(ARef.from0(1, 2 + i), num(values(i)))
        .put(ARef.from0(2, 2 + i), num(years(4 - i)))
        .put(ARef.from0(3 + i, 0), num(years(i)))
        .put(ARef.from0(3 + i, 1), num(values(i)))
    }

  /** Every legacy-lookup miss shape: exact and approximate, both axes, MATCH through INDEX. */
  private val misses: List[String] = List(
    "VLOOKUP(2030,A3:B7,2,FALSE)",
    "VLOOKUP(2000,A3:B7,2,TRUE)",
    "HLOOKUP(2030,D1:H2,2,FALSE)",
    "HLOOKUP(2000,D1:H2,2,TRUE)",
    "MATCH(2030,A3:A7,0)",
    "MATCH(2000,A3:A7,1)",
    "MATCH(3000,C3:C7,-1)",
    "INDEX(B3:B7,MATCH(2030,A3:A7,0))"
  )

  private val na: XLResult[CellValue] = Right(CellValue.Error(CellError.NA))

  private def lookupCase(formula: String, expected: XLResult[CellValue]): Unit =
    assertEquals(lookups.evaluateFormula(s"=$formula"), expected, formula)

  test("GH-662: a bare lookup miss is the cached #N/A, not a host failure") {
    misses.foreach(p => lookupCase(p, na))
  }

  test("GH-662: IFNA sees every lookup miss") {
    misses.foreach(p => lookupCase(s"IFNA($p,0)", Right(num(0))))
  }

  test("GH-662: ISNA is TRUE over every lookup miss") {
    misses.foreach(p => lookupCase(s"ISNA($p)", Right(CellValue.Bool(true))))
  }

  test("GH-662: ISERR is FALSE over a lookup miss — it is #N/A, the one code ISERR excludes") {
    misses.foreach(p => lookupCase(s"ISERR($p)", Right(CellValue.Bool(false))))
  }

  test("GH-662: IFERROR and ISERROR keep trapping the miss") {
    misses.foreach { p =>
      lookupCase(s"IFERROR($p,0)", Right(num(0)))
      lookupCase(s"ISERROR($p)", Right(CellValue.Bool(true)))
    }
  }

  test("GH-662: ERROR.TYPE of a lookup miss is 7") {
    misses.foreach(p => lookupCase(s"ERROR.TYPE($p)", Right(num(7))))
  }

  test("GH-662: the dogfood cell — IF(ISNA(miss),0,1) is 0, never a cached 1") {
    misses.foreach(p => lookupCase(s"IF(ISNA($p),0,1)", Right(num(0))))
  }

  test("GH-662: an unguarded miss absorbs through arithmetic and aggregates as #N/A") {
    misses.foreach { p =>
      lookupCase(s"$p+1", na)
      lookupCase(s"SUM($p)", na)
    }
  }

  test("GH-662: LET binds the miss as the #N/A value, so IFNA over the binding fires") {
    misses.foreach(p => lookupCase(s"LET(v,$p,IFNA(v,0))", Right(num(0))))
  }

  test("GH-662: a miss in an IF condition stays fatal and promotes at the boundary") {
    lookupCase("IF(VLOOKUP(2030,A3:B7,2,FALSE)>0,1,2)", na)
  }

  test("GH-662: hits are untouched") {
    lookupCase("VLOOKUP(2023,A3:B7,2,FALSE)", Right(num(214)))
    lookupCase("HLOOKUP(2023,D1:H2,2,FALSE)", Right(num(214)))
    lookupCase("MATCH(2023,A3:A7,0)", Right(num(3)))
    lookupCase("INDEX(B3:B7,MATCH(2023,A3:A7,0))", Right(num(214)))
  }

  test("GH-662: the miss travels the Left channel with the diagnostic as its context") {
    def leftOf(formula: String): EvalError =
      FormulaParser.parse(formula) match
        case Right(expr) =>
          Evaluator.instance.eval(expr, lookups) match
            case Left(err) => err
            case Right(v) => fail(s"$formula: expected a Left, got $v")
        case Left(err) => fail(s"$formula: $err")
    leftOf("=VLOOKUP(2030,A3:B7,2,FALSE)") match
      case EvalError.ErrorValue(CellError.NA, Some(ctx)) =>
        assertEquals(ctx, "VLOOKUP exact match not found: VLOOKUP(2030, A3:B7, 2, FALSE)")
      case other => fail(s"expected ErrorValue(NA, ctx), got $other")
    leftOf("=VLOOKUP(2000,A3:B7,2,TRUE)") match
      case EvalError.ErrorValue(CellError.NA, Some(ctx)) =>
        assertEquals(ctx, "VLOOKUP approximate match not found: VLOOKUP(2000, A3:B7, 2, TRUE)")
      case other => fail(s"expected ErrorValue(NA, ctx), got $other")
    leftOf("=HLOOKUP(2030,D1:H2,2,FALSE)") match
      case EvalError.ErrorValue(CellError.NA, Some(ctx)) =>
        assertEquals(ctx, "HLOOKUP exact match not found: HLOOKUP(2030, D1:H2, 2, FALSE)")
      case other => fail(s"expected ErrorValue(NA, ctx), got $other")
    leftOf("=MATCH(2030,A3:A7,0)") match
      case EvalError.ErrorValue(CellError.NA, Some(ctx)) =>
        assert(ctx.contains("no match found"), ctx)
        assert(ctx.contains("MATCH(2030, A3:A7, 0)"), ctx)
      case other => fail(s"expected ErrorValue(NA, ctx), got $other")
  }

  test("GH-662: array carriage — a miss inside a broadcast IF is #N/A per element, not #VALUE!") {
    val (out, spilled) = lookups
      .evaluateArrayFormula("=IF(A3:A4>0,VLOOKUP(2030,A3:B7,2,FALSE),0)", ARef.from0(9, 0))
      .fold(e => fail(e.toString), identity)
    assertEquals(spilled.height, 2)
    spilled.cells.foreach { r =>
      assertEquals(out(r).value, CellValue.Error(CellError.NA), r.toA1)
    }
  }

  test("GH-662: a bare miss as an array formula spills a 1×1 #N/A at the origin") {
    val origin = ARef.from0(9, 5)
    val (out, spilled) = lookups
      .evaluateArrayFormula("=VLOOKUP(2030,A3:B7,2,FALSE)", origin)
      .fold(e => fail(e.toString), identity)
    assertEquals(spilled, CellRange(origin, origin))
    assertEquals(out(origin).value, CellValue.Error(CellError.NA))
  }
