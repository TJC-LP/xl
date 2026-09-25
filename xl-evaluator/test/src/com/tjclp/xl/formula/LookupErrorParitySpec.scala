package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.sheets.Sheet

/**
 * #670: Excel error-code parity for the lookup family after #662 — an error-typed lookup_value
 * propagates its own code, MATCH's match_type reads by sign, LOOKUP and XMATCH exist and raise the
 * typed `#N/A`, text keys match approximately (case-insensitive, over sorted text), and the
 * remaining host failures (INDEX out of bounds, XLOOKUP shape mismatch) are Excel error values.
 *
 * Expected values come from a LibreOffice 25.8 recalculation of the same grid (the uncached cells
 * of an xlsx computed on load, read back as CSV). Where LibreOffice and Excel's documented
 * behaviour part, the case says so and follows Excel.
 */
class LookupErrorParitySpec extends FunSuite:

  private val fruit = Vector("apple", "banana", "cherry", "grape", "melon")

  // The oracle grid: A1:A5 10..50, B1:B5 a..e, C1 #DIV/0! (cached), D1:D5 sorted fruit with E1:E5
  // 1..5, F1:F5 50..10, G1 blank, H1:H3 "Empty"/"x"/0, M1:Q2 a horizontal 1..5 over v..z, M3:Q4
  // horizontal fruit over 100..500, R1:R5 fruit descending, U1:U2 TRUE/FALSE
  private val sheet: Sheet =
    val base = (0 until 5).foldLeft(Sheet(SheetName.unsafe("S"))) { (s, i) =>
      s.put(ARef.from0(0, i), num(10 * (i + 1)))
        .put(ARef.from0(1, i), CellValue.Text(("a".charAt(0) + i).toChar.toString))
        .put(ARef.from0(3, i), CellValue.Text(fruit(i)))
        .put(ARef.from0(4, i), num(i + 1))
        .put(ARef.from0(5, i), num(10 * (5 - i)))
        .put(ARef.from0(12 + i, 0), num(i + 1))
        .put(ARef.from0(12 + i, 1), CellValue.Text("vwxyz".substring(i, i + 1)))
        .put(ARef.from0(12 + i, 2), CellValue.Text(fruit(i)))
        .put(ARef.from0(12 + i, 3), num(100 * (i + 1)))
        .put(ARef.from0(17, i), CellValue.Text(fruit(4 - i)))
    }
    base
      .put(ref"C1", CellValue.Formula("1/0", Some(CellValue.Error(CellError.Div0))))
      .put(ref"C2", CellValue.Error(CellError.Ref))
      .put(ref"H1", CellValue.Text("Empty"))
      .put(ref"H2", CellValue.Text("x"))
      .put(ref"H3", num(0))
      .put(ref"U1", CellValue.Bool(true))
      .put(ref"U2", CellValue.Bool(false))

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def text(s: String): CellValue = CellValue.Text(s)
  private def err(e: CellError): CellValue = CellValue.Error(e)

  private def check(formula: String, expected: CellValue)(implicit loc: munit.Location): Unit =
    assertEquals(sheet.evaluateFormula(formula), Right(expected), formula)

  private def leftOf(formula: String)(implicit loc: munit.Location): EvalError =
    FormulaParser.parse(formula) match
      case Right(expr) =>
        Evaluator.instance.eval(expr, sheet) match
          case Left(e) => e
          case Right(v) => fail(s"$formula: expected a Left, got $v")
      case Left(e) => fail(s"$formula: $e")

  // ===== (a) an error-typed lookup_value propagates its own code =====

  test("#670(a): VLOOKUP/HLOOKUP/MATCH/XLOOKUP over an error lookup_value return that error") {
    // LibreOffice: every row #DIV/0!
    check("=VLOOKUP(1/0,A1:B5,2,FALSE)", err(CellError.Div0))
    check("=VLOOKUP(1/0,A1:B5,2,TRUE)", err(CellError.Div0))
    check("=HLOOKUP(1/0,M1:Q2,2,FALSE)", err(CellError.Div0))
    check("=MATCH(1/0,A1:A5,0)", err(CellError.Div0))
    check("=MATCH(1/0,A1:A5,1)", err(CellError.Div0))
    check("=VLOOKUP(C1,A1:B5,2,FALSE)", err(CellError.Div0))
    check("=VLOOKUP(C1,A1:B5,2,TRUE)", err(CellError.Div0))
    check("=HLOOKUP(C1,M1:Q2,2,FALSE)", err(CellError.Div0))
    check("=MATCH(C1,A1:A5,0)", err(CellError.Div0))
    check("=MATCH(C1,A1:A5,-1)", err(CellError.Div0))
    check("=XLOOKUP(C1,A1:A5,B1:B5)", err(CellError.Div0))
    check("=VLOOKUP(NA(),A1:B5,2,FALSE)", err(CellError.NA))
  }

  test("#670(a): XLOOKUP's if_not_found does not catch an error lookup_value") {
    // Excel: #DIV/0! (LibreOffice answers Err:504 for the inline 1/0, #DIV/0! through C1)
    check("=XLOOKUP(C1,A1:A5,B1:B5,\"nf\")", err(CellError.Div0))
    check("=XLOOKUP(C2,A1:A5,B1:B5,\"nf\")", err(CellError.Ref))
  }

  test("#670(a): the propagated code is visible to ERROR.TYPE and not to IFNA") {
    check("=ERROR.TYPE(VLOOKUP(C1,A1:B5,2,FALSE))", num(2))
    check("=IFNA(MATCH(C2,A1:A5,0),\"miss\")", err(CellError.Ref))
  }

  // ===== (c) MATCH match_type reads by sign after truncation =====

  test("#670(c): MATCH match_type — any positive is 1, any negative is -1") {
    // Excel coerces by sign; LibreOffice answers Err:504 for 5 and -3 but agrees on 1.5 and -1.5
    check("=MATCH(25,A1:A5,5)", num(2))
    check("=MATCH(25,F1:F5,-3)", num(3))
    check("=MATCH(25,A1:A5,1.5)", num(2))
    check("=MATCH(25,F1:F5,-1.5)", num(3))
  }

  test("#670(c): a fractional match_type below 1 truncates to exact (LibreOffice: #N/A)") {
    check("=MATCH(25,A1:A5,0.5)", err(CellError.NA))
    check("=MATCH(25,A1:A5,-0.5)", err(CellError.NA))
    check("=MATCH(30,A1:A5,0.5)", num(3))
  }

  // ===== (d) LOOKUP =====

  test("#670(d): LOOKUP vector form — approximate over a sorted vector") {
    check("=LOOKUP(25,A1:A5,B1:B5)", text("b"))
    check("=LOOKUP(60,A1:A5,B1:B5)", text("e"))
    check("=LOOKUP(30,A1:A5)", num(30))
    check("=LOOKUP(2,M1:Q1,M2:Q2)", text("w"))
  }

  test("#670(d): LOOKUP result_vector may lie on the other axis") {
    check("=LOOKUP(25,A1:A5,M2:Q2)", text("w"))
  }

  test("#670(d): LOOKUP array form — first row or column by shape, result from the last") {
    check("=LOOKUP(25,A1:B5)", text("b"))
    check("=LOOKUP(3.5,M1:Q2)", text("x"))
    check("=LOOKUP(25,M1:Q4)", num(500))
  }

  test("#670(d): LOOKUP text keys compare case-insensitively over sorted text") {
    check("=LOOKUP(\"c\",D1:D5,E1:E5)", num(2))
    check("=LOOKUP(\"CHERRY\",D1:D5,E1:E5)", num(3))
    check("=LOOKUP(\"zzz\",D1:D5)", text("melon"))
  }

  test("#670(d): a LOOKUP miss is the typed #N/A through lookupNotFound") {
    check("=LOOKUP(5,A1:A5,B1:B5)", err(CellError.NA))
    check("=LOOKUP(\"aaa\",D1:D5,E1:E5)", err(CellError.NA))
    check("=LOOKUP(G1,A1:A5,B1:B5)", err(CellError.NA))
    check("=IFNA(LOOKUP(5,A1:A5,B1:B5),\"miss\")", text("miss"))
    leftOf("=LOOKUP(5,A1:A5,B1:B5)") match
      case EvalError.ErrorValue(CellError.NA, Some(ctx)) =>
        assert(ctx.contains("LOOKUP(5, A1:A5, B1:B5)"), ctx)
      case other => fail(s"expected ErrorValue(NA, ctx), got $other")
  }

  test("#670(d): LOOKUP over an error lookup_value returns that error") {
    check("=LOOKUP(1/0,A1:A5,B1:B5)", err(CellError.Div0))
  }

  // ===== (d) XMATCH =====

  test("#670(d): XMATCH exact (the default) and its miss") {
    check("=XMATCH(30,A1:A5)", num(3))
    check("=XMATCH(3,M1:Q1)", num(3))
    check("=XMATCH(\"CHERRY\",D1:D5)", num(3))
    check("=XMATCH(\"b*\",B1:B5)", err(CellError.NA))
    check("=XMATCH(25,A1:A5)", err(CellError.NA))
    check("=IFNA(XMATCH(25,A1:A5),\"miss\")", text("miss"))
    leftOf("=XMATCH(25,A1:A5)") match
      case EvalError.ErrorValue(CellError.NA, Some(ctx)) =>
        assert(ctx.contains("XMATCH(25, A1:A5, 0, 1)"), ctx)
      case other => fail(s"expected ErrorValue(NA, ctx), got $other")
  }

  test("#670(d): XMATCH match_mode -1 / 1 — exact or next smaller / larger, any order") {
    check("=XMATCH(25,A1:A5,-1)", num(2))
    check("=XMATCH(25,A1:A5,1)", num(3))
    check("=XMATCH(25,F1:F5,1)", num(3))
    check("=XMATCH(25,F1:F5,-1)", num(4))
    check("=XMATCH(5,A1:A5,-1)", err(CellError.NA))
    check("=XMATCH(55,A1:A5,1)", err(CellError.NA))
    check("=XMATCH(\"c\",D1:D5,-1)", num(2))
    check("=XMATCH(\"c\",D1:D5,1)", num(3))
  }

  test("#670(d): XMATCH match_mode 2 — wildcards") {
    check("=XMATCH(\"b*\",B1:B5,2)", num(2))
    check("=XMATCH(\"ch?rry\",D1:D5,2)", num(3))
    check("=XMATCH(\"*\",B1:B5,2)", num(1))
    check("=XMATCH(\"~*\",B1:B5,2)", err(CellError.NA))
  }

  test("#670(d): XMATCH search_mode — reverse and binary") {
    check("=XMATCH(30,A1:A5,0,-1)", num(3))
    check("=XMATCH(30,A1:A5,0,2)", num(3))
    check("=XMATCH(30,F1:F5,0,-2)", num(3))
    check("=XMATCH(25,A1:A5,-1,2)", num(2))
    check("=XMATCH(25,A1:A5,1,2)", num(3))
    check("=XMATCH(25,F1:F5,-1,-2)", num(4))
  }

  test("#670(d): XMATCH reverse search returns the LAST duplicate") {
    val dup = sheet
      .put(ref"W1", num(7))
      .put(ref"W2", num(8))
      .put(ref"W3", num(7))
    assertEquals(dup.evaluateFormula("=XMATCH(7,W1:W3)"), Right(num(1)))
    assertEquals(dup.evaluateFormula("=XMATCH(7,W1:W3,0,-1)"), Right(num(3)))
  }

  test("#670(d): XMATCH invalid modes and a 2-D lookup_array are #VALUE!") {
    // LibreOffice: Err:504 for search_mode 0 and the 2-D array; #N/A for match_mode 3 (Excel:
    // #VALUE!, the answer for any mode outside its table)
    check("=XMATCH(25,A1:A5,3)", err(CellError.Value))
    check("=XMATCH(25,A1:A5,0,0)", err(CellError.Value))
    check("=XMATCH(30,A1:B5)", err(CellError.Value))
  }

  test("#670(d): XMATCH over an error lookup_value returns that error; a blank matches nothing") {
    check("=XMATCH(1/0,A1:A5)", err(CellError.Div0))
    check("=XMATCH(G1,H1:H3)", err(CellError.NA))
    check("=XMATCH(TRUE,U1:U2)", num(1))
  }

  test("#670(d): XMATCH reads an empty optional slot as omitted") {
    check("=XMATCH(30,A1:A5,,)", num(3))
  }

  test("#670(d): XMATCH and LOOKUP round-trip through the printer") {
    List(
      "=XMATCH(30, A1:A5)",
      "=XMATCH(25, A1:A5, -1, 2)",
      "=LOOKUP(25, A1:A5, B1:B5)",
      "=LOOKUP(25, A1:B5)"
    ).foreach { f =>
      FormulaParser.parse(f) match
        case Right(expr) =>
          assertEquals(FormulaPrinter.print(expr), f)
          assertEquals(FormulaParser.parse(FormulaPrinter.print(expr)), Right(expr))
        case Left(e) => fail(s"$f: $e")
    }
  }

  // ===== (g) text approximate match =====

  test("#670(g): VLOOKUP/HLOOKUP range_lookup TRUE match text keys approximately") {
    check("=VLOOKUP(\"c\",D1:E5,2,TRUE)", num(2))
    check("=VLOOKUP(\"cherry\",D1:E5,2)", num(3))
    check("=VLOOKUP(\"CHERRY\",D1:E5,2,TRUE)", num(3))
    check("=VLOOKUP(\"Z\",D1:E5,2,TRUE)", num(5))
    check("=VLOOKUP(\"aaa\",D1:E5,2,TRUE)", err(CellError.NA))
    check("=HLOOKUP(\"c\",M3:Q4,2,TRUE)", num(200))
  }

  test("#670(g): MATCH 1 / -1 match text keys approximately") {
    check("=MATCH(\"c\",D1:D5,1)", num(2))
    check("=MATCH(\"c\",D1:D5)", num(2))
    check("=MATCH(\"c\",R1:R5,-1)", num(3))
    check("=MATCH(\"zzz\",R1:R5,-1)", err(CellError.NA))
  }

  test("#670(g): XLOOKUP next-smaller matches text keys") {
    check("=XLOOKUP(\"c\",D1:D5,E1:E5,,-1)", num(2))
    check("=XLOOKUP(25,A1:A5,B1:B5,,-1)", text("b"))
  }

  // ===== (h) and siblings: host failures become Excel error values =====

  test("#670: INDEX out of bounds is #REF!, a negative position #VALUE!") {
    // LibreOffice: #REF!, #REF!, Err:502 (its #VALUE! export); ERROR.TYPE(#REF!) is 4
    check("=INDEX(A1:A5,99)", err(CellError.Ref))
    check("=INDEX(A1:B5,1,3)", err(CellError.Ref))
    check("=INDEX(A1:A5,-1)", err(CellError.Value))
    check("=ERROR.TYPE(INDEX(A1:A5,99))", num(4))
    check("=IFERROR(INDEX(A1:A5,99),\"oob\")", text("oob"))
  }

  test("#670: XLOOKUP lookup/return shape mismatch is #VALUE!") {
    check("=XLOOKUP(30,A1:A5,B1:B4)", err(CellError.Value))
    check("=ERROR.TYPE(XLOOKUP(30,A1:A5,B1:B4))", num(3))
  }

  // ===== renderLookupValue =====

  test("#670: a blank lookup_value does not match the text \"Empty\" (nor 0)") {
    // LibreOffice: #N/A for VLOOKUP and MATCH over H1:H3 = "Empty", "x", 0
    check("=VLOOKUP(G1,H1:H3,1,FALSE)", err(CellError.NA))
    check("=HLOOKUP(G1,H1:H3,1,FALSE)", err(CellError.NA))
    check("=MATCH(G1,H1:H3,0)", err(CellError.NA))
  }

  test("#670: a blank lookup_value matches nothing in every lookup, lifted too") {
    // LibreOffice: #N/A for each, and SUMPRODUCT(--ISNA(MATCH(G1:G2,H1:H3,0))) is 2
    check("=XLOOKUP(G1,H1:H4,B1:B4)", err(CellError.NA))
    check("=XMATCH(G1,H1:H4)", err(CellError.NA))
    check("=MATCH(G1,H1:H4,0)", err(CellError.NA))
    check("=LOOKUP(G1,H1:H3)", err(CellError.NA))
    check("=SUMPRODUCT(--ISNA(MATCH(G1:G2,H1:H3,0)))", num(2))
    check("=SUMPRODUCT(--ISNA(XMATCH(G1:G2,H1:H3)))", num(2))
    // an explicit 0 still finds the 0 key
    check("=MATCH(0,H1:H4,0)", num(3))
  }

  test("#670(g): approximate matching never crosses text and numbers") {
    // LibreOffice: #N/A for both — the text "25" is not the number 25
    check("=VLOOKUP(\"25\",A1:B5,2,TRUE)", err(CellError.NA))
    check("=MATCH(\"25\",A1:A5,1)", err(CellError.NA))
  }

  test("#670(g): MATCH and LOOKUP land on the last of equal keys; XMATCH on the first") {
    // LibreOffice over {10,20,20,30}: MATCH(20,…,1) 3, MATCH(25,…,1) 3, LOOKUP(20,…) 3,
    // XMATCH(20,…,-1) 2, XMATCH(25,…,-1,-1) 3; descending {30,20,20,10}: MATCH(15,…,-1) 3
    val dup = sheet
      .put(ref"W1", num(10))
      .put(ref"W2", num(20))
      .put(ref"W3", num(20))
      .put(ref"W4", num(30))
      .put(ref"X1", num(30))
      .put(ref"X2", num(20))
      .put(ref"X3", num(20))
      .put(ref"X4", num(10))
    def dupCase(formula: String, expected: Int)(implicit loc: munit.Location): Unit =
      assertEquals(dup.evaluateFormula(formula), Right(num(expected)), formula)
    dupCase("=MATCH(20,W1:W4,1)", 3)
    dupCase("=MATCH(25,W1:W4,1)", 3)
    dupCase("=MATCH(15,X1:X4,-1)", 3)
    dupCase("=MATCH(20,X1:X4,-1)", 3)
    dupCase("=LOOKUP(20,W1:W4,E1:E4)", 3)
    dupCase("=XMATCH(20,W1:W4,-1)", 2)
    dupCase("=XMATCH(25,W1:W4,-1,-1)", 3)
    dupCase("=XMATCH(15,W1:W4,1)", 2)
  }

  test("#670(d): a LOOKUP position past the end of result_vector is #N/A") {
    // LibreOffice: LOOKUP(2,{1,2,3},{"a","b"}) is "b", LOOKUP(3,…) #N/A
    check("=LOOKUP(2,M1:O1,M2:N2)", text("w"))
    check("=LOOKUP(3,M1:O1,M2:N2)", err(CellError.NA))
  }

  test("#670: VLOOKUP exact-matches a logical key") {
    // Excel: TRUE (LibreOffice answers 1 — its logicals are numbers)
    check("=VLOOKUP(TRUE,U1:U2,1,FALSE)", CellValue.Bool(true))
    check("=MATCH(FALSE,U1:U2,0)", num(2))
  }

  test("#670: ERROR.TYPE classifies XMATCH's #VALUE!") {
    check("=ERROR.TYPE(XMATCH(25,A1:A5,3))", num(3))
  }

  test("#670/#681: lookup diagnostics render booleans and range_lookup as TRUE/FALSE") {
    leftOf("=VLOOKUP(TRUE,A1:B5,2,FALSE)") match
      case EvalError.ErrorValue(CellError.NA, Some(ctx)) =>
        assertEquals(ctx, "VLOOKUP exact match not found: VLOOKUP(TRUE, A1:B5, 2, FALSE)")
      case other => fail(s"expected ErrorValue(NA, ctx), got $other")
    leftOf("=HLOOKUP(0,M1:Q2,2,TRUE)") match
      case EvalError.ErrorValue(CellError.NA, Some(ctx)) =>
        assertEquals(ctx, "HLOOKUP approximate match not found: HLOOKUP(0, M1:Q2, 2, TRUE)")
      case other => fail(s"expected ErrorValue(NA, ctx), got $other")
  }
