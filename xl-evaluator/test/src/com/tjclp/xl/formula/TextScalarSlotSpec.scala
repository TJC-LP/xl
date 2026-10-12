package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellError, CellValue}
import com.tjclp.xl.codec.CodecError
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{ArrayArithmetic, EvalError, Evaluator, ScalarCoercion}
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

import scala.concurrent.{ExecutionContext, Future}
import scala.concurrent.duration.*

/**
 * #709: a text value in a typed scalar slot — a numeric, integer or date argument, the IF/IFS
 * condition, NOT's operand, a computed AND/OR value — is Excel's `#VALUE!`, an error value the cell
 * caches and IFERROR/ISERROR/ERROR.TYPE see, never a host failure that leaves the cell uncached.
 * Numeric text keeps Excel's rule: `"5"` is 5 in a numeric, integer or date slot, and `#VALUE!` in
 * a logical one.
 *
 * The table is LibreOffice 24.2's recalculation of the same plain `<f>` cells (the cell each
 * formula sits in is part of the case). The rows where Excel and LibreOffice differ, and the shapes
 * LibreOffice cannot compute, are pinned after it, each with its reason.
 */
class TextScalarSlotSpec extends FunSuite:

  // the magnitude tests run as Futures so a regression that builds a huge BigInteger fails here
  // instead of hanging the suite (munit applies the timeout to Future-valued tests only)
  override val munitTimeout: Duration = 20.seconds

  private def N(n: String): CellValue = CellValue.Number(BigDecimal(n))
  private def T(s: String): CellValue = CellValue.Text(s)
  private def B(b: Boolean): CellValue = CellValue.Bool(b)
  private val Value: CellValue = CellValue.Error(CellError.Value)

  private val a = Vector("1", "-2", "3.5", "0", "10", "7", "-4.25", "100", "2", "5")
  private val c = Vector(
    "apple",
    "banana",
    "cherry",
    "date",
    "elder",
    "fig",
    "grape",
    "honeydew",
    "iceberg",
    "jam"
  )
  private val g: Vector[CellValue] =
    Vector(N("3"), N("-4"), N("5"), T("n/a"), N("-6"), N("7"), N("8"), N("-9"), N("10"), N("11"))

  /**
   * Sheet1: A1:A10 numbers, B1:B10 = 2, C1:C10 text, G1:G10 numbers with the placeholder "n/a" in
   * G4, F1 "5", F2 "2024-01-15", F3 "TRUE", F4 " 5 " (all text).
   */
  private val book: Workbook =
    val sheet1 = (0 until 10).foldLeft(Sheet(SheetName.unsafe("Sheet1"))) { (s, i) =>
      s.put(ARef.from0(0, i), N(a(i)))
        .put(ARef.from0(1, i), N("2"))
        .put(ARef.from0(2, i), T(c(i)))
        .put(ARef.from0(6, i), g(i))
    }
    val withCells = sheet1
      .put(ref"F1", T("5"))
      .put(ref"F2", T("2024-01-15"))
      .put(ref"F3", T("TRUE"))
      .put(ref"F4", T(" 5 "))
    Workbook(Vector(withCells))

  /** `formula` as a plain formula cell at `Sheet1!at`, evaluated at its own position. */
  private def plain(at: String, formula: String): Either[String, CellValue] =
    plainIn(book, at, formula)

  private def plainIn(wb: Workbook, at: String, formula: String): Either[String, CellValue] =
    val ref = ARef.parse(at).fold(msg => fail(msg), identity)
    wb.sheets.headOption match
      case None => Left("no sheet")
      case Some(target) =>
        val placed = target.put(ref, CellValue.Formula(formula, None))
        placed.evaluateCell(ref, Clock.system, Some(wb.put(placed))).left.map(_.message)

  private def agree(at: String, formula: String, expected: CellValue): Unit =
    assertEquals(plain(at, formula), Right(expected), s"$at: =$formula")

  private val libreOffice: List[(String, String, CellValue)] = List(
    // the issue's table
    ("J1", "ABS(C5)", Value),
    ("J2", "IF(C6,1,0)", Value),
    ("J3", "NOT(C7)", Value),
    ("J4", "ROUND(C8,0)", Value),
    ("J5", "IFERROR(ABS(C9),\"caught\")", T("caught")),
    // integer, date and numeric slots over a text cell
    ("J6", "DATE(2024,C5,1)", Value),
    ("J7", "YEAR(C5)", Value),
    ("J8", "MONTH(C5)", Value),
    ("J9", "DAY(C1)", Value),
    ("J10", "SQRT(C5)", Value),
    ("J11", "LEFT(\"abc\",C5)", Value),
    ("J12", "MID(\"abcdef\",C1,2)", Value),
    ("J13", "ROUND(1.5,C1)", Value),
    ("J14", "POWER(C1,2)", Value),
    ("J15", "MOD(C1,2)", Value),
    ("J16", "EDATE(DATE(2024,1,1),C5)", Value),
    ("J17", "EDATE(C1,1)", Value),
    ("J18", "EOMONTH(C1,0)", Value),
    ("J19", "CHOOSE(C5,1,2)", Value),
    // the operators over a single text cell
    ("J20", "C5+0", Value),
    ("J21", "C1*1", Value),
    ("J22", "-C1", Value),
    ("J23", "C1^2", Value),
    ("J24", "C1%", Value),
    // the error value is visible to the error functions
    ("J25", "ISERROR(ABS(C5))", B(true)),
    ("J26", "ERROR.TYPE(ABS(C5))", N("3")),
    ("J27", "ERROR.TYPE(IF(C6,1,0))", N("3")),
    ("J28", "SUM(ABS(C5))", Value),
    // literals, computed text and references INDIRECT or a sheet name resolve
    ("J29", "ABS(\"abc\")", Value),
    ("J30", "IF(\"abc\",1,0)", Value),
    ("J31", "NOT(\"abc\")", Value),
    ("J32", "ROUND(\"abc\",0)", Value),
    ("J33", "YEAR(\"abc\")", Value),
    ("J34", "ABS(\"\")", Value),
    ("J35", "IF(\"\",1,0)", Value),
    ("J36", "YEAR(\"\")", Value),
    ("J37", "DATE(2024,\"\",1)", Value),
    ("J38", "LEFT(\"abc\",\"x\")", Value),
    ("J39", "ABS(C5&\"\")", Value),
    ("J40", "IF(C5&\"\",1,0)", Value),
    ("J41", "ABS(INDIRECT(\"C5\"))", Value),
    ("J42", "IF(INDIRECT(\"C5\"),1,0)", Value),
    ("J43", "NOT(INDIRECT(\"C5\"))", Value),
    ("J44", "ABS(Sheet1!C5)", Value),
    ("J45", "IF(Sheet1!C6,1,0)", Value),
    // computed AND/OR values: a text value is #VALUE! (a reference's text is ignored)
    ("J46", "AND(C1&\"\")", Value),
    ("J47", "OR(C1&\"\")", Value),
    ("J48", "AND(TRUE,\"abc\")", Value),
    ("J49", "OR(FALSE,\"abc\")", Value),
    ("J50", "AND({\"a\",TRUE})", Value),
    ("J51", "OR({\"a\",FALSE})", Value),
    ("J52", "AND(IF({1,0},\"a\",TRUE))", Value),
    ("J53", "SUMPRODUCT(--AND(C1:C2&\"\"))", Value),
    ("J54", "AND(C1)", Value),
    ("J55", "AND(C1,TRUE)", B(true)),
    // numeric text is its number in a numeric, integer or date slot
    ("J56", "ABS(F1)", N("5")),
    ("J57", "F1+0", N("5")),
    ("J58", "ABS(F4)", N("5")),
    ("J59", "MONTH(DATE(2024,F1,1))", N("5")),
    ("J60", "YEAR(F1)", N("1900")),
    ("J61", "YEAR(\"45000\")", N("2023")),
    ("J62", "ABS(\"5\")", N("5")),
    ("J63", "IF(F3,1,0)", N("1")),
    // a plain cell intersects a column: the text row is #VALUE!, the others their value
    ("H3", "ABS(G1:G10)", N("5")),
    ("H4", "ABS(G1:G10)", Value),
    ("H5", "ABS(G1:G10)", N("6")),
    ("I1", "IF(C1:C10,1,0)", Value),
    ("I3", "MONTH(DATE(2024,G1:G10,1))", N("5")),
    ("I4", "DATE(2024,G1:G10,1)", Value),
    ("I5", "ROUND(G1:G10,0)", N("-6")),
    ("I6", "NOT(G1:G10)", B(false))
  )

  test("#709: a text value in a typed scalar slot is #VALUE!, as LibreOffice computes it") {
    val diffs = libreOffice.flatMap { case (at, formula, expected) =>
      val got = plain(at, formula)
      Option.when(got != Right(expected))(s"$at =$formula: expected $expected, got $got")
    }
    assert(diffs.isEmpty, diffs.mkString("\n"))
  }

  test("#709: Excel's logical slots refuse numeric and boolean-looking text where LO converts") {
    // Excel converts only the text TRUE/FALSE in a logical slot and never converts text to a number
    // there; LibreOffice reads "5" as 5 (IF and NOT give 1 and FALSE). Likewise "TRUE" is not a
    // number to Excel (`="TRUE"+0` is #VALUE!) where LibreOffice reads it as 1.
    agree("K1", "IF(F1,1,0)", Value)
    agree("K2", "NOT(\"5\")", Value)
    agree("K3", "IF(\"5\",1,0)", Value)
    agree("K4", "ABS(F3)", Value)
  }

  test("#709: IFS and LET carry the same #VALUE! (LibreOffice 24.2 lacks both)") {
    agree("K5", "IFS(C5,1,TRUE,2)", Value)
    agree("K6", "LET(x,C5,ABS(x))", Value)
    agree("K7", "LET(x,C5,IF(x,1,0))", Value)
    agree("K8", "IFERROR(LET(x,C5,ABS(x)),\"caught\")", T("caught"))
  }

  test("#709: host failures in a typed slot stay loud") {
    assert(plain("K9", "ABS(Missing!A1)").isLeft, "a missing sheet is not #VALUE!")
    assert(plain("K10", "IF(Missing!A1,1,0)").isLeft, "a missing sheet is not #VALUE!")
  }

  test("#709: range folds over date cells keep refusing numeric text") {
    // Only a single reference in a date slot reads numeric text as a serial; a holiday list or an
    // XIRR/XNPV dates range still refuses it (LibreOffice 24.2: NETWORKDAYS 22; Excel documents
    // #VALUE! for an invalid XIRR/XNPV date).
    val sheet = Sheet(SheetName.unsafe("Sheet1"))
      .put(ref"M1", N("44927"))
      .put(ref"M2", N("44957"))
      .put(ref"M3", T("44935"))
      .put(ref"N1", N("-1000"))
      .put(ref"N2", N("300"))
      .put(ref"N3", N("800"))
      .put(ref"O1", N("44927"))
      .put(ref"O2", T("45000"))
      .put(ref"O3", N("45100"))
    val wb = Workbook(Vector(sheet))
    val cases = List(
      ("P1", "NETWORKDAYS(M1,M2,M3)", N("22")),
      ("P2", "XIRR(N1:N3,O1:O3)", Value),
      ("P3", "XNPV(0.1,N1:N3,O1:O3)", Value)
    )
    val diffs = cases.flatMap { case (at, formula, expected) =>
      val got = plainIn(wb, at, formula)
      Option.when(got != Right(expected))(s"$at =$formula: expected $expected, got $got")
    }
    assert(diffs.isEmpty, diffs.mkString("\n"))
  }

  test("#709: an aggregate over a single numeric-text reference ignores it; ABS reads it") {
    // Excel's SUM/COUNT skip text that arrives through a reference (only a literal like SUM("5")
    // converts), while a scalar numeric slot reads '5 as 5. The aggregates route a bare reference
    // through their own cell triage before the looser scalar decoder can see it.
    agree("L1", "SUM(F1)", N("0"))
    agree("L2", "COUNT(F1)", N("0"))
    agree("L3", "ABS(F1)", N("5"))
  }

  test("#709: a rich-text condition element is read through its plain text") {
    val rich = (s: String) => CellValue.RichText(RichText.plain(s))
    assertEquals(ArrayArithmetic.conditionTruthy("IF", rich("TRUE")), Right(true))
    assertEquals(ArrayArithmetic.conditionTruthy("IF", rich("false")), Right(false))
    ArrayArithmetic.conditionTruthy("IF", rich("abc")) match
      case Left(EvalError.ErrorValue(CellError.Value, _)) => ()
      case other => fail(s"expected #VALUE!, got $other")
  }

  test("#709: the #VALUE! diagnostic quotes at most 40 characters of the cell's text") {
    val long = "x" * 200
    val sheet = Sheet(SheetName.unsafe("Sheet1")).put(ref"A1", T(long))
    Evaluator.instance.eval(TExpr.ref(ref"A1", TExpr.decodeNumeric), sheet) match
      case Left(EvalError.ErrorValue(CellError.Value, Some(msg))) =>
        assert(msg.contains(s"'${"x" * 40}…'"), msg)
        assert(!msg.contains("x" * 41), msg)
      case other => fail(s"expected #VALUE! with a message, got $other")
    Evaluator.instance.eval(
      TExpr.ref(ref"A1", TExpr.decodeNumeric),
      sheet.put(ref"A1", T("abc"))
    ) match
      case Left(EvalError.ErrorValue(CellError.Value, Some(msg))) =>
        assert(msg.contains("'abc'"), msg)
      case other => fail(s"expected #VALUE! with a message, got $other")
  }

  test("#709: rich text through a reference reads as its plain text in every scalar slot") {
    // A rich-text cell is text to Excel: the formatting runs never change what a slot reads.
    val rich = (s: String) => CellValue.RichText(RichText.plain(s))
    val sheet = Sheet(SheetName.unsafe("Sheet1"))
      .put(ref"A1", rich("TRUE"))
      .put(ref"A2", rich("3"))
      .put(ref"A3", rich("abc"))
      .put(ref"A4", rich("FALSE"))
      .put(ref"A5", rich("45000"))
    val wb = Workbook(Vector(sheet))
    val cases = List(
      ("B1", "IF(A1,1,0)", N("1")),
      ("B2", "LEFT(\"abcdef\",A2)", T("abc")),
      ("B3", "NOT(A4)", B(true)),
      ("B4", "ABS(A2)", N("3")),
      ("B5", "YEAR(A5)", N("2023")),
      ("B6", "IF(A3,1,0)", Value),
      ("B7", "LEFT(\"abcdef\",A3)", Value),
      ("B8", "ABS(A3)", Value),
      ("B9", "YEAR(A3)", Value)
    )
    val diffs = cases.flatMap { case (at, formula, expected) =>
      val got = plainIn(wb, at, formula)
      Option.when(got != Right(expected))(s"$at =$formula: expected $expected, got $got")
    }
    assert(diffs.isEmpty, diffs.mkString("\n"))
  }

  test("#709: an out-of-range date serial is #VALUE! on every path until #726 makes it #NUM!") {
    // Excel gives #NUM! for a serial outside 0..2958465 (#726). Until then every path agrees on
    // #VALUE!, an error value the cell caches: numeric text or a number through a reference, a
    // literal number or literal text, a computed value and a LET binding. The decoder names the
    // reason with its own refusal, so the #NUM! fix changes one mapping.
    val sheet = Sheet(SheetName.unsafe("Sheet1"))
      .put(ref"A1", T("-5"))
      .put(ref"A2", T("3000000"))
      .put(ref"A3", N("-5"))
      .put(ref"A4", N("2958466"))
      .put(ref"A5", CellValue.Formula("A3*1", Some(N("-5"))))
    val wb = Workbook(Vector(sheet))
    val cases = List(
      ("B1", "YEAR(A1)", Value),
      ("B2", "YEAR(A2)", Value),
      ("B3", "YEAR(A3)", Value),
      ("B4", "MONTH(A4)", Value),
      ("B5", "DAY(A5)", Value),
      ("B6", "YEAR(-5)", Value),
      ("B7", "YEAR(\"-5\")", Value),
      ("B8", "YEAR(3000000)", Value),
      ("B9", "YEAR(A3*1)", Value),
      ("B10", "EDATE(-1,1)", Value),
      ("B11", "LET(x,-5,YEAR(x))", Value),
      ("B12", "IFERROR(YEAR(-5),\"caught\")", T("caught")),
      ("B13", "ERROR.TYPE(YEAR(A3))", N("3")),
      // the edges stay dates
      ("B14", "YEAR(1)", N("1900")),
      ("B15", "YEAR(2958465)", N("9999"))
    )
    val diffs = cases.flatMap { case (at, formula, expected) =>
      val got = plainIn(wb, at, formula)
      Option.when(got != Right(expected))(s"$at =$formula: expected $expected, got $got")
    }
    assert(diffs.isEmpty, diffs.mkString("\n"))
    // the decoder's refusal is the one sentinel instance the evaluator recognises (by identity)
    List(T("-5"), N("-5"), CellValue.Formula("x", Some(T("3000000")))).foreach { v =>
      TExpr.decodeAsDateScalar(Cell(ref"A1", v)) match
        case Left(err) => assert(err eq ScalarCoercion.DateSerialOutOfRange, s"$v: $err")
        case other => fail(s"$v: expected the out-of-range refusal, got $other")
    }
    assertEquals(TExpr.decodeAsDateScalar(Cell(ref"A1", T("45000"))).map(_.getYear), Right(2023))
  }

  test("#709: a caller's decoder cannot forge the out-of-range refusals") {
    // decodeOrCarried recognises the date and integer out-of-range refusals by instance, not by
    // their text: an equal ParseError from another decoder over a number cell stays CodecFailed.
    val sheet = Sheet(SheetName.unsafe("Sheet1")).put(ref"A1", N("5"))
    List(ScalarCoercion.DateSerialOutOfRange, ScalarCoercion.IntOutOfRange).foreach { sentinel =>
      val forged = sentinel match
        case CodecError.ParseError(v, t, d) => CodecError.ParseError(v, t, d)
        case CodecError.TypeMismatch(e, a) => CodecError.TypeMismatch(e, a)
      assertEquals(forged, sentinel)
      assert(!(forged eq sentinel))
      val decode: Cell => Either[CodecError, BigDecimal] = _ => Left(forged)
      Evaluator.instance.eval(TExpr.ref(ref"A1", decode), sheet) match
        case Left(EvalError.CodecFailed(at, err)) =>
          assertEquals(at, ref"A1")
          assertEquals(err, forged)
        case other => fail(s"expected CodecFailed, got $other")
    }
  }

  test("#709: a sheet-qualified reference names its sheet in the #VALUE! diagnostic") {
    val main = Sheet(SheetName.unsafe("Sheet1")).put(ref"C5", N("1")).put(ref"D5", T("x"))
    val other = Sheet(SheetName.unsafe("Other")).put(ref"C5", T("n/a"))
    val spaced = Sheet(SheetName.unsafe("My Data")).put(ref"C5", T("-5"))
    val wb = Workbook(Vector(main, other, spaced))
    def message[A](expr: TExpr[A]): String =
      Evaluator.instance.eval(expr, main, workbook = Some(wb)) match
        case Left(EvalError.ErrorValue(CellError.Value, Some(msg))) => msg
        case other => fail(s"expected #VALUE! with a message, got $other")
    val numeric = message(
      TExpr.SheetRef(SheetName.unsafe("Other"), ref"C5", Anchor.Relative, TExpr.decodeNumericScalar)
    )
    assert(numeric.startsWith("Other!C5: "), numeric)
    val date = message(
      TExpr.SheetRef(
        SheetName.unsafe("My Data"),
        ref"C5",
        Anchor.Relative,
        TExpr.decodeAsDateScalar
      )
    )
    assert(date.startsWith("'My Data'!C5: "), date)
    // the unqualified reference keeps the bare address
    val bare = message(TExpr.ref(ref"D5", TExpr.decodeNumericScalar))
    assert(bare.startsWith("D5: "), bare)
  }

  test("#709: integer slots refuse a magnitude outside the Int range without building it") {
    // Before the guard, truncating "1e2147483647" threw ArithmeticException (BigInteger would
    // overflow) and "1e10000000" spent ~10 s building 10^10000000. Now the size is read from
    // precision and scale first: out of the Int range is the int-slot #VALUE!, on every path
    // (LibreOffice 24.2 refuses these too, with Err:502); a magnitude below 1 truncates to 0.
    Future {
      val rich = (s: String) => CellValue.RichText(RichText.plain(s))
      val sheet = Sheet(SheetName.unsafe("Sheet1"))
        .put(ref"A1", T("1e2147483647"))
        .put(ref"A2", rich("1e2147483647"))
        .put(ref"A3", CellValue.Formula("x", Some(T("1e2147483647"))))
        .put(ref"A4", T("1e10000000"))
        .put(ref"A5", T("1e-10000000"))
        .put(ref"A6", N("1E10"))
        .put(ref"A7", T("1e10"))
        .put(ref"A8", N("2147483647"))
        .put(ref"A9", T("2147483648"))
        .put(ref"A10", T("0E+2147483647"))
        .put(ref"A11", T("-1e-2147483647"))
        .put(ref"A12", CellValue.Formula("x", Some(N("1E10"))))
      val wb = Workbook(Vector(sheet))
      val cases = List(
        ("B1", "LEFT(\"abc\",A1)", Value),
        ("B2", "LEFT(\"abc\",A2)", Value),
        ("B3", "LEFT(\"abc\",A3)", Value),
        ("B4", "LEFT(\"abc\",A4)", Value),
        ("B5", "LEFT(\"abc\",A5)", T("")),
        ("B6", "LEFT(\"abc\",A6)", Value),
        ("B7", "LEFT(\"abc\",A7)", Value),
        ("B8", "LEFT(\"abc\",A8)", T("abc")),
        ("B9", "LEFT(\"abc\",A9)", Value),
        ("B10", "LEFT(\"abc\",A10)", T("")),
        ("B11", "LEFT(\"abc\",A11)", T("")),
        ("B12", "LEFT(\"abc\",A12)", Value),
        ("B13", "MID(\"abcdef\",2,A1)", Value),
        ("B14", "LEFT(\"abc\",1E10)", Value),
        ("B15", "LEFT(\"abc\",\"1e2147483647\")", Value),
        ("B16", "LEFT(\"abc\",\"1e-2147483647\")", T("")),
        ("B17", "IFERROR(LEFT(\"abc\",A1),\"caught\")", T("caught"))
      )
      val diffs = cases.flatMap { case (at, formula, expected) =>
        val got = plainIn(wb, at, formula)
        Option.when(got != Right(expected))(s"$at =$formula: expected $expected, got $got")
      }
      assert(diffs.isEmpty, diffs.mkString("\n"))
    }(using ExecutionContext.global)
  }

  test("#709: a formula's cached non-numeric text is #VALUE! in a numeric slot") {
    val rich = (s: String) => CellValue.RichText(RichText.plain(s))
    val sheet = Sheet(SheetName.unsafe("Sheet1"))
      .put(ref"A1", CellValue.Formula("IF(TRUE,\"n/a\")", Some(T("n/a"))))
      .put(ref"A2", CellValue.Formula("\"5\"", Some(T("5"))))
      // a formula whose cached value is rich text reads through its plain text, like plain text
      .put(ref"A3", CellValue.Formula("IF(TRUE,\"n/a\")", Some(rich("n/a"))))
      .put(ref"A4", CellValue.Formula("\"5\"", Some(rich("5"))))
    val wb = Workbook(Vector(sheet))
    assertEquals(plainIn(wb, "B1", "ABS(A1)"), Right(Value))
    assertEquals(plainIn(wb, "B2", "ABS(A2)"), Right(N("5")))
    assertEquals(plainIn(wb, "B3", "IFERROR(ABS(A1),\"caught\")"), Right(T("caught")))
    assertEquals(plainIn(wb, "B4", "ABS(A3)"), Right(Value))
    assertEquals(plainIn(wb, "B5", "ABS(A4)"), Right(N("5")))
  }

  test("#709: a caller decoder's own refusal reason survives in the #VALUE! diagnostic") {
    val sheet = Sheet(SheetName.unsafe("Sheet1")).put(ref"A1", T("abc"))
    def message(decode: Cell => Either[CodecError, BigDecimal]): String =
      Evaluator.instance.eval(TExpr.ref(ref"A1", decode), sheet) match
        case Left(EvalError.ErrorValue(CellError.Value, Some(msg))) => msg
        case other => fail(s"expected #VALUE! with a message, got $other")
    // the built-in decoder's wording is unchanged: nothing appended
    assertEquals(message(TExpr.decodeNumericScalar), "A1: text 'abc' in a typed argument")
    // a caller's reason is kept
    val widget = message(_ => Left(CodecError.ParseError("abc", "Widget", "not a widget code")))
    assertEquals(
      widget,
      "A1: text 'abc' in a typed argument (cannot parse 'abc' as Widget: not a widget code)"
    )
    // and bounded: a 500-character reason is cut at 120 code points
    val long = message(_ => Left(CodecError.ParseError("abc", "Widget", "z" * 500)))
    assert(long.endsWith("…)"), long)
    assert(!long.contains("z" * 121), long)
  }

  test("#709: the diagnostic cut never splits a surrogate pair") {
    val emoji = "\uD83D\uDE00" // one code point, two UTF-16 chars
    // 39 code points then the emoji: the 40th code point is whole, nothing is cut
    val exactly40 = "x" * 39 + emoji
    assertEquals(ScalarCoercion.quotedText(exactly40), exactly40)
    // the emoji straddles char 40: it is kept whole before the ellipsis, never half a pair
    val straddling = "x" * 39 + emoji + "tail"
    assertEquals(ScalarCoercion.quotedText(straddling), "x" * 39 + emoji + "\u2026")
    // a long run of emoji is cut at 40 code points (80 chars)
    val emojis = emoji * 50
    assertEquals(ScalarCoercion.quotedText(emojis), emoji * 40 + "\u2026")
  }

  test("#709: numeric text in an XIRR/XNPV values range is still refused") {
    // The value fold is strict like the dates fold: '300 among the cash flows is #VALUE!. Excel's
    // XNPV documents #VALUE! for a nonnumeric argument, and XIRR shares the cash-flow model.
    // LibreOffice 24.2 differs: it drops the text pair (XIRR -0.3755, XNPV -235.34 here).
    val sheet = Sheet(SheetName.unsafe("Sheet1"))
      .put(ref"N1", N("-1000"))
      .put(ref"N2", T("300"))
      .put(ref"N3", N("800"))
      .put(ref"O1", N("44927"))
      .put(ref"O2", N("45000"))
      .put(ref"O3", N("45100"))
    val wb = Workbook(Vector(sheet))
    assertEquals(plainIn(wb, "P1", "XIRR(N1:N3,O1:O3)"), Right(Value))
    assertEquals(plainIn(wb, "P2", "XNPV(0.1,N1:N3,O1:O3)"), Right(Value))
  }
