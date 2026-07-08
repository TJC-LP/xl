package com.tjclp.xl.formula

import munit.FunSuite
import com.tjclp.xl.*
import com.tjclp.xl.unsafe.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet

/**
 * GH-335: Excel semantics for the ordered comparison operators (< <= > >=).
 *
 * Excel compares polymorphically, in both scalar and array/broadcast positions:
 *   - text vs text: case-insensitive, lexicographic
 *   - number vs number: numeric order (dates ARE numbers via their Excel serial)
 *   - boolean vs boolean: FALSE < TRUE
 *   - cross-type: number < text < logical (no numeric parsing of text under comparison)
 *   - empty cells coerce relative to the other operand: 0 vs numbers, "" vs text, FALSE vs
 *     booleans
 *
 * The dogfooding repro: SUMPRODUCT over a text column raised
 * TypeMismatch(numeric argument,number,z) instead of comparing lexicographically.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class ComparisonSemanticsSpec extends FunSuite:

  private def sheetWith(entries: (ARef, Any)*): Sheet =
    entries.foldLeft(Sheet("Test")) { case (s, (ref, value)) =>
      val cv = value match
        case cv: CellValue => cv
        case str: String => CellValue.Text(str)
        case n: Int => CellValue.Number(BigDecimal(n))
        case n: BigDecimal => CellValue.Number(n)
        case b: Boolean => CellValue.Bool(b)
        case d: java.time.LocalDate => CellValue.DateTime(d.atStartOfDay())
        case _ => CellValue.Text(value.toString)
      s.put(ref, cv).unsafe
    }

  private def evalAs[A](formula: String, sheet: Sheet): Either[String, A] =
    FormulaParser.parse(formula) match
      case Right(expr) =>
        Evaluator.eval(expr, sheet) match
          case Right(value) => Right(value.asInstanceOf[A])
          case Left(err) => Left(s"Eval error: $err")
      case Left(err) => Left(s"Parse error: $err")

  private def assertBool(formula: String, expected: Boolean, sheet: Sheet = Sheet("Test"))(implicit
    loc: munit.Location
  ): Unit =
    evalAs[Any](formula, sheet) match
      case Right(b: Boolean) => assertEquals(b, expected, s"$formula")
      case Right(other) => fail(s"$formula: expected Boolean, got $other")
      case Left(err) => fail(s"$formula failed: $err")

  private def assertNum(formula: String, expected: Int, sheet: Sheet)(implicit
    loc: munit.Location
  ): Unit =
    evalAs[Any](formula, sheet) match
      case Right(n: BigDecimal) => assertEquals(n, BigDecimal(expected), s"$formula")
      case Right(other) => fail(s"$formula: expected BigDecimal, got $other")
      case Left(err) => fail(s"$formula failed: $err")

  // ===== Issue repro (a): text range < text literal inside SUMPRODUCT =====

  test("GH-335: SUMPRODUCT((B1:B2<\"z\")*1) counts text below z") {
    val sheet = sheetWith(ref"B1" -> "apple", ref"B2" -> "banana")
    assertNum("=SUMPRODUCT((B1:B2<\"z\")*1)", 2, sheet)
  }

  // ===== Issue repro (b): dates-as-text column compared to a text anchor =====

  test("GH-335: SUMPRODUCT over dates-as-text compares lexicographically") {
    val sheet = sheetWith(
      ref"F2" -> "15/07/20",
      ref"F3" -> "05/08/20",
      ref"F4" -> "01/01/21",
      ref"F5" -> "20/06/20"
    )
    // Lexicographic: only "01/01/21" < "05/08/20"
    assertNum("=SUMPRODUCT(($F$2:$F$5<$F$3)*1)", 1, sheet)
  }

  // ===== Scalar text comparisons =====

  test("text vs text compares case-insensitively, lexicographically") {
    assertBool("=\"apple\"<\"BANANA\"", true)
    assertBool("=\"B\">\"a\"", true)
    assertBool("=\"apple\"<=\"APPLE\"", true)
    assertBool("=\"apple\">=\"APPLE\"", true)
    assertBool("=\"zebra\"<\"apple\"", false)
  }

  test("text cell vs text cell comparison") {
    val sheet = sheetWith(ref"A1" -> "apple", ref"A2" -> "Banana")
    assertBool("=A1<A2", true, sheet)
    assertBool("=A2<A1", false, sheet)
  }

  // ===== Cross-type ordering: number < text < logical =====

  test("number < text (even numeric-looking text)") {
    assertBool("=5<\"5\"", true)
    assertBool("=\"2\">100", true)
    assertBool("=100<\"2\"", true)
  }

  test("text < logical") {
    assertBool("=\"zzz\"<TRUE", true)
    assertBool("=\"zzz\"<FALSE", true)
    assertBool("=FALSE>\"zzz\"", true)
  }

  test("number < logical, FALSE < TRUE") {
    assertBool("=TRUE>100", true)
    assertBool("=FALSE>100", true)
    assertBool("=FALSE<TRUE", true)
    assertBool("=TRUE<=TRUE", true)
    assertBool("=TRUE<FALSE", false)
  }

  // ===== Empty-cell coercion =====

  test("empty cell compares as 0 against numbers") {
    val sheet = Sheet("Test")
    assertBool("=A1<5", true, sheet)
    assertBool("=A1>-1", true, sheet)
    assertBool("=A1>0", false, sheet)
  }

  test("empty cell compares as empty string against text") {
    val sheet = Sheet("Test")
    assertBool("=A1<\"a\"", true, sheet)
    assertBool("=A1>\"a\"", false, sheet)
  }

  test("empty cell compares as FALSE against booleans") {
    val sheet = Sheet("Test")
    assertBool("=A1<TRUE", true, sheet)
    assertBool("=A1<FALSE", false, sheet)
    assertBool("=A1>=FALSE", true, sheet)
  }

  test("empty cell equals 0 and empty string (scalar and array)") {
    val sheet = sheetWith(ref"A2" -> 5)
    assertBool("=A1=0", true, sheet)
    assertBool("=A1=\"\"", true, sheet)
    assertNum("=SUMPRODUCT((A1:A2=0)*1)", 1, sheet)
  }

  // ===== Numbers and dates =====

  test("numeric comparisons unchanged") {
    assertBool("=1<2", true)
    assertBool("=2<=2", true)
    assertBool("=3>4", false)
  }

  test("date cells compare as serial numbers") {
    val sheet = sheetWith(
      ref"A1" -> java.time.LocalDate.of(2020, 1, 1),
      ref"A2" -> java.time.LocalDate.of(2021, 1, 1)
    )
    assertBool("=A1<A2", true, sheet)
    assertBool("=A1<DATE(2020,6,1)", true, sheet)
    assertBool("=A2<DATE(2020,6,1)", false, sheet)
    assertNum("=SUMPRODUCT((A1:A2<DATE(2020,6,1))*1)", 1, sheet)
  }

  // ===== Array/broadcast paths =====

  test("mixed-type array vs number scalar uses type ranks elementwise") {
    val sheet = sheetWith(ref"A1" -> 1, ref"A2" -> "apple", ref"A3" -> true)
    // 1>100 FALSE; "apple">100 TRUE (text>number); TRUE>100 TRUE (logical>number)
    assertNum("=SUMPRODUCT((A1:A3>100)*1)", 2, sheet)
  }

  test("text array vs text cell anchor broadcasts") {
    val sheet = sheetWith(ref"A1" -> "apple", ref"A2" -> "banana", ref"B1" -> "b")
    // case-insensitive lexicographic: "apple"<"b" TRUE, "banana">"b" (prefix) -> not < "b"
    assertNum("=SUMPRODUCT((A1:A2<B1)*1)", 1, sheet)
  }

  test("empty cells in range compare as 0 against numeric scalar") {
    val sheet = sheetWith(ref"A2" -> 5)
    // A1 empty -> 0<3 TRUE; 5<3 FALSE
    assertNum("=SUMPRODUCT((A1:A2<3)*1)", 1, sheet)
  }

  test("array comparison feeds boolean arithmetic (issue shape)") {
    val sheet = sheetWith(ref"A1" -> "x", ref"A2" -> "y", ref"B1" -> 10, ref"B2" -> 20)
    // (text<"y") -> [TRUE, FALSE]; * B gives [10, 0]
    assertNum("=SUMPRODUCT((A1:A2<\"y\")*B1:B2)", 10, sheet)
  }
