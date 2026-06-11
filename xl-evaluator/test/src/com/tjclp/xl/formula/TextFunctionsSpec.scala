package com.tjclp.xl.formula

import com.tjclp.xl.*
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.sheets.Sheet
import munit.ScalaCheckSuite
import org.scalacheck.Prop.*
import org.scalacheck.Gen

import java.time.LocalDate

/**
 * Tests for the 6 text functions added in TJC-1055 / GH-116.
 *
 * Functions: TRIM, MID, FIND, SUBSTITUTE, VALUE, TEXT.
 *
 * Each remaining test kills a specific bug class — redundant boundary cases and overlapping
 * properties were dropped to bring this spec in line with the repo's per-category density (~10
 * tests / function). Pinning decisions:
 *   - Type coercion: text functions accept Number / Bool via Excel-style coercion (TRIM(123) ==
 *     "123", TRIM(true) == "TRUE").
 *   - Negative TEXT formatting preserves the default minus sign for single-section formats and also
 *     supports explicit two-section negative formats.
 */
class TextFunctionsSpec extends ScalaCheckSuite:
  private val emptySheet = new Sheet(name = SheetName.unsafe("Test"))
  private val evaluator = Evaluator.instance

  private def sheetWith(cells: (ARef, CellValue)*): Sheet =
    cells.foldLeft(emptySheet) { case (s, (r, v)) => s.put(r, v) }

  /** Number of non-overlapping occurrences of `needle` in `haystack`, scanning forward. */
  private def countOccurrences(haystack: String, needle: String): Int =
    if needle.isEmpty then 0
    else
      @annotation.tailrec
      def loop(idx: Int, acc: Int): Int =
        val next = haystack.indexOf(needle, idx)
        if next < 0 then acc else loop(next + needle.length, acc + 1)
      loop(0, 0)

  /** Small alphanumeric string generator — keeps shrinks readable. */
  private val smallString: Gen[String] =
    Gen.choose(0, 20).flatMap(n => Gen.listOfN(n, Gen.alphaNumChar).map(_.mkString))

  /** Non-empty single-or-multi char needle. */
  private val smallNeedle: Gen[String] =
    Gen.choose(1, 4).flatMap(n => Gen.listOfN(n, Gen.alphaNumChar).map(_.mkString))

  /**
   * Generate (haystack, needle) pairs where needle is guaranteed to be a substring of haystack.
   *
   * Avoids the high-discard-rate trap of `forAll(s, x){ s.contains(x) ==> ... }` where random
   * alphanumeric pairs almost never satisfy the precondition.
   */
  private val haystackWithNeedle: Gen[(String, String)] =
    for
      s <- smallString.suchThat(_.nonEmpty)
      start <- Gen.choose(0, s.length - 1)
      len <- Gen.choose(1, s.length - start)
    yield (s, s.substring(start, start + len))

  // ============================================================
  // §1. TRIM scalars
  // ============================================================

  test("TRIM: collapses internal runs and strips leading/trailing ASCII spaces") {
    val expr = TExpr.trim(TExpr.Lit("  hello   world  "))
    assertEquals(evaluator.eval(expr, emptySheet), Right("hello world"))
  }

  test("TRIM: whitespace-only ASCII spaces collapse to empty string") {
    val expr = TExpr.trim(TExpr.Lit("   "))
    assertEquals(evaluator.eval(expr, emptySheet), Right(""))
  }

  test("TRIM: scalar literals coerce to text") {
    assertEquals(emptySheet.evaluateFormula("=TRIM(123)"), Right(CellValue.Text("123")))
    assertEquals(emptySheet.evaluateFormula("=TRIM(TRUE)"), Right(CellValue.Text("TRUE")))
  }

  test("TRIM: collapses ASCII-space runs around a tab; tab is preserved") {
    val expr = TExpr.trim(TExpr.Lit("a   \t   b"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("a \t b"))
  }

  test("TRIM: internal nbsp run is not collapsed") {
    val expr = TExpr.trim(TExpr.Lit("a  b"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("a  b"))
  }

  test("TRIM: zero-width space (U+200B) is not stripped") {
    val expr = TExpr.trim(TExpr.Lit("​hello"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("​hello"))
  }

  test("TRIM: BOM (U+FEFF) is not stripped") {
    val expr = TExpr.trim(TExpr.Lit("﻿hello"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("﻿hello"))
  }

  // ============================================================
  // §2. MID scalars
  // ============================================================

  test("MID: extracts middle substring (issue golden)") {
    val expr = TExpr.mid(TExpr.Lit("Hello"), TExpr.Lit(2), TExpr.Lit(3))
    assertEquals(evaluator.eval(expr, emptySheet), Right("ell"))
  }

  test("MID: start one past end returns empty (boundary)") {
    val expr = TExpr.mid(TExpr.Lit("Hello"), TExpr.Lit(6), TExpr.Lit(1))
    assertEquals(evaluator.eval(expr, emptySheet), Right(""))
  }

  test("MID: start+len beyond length clamps to remainder") {
    val expr = TExpr.mid(TExpr.Lit("Hello"), TExpr.Lit(4), TExpr.Lit(100))
    assertEquals(evaluator.eval(expr, emptySheet), Right("lo"))
  }

  test("MID: start=0 returns EvalFailed naming MID and start") {
    val expr = TExpr.mid(TExpr.Lit("Hello"), TExpr.Lit(0), TExpr.Lit(3))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft, s"start=0 must fail; got $result")
    val msg = result.left.toOption.map(_.toString).getOrElse("")
    assert(msg.contains("MID") && msg.toLowerCase.contains("start"), s"Error msg: $msg")
  }

  test("MID: negative length returns EvalFailed naming MID and length") {
    val expr = TExpr.mid(TExpr.Lit("Hello"), TExpr.Lit(1), TExpr.Lit(-1))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft, s"len=-1 must fail; got $result")
    val msg = result.left.toOption.map(_.toString).getOrElse("")
    assert(msg.contains("MID") && msg.toLowerCase.contains("length"), s"Error msg: $msg")
  }

  test("MID: start=Int.MaxValue does not overflow (returns empty)") {
    val expr = TExpr.mid(TExpr.Lit("Hello"), TExpr.Lit(Int.MaxValue), TExpr.Lit(3))
    assertEquals(evaluator.eval(expr, emptySheet), Right(""))
  }

  // ============================================================
  // §3. FIND scalars
  // ============================================================

  test("FIND: locates first occurrence (1-indexed, issue golden)") {
    val expr = TExpr.find(TExpr.Lit("l"), TExpr.Lit("Hello"))
    assertEquals(evaluator.eval(expr, emptySheet), Right(BigDecimal(3)))
  }

  test("FIND: start parameter advances past first match") {
    val expr = TExpr.find(TExpr.Lit("l"), TExpr.Lit("Hello"), Some(TExpr.Lit(4)))
    assertEquals(evaluator.eval(expr, emptySheet), Right(BigDecimal(4)))
  }

  test("FIND: case-sensitive — lowercase 'h' not found in 'Hello'") {
    val expr = TExpr.find(TExpr.Lit("h"), TExpr.Lit("Hello"))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft, s"case-sensitive miss must fail; got $result")
  }

  test("FIND: regex metachars treated literally (no regex injection)") {
    val expr = TExpr.find(TExpr.Lit("."), TExpr.Lit("a.b"))
    assertEquals(evaluator.eval(expr, emptySheet), Right(BigDecimal(2)))
  }

  test("FIND: empty needle returns start position (Excel quirk)") {
    val r1 = evaluator.eval(TExpr.find(TExpr.Lit(""), TExpr.Lit("Hello")), emptySheet)
    assertEquals(r1, Right(BigDecimal(1)))
    val r2 = evaluator.eval(
      TExpr.find(TExpr.Lit(""), TExpr.Lit("Hello"), Some(TExpr.Lit(3))),
      emptySheet
    )
    assertEquals(r2, Right(BigDecimal(3)))
  }

  test("FIND: start=0 fails (Excel min start is 1)") {
    val expr = TExpr.find(TExpr.Lit("l"), TExpr.Lit("Hello"), Some(TExpr.Lit(0)))
    assert(evaluator.eval(expr, emptySheet).isLeft)
  }

  test("FIND: start past end of text fails — Excel: start_num > length → #VALUE!") {
    // Empty-needle case is the trap: without the strict bound, =FIND("", "abc", 4)
    // would silently succeed and return 4. Excel returns #VALUE! per docs.
    val emptyNeedle =
      TExpr.find(TExpr.Lit(""), TExpr.Lit("abc"), Some(TExpr.Lit(4)))
    assert(evaluator.eval(emptyNeedle, emptySheet).isLeft, "empty-needle past-end must fail")

    val nonEmptyNeedle =
      TExpr.find(TExpr.Lit("a"), TExpr.Lit("abc"), Some(TExpr.Lit(4)))
    assert(evaluator.eval(nonEmptyNeedle, emptySheet).isLeft, "non-empty-needle past-end must fail")
  }

  // ============================================================
  // §4. SUBSTITUTE scalars
  // ============================================================

  test("SUBSTITUTE: replaces all occurrences when instance omitted") {
    val expr = TExpr.substitute(TExpr.Lit("Hello"), TExpr.Lit("l"), TExpr.Lit("L"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("HeLLo"))
  }

  test("SUBSTITUTE: instance=2 replaces only the second occurrence (1-indexed)") {
    val expr =
      TExpr.substitute(TExpr.Lit("Hello"), TExpr.Lit("l"), TExpr.Lit("L"), Some(TExpr.Lit(2)))
    assertEquals(evaluator.eval(expr, emptySheet), Right("HelLo"))
  }

  test("SUBSTITUTE: regex metachars in old_text treated literally") {
    val expr = TExpr.substitute(TExpr.Lit("a.b.c"), TExpr.Lit("."), TExpr.Lit("X"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("aXbXc"))
  }

  test("SUBSTITUTE: replacement longer than match — no infinite loop") {
    val expr = TExpr.substitute(TExpr.Lit("Hello"), TExpr.Lit("l"), TExpr.Lit("ll"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("Hellllo"))
  }

  test("SUBSTITUTE: empty old_text returns text unchanged (Excel quirk)") {
    val expr = TExpr.substitute(TExpr.Lit("Hello"), TExpr.Lit(""), TExpr.Lit("X"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("Hello"))
  }

  test("SUBSTITUTE: instance=0 returns EvalFailed") {
    val expr =
      TExpr.substitute(TExpr.Lit("Hello"), TExpr.Lit("l"), TExpr.Lit("L"), Some(TExpr.Lit(0)))
    assert(evaluator.eval(expr, emptySheet).isLeft)
  }

  test("SUBSTITUTE: empty old_text with instance<1 still errors (instance validates first)") {
    val expr =
      TExpr.substitute(TExpr.Lit("Hello"), TExpr.Lit(""), TExpr.Lit("X"), Some(TExpr.Lit(0)))
    assert(evaluator.eval(expr, emptySheet).isLeft)
  }

  // ============================================================
  // §5. VALUE scalars
  // ============================================================

  test("VALUE: parses decimal with exact precision (kills Double impl)") {
    val expr = TExpr.value(TExpr.Lit("123.45"))
    assertEquals(evaluator.eval(expr, emptySheet), Right(BigDecimal("123.45")))
  }

  test("VALUE: strips currency symbol and thousands commas") {
    val expr = TExpr.value(TExpr.Lit("$1,234.56"))
    assertEquals(evaluator.eval(expr, emptySheet), Right(BigDecimal("1234.56")))
  }

  test("VALUE: percent string divided by 100") {
    val expr = TExpr.value(TExpr.Lit("45.5%"))
    assertEquals(evaluator.eval(expr, emptySheet), Right(BigDecimal("0.455")))
  }

  test("VALUE: accounting parentheses denote negative") {
    val expr = TExpr.value(TExpr.Lit("(1,234)"))
    assertEquals(evaluator.eval(expr, emptySheet), Right(BigDecimal(-1234)))
  }

  test("VALUE: sign + currency interaction ($-1,234.56)") {
    val expr = TExpr.value(TExpr.Lit("$-1,234.56"))
    assertEquals(evaluator.eval(expr, emptySheet), Right(BigDecimal("-1234.56")))
  }

  test("VALUE: empty string returns 0 (Excel quirk)") {
    val expr = TExpr.value(TExpr.Lit(""))
    assertEquals(evaluator.eval(expr, emptySheet), Right(BigDecimal(0)))
  }

  test("VALUE: accounting parens with inner negative sign is rejected (no double-negate)") {
    val expr = TExpr.value(TExpr.Lit("(-5)"))
    assert(evaluator.eval(expr, emptySheet).isLeft)
  }

  test("VALUE: accounting parens with inner positive sign is rejected") {
    val expr = TExpr.value(TExpr.Lit("(+5)"))
    assert(evaluator.eval(expr, emptySheet).isLeft)
  }

  test("VALUE: accounting parens with inner signed currency is rejected") {
    val beforeCurrency = TExpr.value(TExpr.Lit("(-$1,234.56)"))
    assert(evaluator.eval(beforeCurrency, emptySheet).isLeft)

    val afterCurrencyNegative = TExpr.value(TExpr.Lit("($-5)"))
    assert(evaluator.eval(afterCurrencyNegative, emptySheet).isLeft)

    val afterCurrencyPositive = TExpr.value(TExpr.Lit("($+5)"))
    assert(evaluator.eval(afterCurrencyPositive, emptySheet).isLeft)
  }

  // ============================================================
  // §6. TEXT scalars
  // ============================================================

  test("TEXT: basic decimal with rounding (issue golden)") {
    val expr = TExpr.text(TExpr.Lit(BigDecimal("1234.567")), TExpr.Lit("0.00"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("1234.57"))
  }

  test("TEXT: half-up rounding (1.555 → '1.56')") {
    val expr = TExpr.text(TExpr.Lit(BigDecimal("1.555")), TExpr.Lit("0.00"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("1.56"))
  }

  test("TEXT: percent format multiplies by 100") {
    val expr = TExpr.text(TExpr.Lit(BigDecimal("0.5")), TExpr.Lit("0%"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("50%"))
  }

  test("TEXT: negative number with single-section format preserves minus sign") {
    val expr = TExpr.text(TExpr.Lit(BigDecimal("-1234.5")), TExpr.Lit("0.00"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("-1234.50"))
  }

  test("TEXT: negative currency via explicit two-section format ('-$1,234.50')") {
    // Explicit negative sections use the absolute value and place the sign/currency
    // exactly as the format code specifies.
    val expr = TExpr.text(TExpr.Lit(BigDecimal("-1234.5")), TExpr.Lit("$#,##0.00;-$#,##0.00"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("-$1,234.50"))
  }

  test("TEXT: empty format string returns empty (Excel quirk)") {
    val expr = TExpr.text(TExpr.Lit(BigDecimal(1234)), TExpr.Lit(""))
    assertEquals(evaluator.eval(expr, emptySheet), Right(""))
  }

  test("TEXT: text input passes through unchanged") {
    val expr = TExpr.text(TExpr.Lit("hello"), TExpr.Lit("0.00"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("hello"))
  }

  test("TEXT: date input supports date tokens") {
    val expr = TExpr.text(TExpr.Lit(LocalDate.of(2025, 1, 15)), TExpr.Lit("yyyy-mm-dd"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("2025-01-15"))
  }

  // ============================================================
  // §7. Property-based laws (highest-leverage four)
  // ============================================================

  property("TRIM is idempotent: trim(trim(s)) == trim(s)") {
    forAll(smallString) { s =>
      val once = evaluator.eval(TExpr.trim(TExpr.Lit(s)), emptySheet)
      val twice = once.flatMap(t => evaluator.eval(TExpr.trim(TExpr.Lit(t)), emptySheet))
      once == twice
    }
  }

  property("FIND/MID/LEN coupling: MID(s, FIND(x,s), LEN(x)) == x when x ⊆ s") {
    forAll(haystackWithNeedle) { case (s, x) =>
      val findR = evaluator.eval(TExpr.find(TExpr.Lit(x), TExpr.Lit(s)), emptySheet)
      val lenR = evaluator.eval(TExpr.len(TExpr.Lit(x)), emptySheet)
      val combined = for
        k <- findR
        n <- lenR
        got <- evaluator.eval(
          TExpr.mid(TExpr.Lit(s), TExpr.Lit(k.toInt), TExpr.Lit(n.toInt)),
          emptySheet
        )
      yield got
      combined == Right(x)
    }
  }

  property(
    "SUBSTITUTE accounting law: len(sub(s,x,y)) - len(s) == count(x in s)*(len(y)-len(x))"
  ) {
    forAll(smallString, smallNeedle, smallString) { (s, x, y) =>
      x.nonEmpty ==> {
        val r = evaluator.eval(
          TExpr.substitute(TExpr.Lit(s), TExpr.Lit(x), TExpr.Lit(y)),
          emptySheet
        )
        val c = countOccurrences(s, x)
        r.fold(_ => false, t => t.length - s.length == c * (y.length - x.length))
      }
    }
  }

  property("VALUE/TEXT round-trip: value(text(n, '0.0000;-0.0000')) == n for 4-decimal n") {
    // Use unscaled-int construction so generator yields exact 4-decimal BigDecimals,
    // and forAllNoShrink so shrinking can't escape the generator's invariants.
    val gen = Gen.choose(-10000000L, 10000000L).map(u => BigDecimal(BigInt(u), 4))
    forAllNoShrink(gen) { n =>
      val r = for
        s <- evaluator.eval(TExpr.text(TExpr.Lit(n), TExpr.Lit("0.0000;-0.0000")), emptySheet)
        back <- evaluator.eval(TExpr.value(TExpr.Lit(s)), emptySheet)
      yield back == n
      r.getOrElse(false)
    }
  }

  // ============================================================
  // §8. Cell-value type matrix — TRIM exercises every CellValue case
  // ============================================================

  test("§8.1 TRIM(A1) where A1 is Empty cell returns ''") {
    val sheet = emptySheet
    assertEquals(sheet.evaluateFormula("=TRIM(A1)"), Right(CellValue.Text("")))
  }

  test("§8.2 TRIM(A1) where A1 is Number coerces to plain string '123'") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(123)))
    assertEquals(sheet.evaluateFormula("=TRIM(A1)"), Right(CellValue.Text("123")))
  }

  test("§8.3 TRIM(A1) where A1 is Bool coerces to 'TRUE'") {
    val sheet = sheetWith(ref"A1" -> CellValue.Bool(true))
    assertEquals(sheet.evaluateFormula("=TRIM(A1)"), Right(CellValue.Text("TRUE")))
  }

  test("§8.4 TRIM(A1) where A1 is Error propagates the error variant") {
    val sheet = sheetWith(ref"A1" -> CellValue.Error(CellError.Ref))
    val result = sheet.evaluateFormula("=TRIM(A1)")
    assert(result.isLeft, s"error must propagate; got $result")
  }

  // ============================================================
  // §9. Composability / nesting
  // ============================================================

  test("§9.1 LEN(TRIM('  hello  ')) == 5") {
    val sheet = emptySheet
    assertEquals(
      sheet.evaluateFormula("""=LEN(TRIM("  hello  "))"""),
      Right(CellValue.Number(BigDecimal(5)))
    )
  }

  test("§9.2 ISERROR(FIND('x','abc')) == TRUE — daily safe-search idiom") {
    val sheet = emptySheet
    assertEquals(
      sheet.evaluateFormula("""=ISERROR(FIND("x", "abc"))"""),
      Right(CellValue.Bool(true))
    )
  }

  test("§9.3 TEXT(VALUE('$1,234'), '#,##0') pipeline through formula engine") {
    val sheet = emptySheet
    assertEquals(
      sheet.evaluateFormula("""=TEXT(VALUE("$1,234"), "#,##0")"""),
      Right(CellValue.Text("1,234"))
    )
  }

  test("§9.4 numeric-returning calls work as Int args (returnsNumeric flag)") {
    // Pre-flag, only FIND + arithmetic were wrapped; SUM/ROUND/ABS in Int-arg
    // positions silently crashed at runtime via the asInstanceOf catch-all.
    // After flagging every BigDecimal-returning spec, any of them composes safely.
    val sheet = sheetWith(
      ref"A1" -> CellValue.Text("Hello, World"),
      ref"B1" -> CellValue.Number(BigDecimal(2)),
      ref"B2" -> CellValue.Number(BigDecimal(1)),
      ref"B3" -> CellValue.Number(BigDecimal(2))
    )
    // =MID(A1, SUM(B1:B3), 5) → start=5 (sum of 2+1+2), substring "o, Wo"
    assertEquals(
      sheet.evaluateFormula("=MID(A1, SUM(B1:B3), 5)"),
      Right(CellValue.Text("o, Wo"))
    )
    // =LEFT(A1, ROUND(B1 + 0.4, 0)) → ROUND(2.4, 0) = 2 → "He"
    assertEquals(
      sheet.evaluateFormula("=LEFT(A1, ROUND(B1 + 0.4, 0))"),
      Right(CellValue.Text("He"))
    )
    // =MID(A1, ABS(-3), 5) → start=3 → "llo, "
    assertEquals(
      sheet.evaluateFormula("=MID(A1, ABS(-3), 5)"),
      Right(CellValue.Text("llo, "))
    )
  }

  // ============================================================
  // §10. End-to-end via sheet.evaluateFormula (one per function — wiring smoke)
  // ============================================================

  test("§10.1 e2e: =TRIM(A1)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("  hello world  "))
    assertEquals(sheet.evaluateFormula("=TRIM(A1)"), Right(CellValue.Text("hello world")))
  }

  test("§10.2 e2e: =MID(A1, 2, 3)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("Hello"))
    assertEquals(sheet.evaluateFormula("=MID(A1, 2, 3)"), Right(CellValue.Text("ell")))
  }

  test("§10.3 e2e: =FIND(\"o\", A1)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("Hello"))
    assertEquals(
      sheet.evaluateFormula("""=FIND("o", A1)"""),
      Right(CellValue.Number(BigDecimal(5)))
    )
  }

  test("§10.4 e2e: =SUBSTITUTE(A1, \"l\", \"L\")") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("Hello"))
    assertEquals(
      sheet.evaluateFormula("""=SUBSTITUTE(A1, "l", "L")"""),
      Right(CellValue.Text("HeLLo"))
    )
  }

  test("§10.5 e2e: =VALUE(A1)") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("$1,234.56"))
    assertEquals(
      sheet.evaluateFormula("=VALUE(A1)"),
      Right(CellValue.Number(BigDecimal("1234.56")))
    )
  }

  test("§10.6 e2e: =TEXT(A1, \"0.00\")") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal("1234.567")))
    assertEquals(
      sheet.evaluateFormula("""=TEXT(A1, "0.00")"""),
      Right(CellValue.Text("1234.57"))
    )
  }

  // ============================================================
  // §11. Parser dispatch edges
  // ============================================================

  test("§11.1 Lowercase function name dispatches via case-insensitive registry") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("  hello  "))
    assertEquals(sheet.evaluateFormula("=trim(A1)"), Right(CellValue.Text("hello")))
  }

  test("§11.2 SUBSTITUTE with comma inside string literal preserves arg boundaries") {
    val sheet = emptySheet
    assertEquals(
      sheet.evaluateFormula("""=SUBSTITUTE("a,b", ",", ";")"""),
      Right(CellValue.Text("a;b"))
    )
  }

  // ============================================================
  // §12. Error-variant fidelity
  // ============================================================

  test("§12.1 MID start=0 message names function and constraint") {
    val expr = TExpr.mid(TExpr.Lit("Hi"), TExpr.Lit(0), TExpr.Lit(3))
    val result = evaluator.eval(expr, emptySheet)
    val msg = result.left.toOption.map(_.toString).getOrElse("")
    assert(msg.contains("MID") && msg.toLowerCase.contains("start"), msg)
  }

  test("§12.2 VALUE error echoes the offending input string") {
    val expr = TExpr.value(TExpr.Lit("not-a-number"))
    val result = evaluator.eval(expr, emptySheet)
    val msg = result.left.toOption.map(_.toString).getOrElse("")
    assert(msg.contains("VALUE") && msg.contains("not-a-number"), msg)
  }

  // ============================================================
  // §13. Sheet dependency graph
  // ============================================================

  test("§13.1 Cell with =TRIM(A1) recalculates when A1 changes") {
    val sheet1 = sheetWith(ref"A1" -> CellValue.Text("  one  "))
    val r1 = sheet1.evaluateFormula("=TRIM(A1)")
    assertEquals(r1, Right(CellValue.Text("one")))

    val sheet2 = sheet1.put(ref"A1", CellValue.Text("  two  "))
    val r2 = sheet2.evaluateFormula("=TRIM(A1)")
    assertEquals(r2, Right(CellValue.Text("two")))

    assertNotEquals(r1, r2)
  }

  // ============================================================
  // §14. Determinism / printer round-trip
  // ============================================================

  test("§14.1 FormulaPrinter round-trip for each of the 6 functions") {
    val cases: List[String] = List(
      """=TRIM(A1)""",
      """=MID(A1, 2, 3)""",
      """=FIND("o", A1)""",
      """=FIND("o", A1, 3)""",
      """=SUBSTITUTE(A1, "l", "L")""",
      """=SUBSTITUTE(A1, "l", "L", 2)""",
      """=VALUE(A1)""",
      """=TEXT(A1, "0.00")"""
    )
    cases.foreach { src =>
      val parsed = com.tjclp.xl.formula.parser.FormulaParser.parse(src)
      assert(parsed.isRight, s"parse failed: $src — $parsed")
      val printed = parsed.toOption
        .map(com.tjclp.xl.formula.printer.FormulaPrinter.print(_, includeEquals = true))
        .getOrElse("")
      val reparsed = com.tjclp.xl.formula.parser.FormulaParser.parse(printed)
      assertEquals(reparsed, parsed, s"round-trip drift: $src -> $printed")
    }
  }
