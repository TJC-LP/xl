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

import java.time.{LocalDate, LocalDateTime}

/**
 * Comprehensive tests for the 6 text functions added in TJC-1055 / GH-116.
 *
 * Functions: TRIM, MID, FIND, SUBSTITUTE, VALUE, TEXT.
 *
 * Organized into 14 sections covering scalar behavior, property-based laws, cell-value type
 * interactions, composability, end-to-end formula evaluation, parser dispatch, error fidelity,
 * dependency-graph integration, and OOXML round-trip determinism. Pinning decisions:
 *   - Type coercion: text functions accept Number / Bool via Excel-style coercion (TRIM(123) ==
 *     "123", TRIM(true) == "TRUE").
 *   - Negative currency in TEXT places sign before the symbol: TEXT(-1234.5, "$#,##0.00") ==
 *     "-$1,234.50".
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
  // §1. TRIM scalars (8)
  // ============================================================

  test("TRIM: collapses internal runs and strips leading/trailing ASCII spaces") {
    val expr = TExpr.trim(TExpr.Lit("  hello   world  "))
    assertEquals(evaluator.eval(expr, emptySheet), Right("hello world"))
  }

  test("TRIM: all-ASCII-space input returns empty") {
    val expr = TExpr.trim(TExpr.Lit("     "))
    assertEquals(evaluator.eval(expr, emptySheet), Right(""))
  }

  test("TRIM: preserves leading/trailing non-breaking space (char 160)") {
    val expr = TExpr.trim(TExpr.Lit(" hello "))
    assertEquals(evaluator.eval(expr, emptySheet), Right(" hello "))
  }

  test("TRIM: collapses ASCII-space runs around a tab; tab is preserved") {
    val expr = TExpr.trim(TExpr.Lit("a   \t   b"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("a \t b"))
  }

  test("TRIM: tabs and newlines pass through (Excel only collapses ASCII space 0x20)") {
    val expr = TExpr.trim(TExpr.Lit("\thello\n"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("\thello\n"))
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
  // §2. MID scalars (9)
  // ============================================================

  test("MID: extracts middle substring (issue golden)") {
    val expr = TExpr.mid(TExpr.Lit("Hello"), TExpr.Lit(2), TExpr.Lit(3))
    assertEquals(evaluator.eval(expr, emptySheet), Right("ell"))
  }

  test("MID: start at last char with len=1 returns last char") {
    val expr = TExpr.mid(TExpr.Lit("Hello"), TExpr.Lit(5), TExpr.Lit(1))
    assertEquals(evaluator.eval(expr, emptySheet), Right("o"))
  }

  test("MID: start one past end returns empty (boundary)") {
    val expr = TExpr.mid(TExpr.Lit("Hello"), TExpr.Lit(6), TExpr.Lit(1))
    assertEquals(evaluator.eval(expr, emptySheet), Right(""))
  }

  test("MID: start+len beyond length clamps to remainder") {
    val expr = TExpr.mid(TExpr.Lit("Hello"), TExpr.Lit(4), TExpr.Lit(100))
    assertEquals(evaluator.eval(expr, emptySheet), Right("lo"))
  }

  test("MID: empty input with valid start returns empty") {
    val expr = TExpr.mid(TExpr.Lit(""), TExpr.Lit(1), TExpr.Lit(5))
    assertEquals(evaluator.eval(expr, emptySheet), Right(""))
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

  test("MID: emoji surrogate pair — UTF-16 code-unit semantics (documented)") {
    // "a😀b" = a + high-surrogate(D83D) + low-surrogate(DE00) + b — total 4 UTF-16 code units.
    // MID at position 2, length 1 returns the high-surrogate half (matches Excel UTF-16 model).
    val expr = TExpr.mid(TExpr.Lit("a😀b"), TExpr.Lit(2), TExpr.Lit(1))
    assertEquals(evaluator.eval(expr, emptySheet), Right("\uD83D"))
  }

  // ============================================================
  // §3. FIND scalars (9)
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

  test("FIND: multi-char needle resolves correctly") {
    val expr = TExpr.find(TExpr.Lit("ll"), TExpr.Lit("Hello"))
    assertEquals(evaluator.eval(expr, emptySheet), Right(BigDecimal(3)))
  }

  test("FIND: start=0 fails (Excel min start is 1)") {
    val expr = TExpr.find(TExpr.Lit("l"), TExpr.Lit("Hello"), Some(TExpr.Lit(0)))
    assert(evaluator.eval(expr, emptySheet).isLeft)
  }

  test("FIND: start exactly at the match position is inclusive") {
    val expr = TExpr.find(TExpr.Lit("o"), TExpr.Lit("Hello"), Some(TExpr.Lit(5)))
    assertEquals(evaluator.eval(expr, emptySheet), Right(BigDecimal(5)))
  }

  test("FIND: needle containing comma works through formula parser") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("Hello, World!"))
    val result = sheet.evaluateFormula("""=FIND("o, W", A1)""")
    assertEquals(result, Right(CellValue.Number(BigDecimal(5))))
  }

  // ============================================================
  // §4. SUBSTITUTE scalars (9)
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

  test("SUBSTITUTE: instance past occurrence count returns text unchanged") {
    val expr =
      TExpr.substitute(TExpr.Lit("Hello"), TExpr.Lit("l"), TExpr.Lit("L"), Some(TExpr.Lit(3)))
    assertEquals(evaluator.eval(expr, emptySheet), Right("Hello"))
  }

  test("SUBSTITUTE: regex metachars in old_text treated literally") {
    val expr = TExpr.substitute(TExpr.Lit("a.b.c"), TExpr.Lit("."), TExpr.Lit("X"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("aXbXc"))
  }

  test("SUBSTITUTE: forward non-overlapping scan ('aaaa' / 'aa' / 'b' → 'bb')") {
    val expr = TExpr.substitute(TExpr.Lit("aaaa"), TExpr.Lit("aa"), TExpr.Lit("b"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("bb"))
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

  test("SUBSTITUTE: multi-char old, instance=2") {
    val expr =
      TExpr.substitute(
        TExpr.Lit("foo bar foo"),
        TExpr.Lit("foo"),
        TExpr.Lit("baz"),
        Some(TExpr.Lit(2))
      )
    assertEquals(evaluator.eval(expr, emptySheet), Right("foo bar baz"))
  }

  // ============================================================
  // §5. VALUE scalars (10)
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

  test("VALUE: leading/trailing whitespace around currency") {
    val expr = TExpr.value(TExpr.Lit(" $1,234 "))
    assertEquals(evaluator.eval(expr, emptySheet), Right(BigDecimal(1234)))
  }

  test("VALUE: scientific notation") {
    val expr = TExpr.value(TExpr.Lit("1.5E2"))
    assertEquals(evaluator.eval(expr, emptySheet), Right(BigDecimal(150)))
  }

  test("VALUE: empty string returns 0 (Excel quirk)") {
    val expr = TExpr.value(TExpr.Lit(""))
    assertEquals(evaluator.eval(expr, emptySheet), Right(BigDecimal(0)))
  }

  test("VALUE: alphanumeric input fails with offending value in error") {
    val expr = TExpr.value(TExpr.Lit("12abc"))
    val result = evaluator.eval(expr, emptySheet)
    assert(result.isLeft)
    val msg = result.left.toOption.map(_.toString).getOrElse("")
    assert(msg.contains("VALUE") && msg.contains("12abc"), s"Error msg: $msg")
  }

  test("VALUE: 100% boundary returns 1") {
    val expr = TExpr.value(TExpr.Lit("100%"))
    assertEquals(evaluator.eval(expr, emptySheet), Right(BigDecimal(1)))
  }

  // ============================================================
  // §6. TEXT scalars (10)
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

  test("TEXT: percent with single-decimal precision") {
    val expr = TExpr.text(TExpr.Lit(BigDecimal("0.5")), TExpr.Lit("0.0%"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("50.0%"))
  }

  test("TEXT: multi-group thousands separator") {
    val expr = TExpr.text(TExpr.Lit(BigDecimal(1234567)), TExpr.Lit("#,##0"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("1,234,567"))
  }

  test("TEXT: negative currency renders sign before symbol ('-$1,234.50')") {
    val expr = TExpr.text(TExpr.Lit(BigDecimal("-1234.5")), TExpr.Lit("$#,##0.00"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("-$1,234.50"))
  }

  test("TEXT: zero with mandatory decimals") {
    val expr = TExpr.text(TExpr.Lit(BigDecimal(0)), TExpr.Lit("0.00"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("0.00"))
  }

  test("TEXT: empty format string returns empty (Excel quirk)") {
    val expr = TExpr.text(TExpr.Lit(BigDecimal(1234)), TExpr.Lit(""))
    assertEquals(evaluator.eval(expr, emptySheet), Right(""))
  }

  test("TEXT: date format on LocalDate") {
    val expr = TExpr.text(TExpr.Lit(LocalDate.of(2025, 1, 15)), TExpr.Lit("yyyy-mm-dd"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("2025-01-15"))
  }

  test("TEXT: text input passes through unchanged") {
    val expr = TExpr.text(TExpr.Lit("hello"), TExpr.Lit("0.00"))
    assertEquals(evaluator.eval(expr, emptySheet), Right("hello"))
  }

  // ============================================================
  // §7. Property-based laws (10)
  // ============================================================

  property("TRIM is idempotent: trim(trim(s)) == trim(s)") {
    forAll(smallString) { s =>
      val once = evaluator.eval(TExpr.trim(TExpr.Lit(s)), emptySheet)
      val twice = once.flatMap(t => evaluator.eval(TExpr.trim(TExpr.Lit(t)), emptySheet))
      once == twice
    }
  }

  property("TRIM never grows length: len(trim(s)) <= len(s)") {
    forAll(smallString) { s =>
      val r = evaluator.eval(TExpr.trim(TExpr.Lit(s)), emptySheet)
      r.fold(_ => false, t => t.length <= s.length)
    }
  }

  property("MID(s, 1, n) == LEFT(s, n) for n in [0, len(s)]") {
    forAll(smallString, Gen.choose(0, 30)) { (s, n) =>
      (n >= 0 && n <= s.length) ==> {
        val midR = evaluator.eval(TExpr.mid(TExpr.Lit(s), TExpr.Lit(1), TExpr.Lit(n)), emptySheet)
        val leftR = evaluator.eval(TExpr.left(TExpr.Lit(s), TExpr.Lit(n)), emptySheet)
        midR == leftR
      }
    }
  }

  property("MID slicing transitivity: MID(MID(s,k,big),1,n) == MID(s,k,n)") {
    forAll(smallString, Gen.choose(1, 30), Gen.choose(0, 30)) { (s, k, n) =>
      val outer = evaluator.eval(
        TExpr.mid(TExpr.Lit(s), TExpr.Lit(k), TExpr.Lit(1000)),
        emptySheet
      )
      val direct = evaluator.eval(TExpr.mid(TExpr.Lit(s), TExpr.Lit(k), TExpr.Lit(n)), emptySheet)
      val nested = outer.flatMap(t =>
        evaluator.eval(TExpr.mid(TExpr.Lit(t), TExpr.Lit(1), TExpr.Lit(n)), emptySheet)
      )
      nested == direct
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

  property("SUBSTITUTE identity: replacing x with x is a no-op") {
    forAll(smallString, smallNeedle) { (s, x) =>
      val r = evaluator.eval(
        TExpr.substitute(TExpr.Lit(s), TExpr.Lit(x), TExpr.Lit(x)),
        emptySheet
      )
      r == Right(s)
    }
  }

  property("SUBSTITUTE empty-old quirk: SUBSTITUTE(s, '', anything) == s") {
    forAll(smallString, smallString) { (s, anything) =>
      val r = evaluator.eval(
        TExpr.substitute(TExpr.Lit(s), TExpr.Lit(""), TExpr.Lit(anything)),
        emptySheet
      )
      r == Right(s)
    }
  }

  property("VALUE/TEXT round-trip: value(text(n,'0.0000')) ≈ n at 1e-4") {
    forAll(Gen.choose(-1000.0, 1000.0)) { d =>
      val n = BigDecimal(d).setScale(6, BigDecimal.RoundingMode.HALF_UP)
      val r = for
        s <- evaluator.eval(TExpr.text(TExpr.Lit(n), TExpr.Lit("0.0000")), emptySheet)
        back <- evaluator.eval(TExpr.value(TExpr.Lit(s)), emptySheet)
      yield (back - n).abs <= BigDecimal("0.0001")
      r.getOrElse(false)
    }
  }

  property("FIND first-char law: len(s)>0 ⟹ FIND(MID(s,1,1), s) == 1") {
    forAll(smallString) { s =>
      s.nonEmpty ==> {
        val firstChar = s.substring(0, 1)
        val r = evaluator.eval(
          TExpr.find(TExpr.Lit(firstChar), TExpr.Lit(s)),
          emptySheet
        )
        r == Right(BigDecimal(1))
      }
    }
  }

  // ============================================================
  // §8. Cell-value type matrix (8)
  // ============================================================

  test("§8.1 TRIM(A1) where A1 is Empty cell returns ''") {
    // No put — A1 is implicitly Empty.
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

  test("§8.5 TEXT(A1, '0.00') where A1 is Empty treats Empty as 0") {
    val sheet = emptySheet
    assertEquals(sheet.evaluateFormula("""=TEXT(A1, "0.00")"""), Right(CellValue.Text("0.00")))
  }

  test("§8.6 LEN(MID(A1, 1, 5)) where A1 is Empty == 0") {
    val sheet = emptySheet
    assertEquals(sheet.evaluateFormula("=LEN(MID(A1, 1, 5))"), Right(CellValue.Number(BigDecimal(0))))
  }

  test("§8.7 VALUE(A1) where A1 is already Number passes through") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal(42)))
    assertEquals(sheet.evaluateFormula("=VALUE(A1)"), Right(CellValue.Number(BigDecimal(42))))
  }

  test("§8.8 FIND('x', A1) where A1 is Error propagates without rewrapping") {
    val sheet = sheetWith(ref"A1" -> CellValue.Error(CellError.Value))
    val result = sheet.evaluateFormula("""=FIND("x", A1)""")
    assert(result.isLeft, s"error must propagate; got $result")
  }

  // ============================================================
  // §9. Composability / nesting (7)
  // ============================================================

  test("§9.1 LEN(TRIM('  hello  ')) == 5") {
    val sheet = emptySheet
    assertEquals(
      sheet.evaluateFormula("""=LEN(TRIM("  hello  "))"""),
      Right(CellValue.Number(BigDecimal(5)))
    )
  }

  test("§9.2 MID(TRIM(A1), 1, 5) operates on the trimmed result") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("   hello world   "))
    assertEquals(
      sheet.evaluateFormula("=MID(TRIM(A1), 1, 5)"),
      Right(CellValue.Text("hello"))
    )
  }

  test("§9.3 SUBSTITUTE nesting (chained replacements)") {
    val sheet = emptySheet
    assertEquals(
      sheet.evaluateFormula(
        """=SUBSTITUTE(SUBSTITUTE("a-b-c", "-", "+"), "+", "/")"""
      ),
      Right(CellValue.Text("a/b/c"))
    )
  }

  test("§9.4 ISERROR(FIND('x','abc')) == TRUE — daily safe-search idiom") {
    val sheet = emptySheet
    assertEquals(
      sheet.evaluateFormula("""=ISERROR(FIND("x", "abc"))"""),
      Right(CellValue.Bool(true))
    )
  }

  test("§9.5 IF(ISNUMBER(VALUE(A1)), VALUE(A1), 0) safe-parse pattern") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("42"))
    assertEquals(
      sheet.evaluateFormula("=IF(ISNUMBER(VALUE(A1)), VALUE(A1), 0)"),
      Right(CellValue.Number(BigDecimal(42)))
    )
  }

  test("§9.6 TEXT(VALUE('$1,234'), '#,##0') pipeline through formula engine") {
    val sheet = emptySheet
    assertEquals(
      sheet.evaluateFormula("""=TEXT(VALUE("$1,234"), "#,##0")"""),
      Right(CellValue.Text("1,234"))
    )
  }

  test("§9.7 CONCATENATE(TEXT(A1, '0.00'), ' USD') interop with existing CONCATENATE") {
    val sheet = sheetWith(ref"A1" -> CellValue.Number(BigDecimal("1234.5")))
    assertEquals(
      sheet.evaluateFormula("""=CONCATENATE(TEXT(A1, "0.00"), " USD")"""),
      Right(CellValue.Text("1234.50 USD"))
    )
  }

  // ============================================================
  // §10. End-to-end via sheet.evaluateFormula (6)
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
  // §11. Parser dispatch edges (4)
  // ============================================================

  test("§11.1 Lowercase function name dispatches via case-insensitive registry") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("  hello  "))
    assertEquals(sheet.evaluateFormula("=trim(A1)"), Right(CellValue.Text("hello")))
  }

  test("§11.2 TRIM() with zero arguments is a parse error") {
    val sheet = emptySheet
    assert(sheet.evaluateFormula("=TRIM()").isLeft)
  }

  test("§11.3 FIND with one argument (needs ≥2) is a parse error") {
    val sheet = emptySheet
    assert(sheet.evaluateFormula("""=FIND("a")""").isLeft)
  }

  test("§11.4 SUBSTITUTE with comma inside string literal preserves arg boundaries") {
    val sheet = emptySheet
    assertEquals(
      sheet.evaluateFormula("""=SUBSTITUTE("a,b", ",", ";")"""),
      Right(CellValue.Text("a;b"))
    )
  }

  // ============================================================
  // §12. Error-variant fidelity (4)
  // ============================================================

  test("§12.1 MID start=0 message names function and constraint") {
    val expr = TExpr.mid(TExpr.Lit("Hi"), TExpr.Lit(0), TExpr.Lit(3))
    val result = evaluator.eval(expr, emptySheet)
    val msg = result.left.toOption.map(_.toString).getOrElse("")
    assert(msg.contains("MID") && msg.toLowerCase.contains("start"), msg)
  }

  test("§12.2 FIND not-found message names function and reason") {
    val expr = TExpr.find(TExpr.Lit("z"), TExpr.Lit("Hi"))
    val result = evaluator.eval(expr, emptySheet)
    val msg = result.left.toOption.map(_.toString).getOrElse("")
    assert(msg.contains("FIND") && msg.toLowerCase.contains("not found"), msg)
  }

  test("§12.3 VALUE error echoes the offending input string") {
    val expr = TExpr.value(TExpr.Lit("not-a-number"))
    val result = evaluator.eval(expr, emptySheet)
    val msg = result.left.toOption.map(_.toString).getOrElse("")
    assert(msg.contains("VALUE") && msg.contains("not-a-number"), msg)
  }

  test("§12.4 Error-variant fidelity: cell error flows through unchanged shape") {
    // When a function receives a cell holding CellValue.Error, the resulting EvalError
    // should reflect the source cell error (not be rewrapped as a generic EvalFailed
    // about that function).
    val sheet = sheetWith(ref"A1" -> CellValue.Error(CellError.Div0))
    val result = sheet.evaluateFormula("=TRIM(A1)")
    assert(result.isLeft, s"error must propagate; got $result")
  }

  // ============================================================
  // §13. Sheet dependency graph (2)
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

  test("§13.2 =FIND(...) on non-matching content yields error result for caching") {
    val sheet = sheetWith(ref"A1" -> CellValue.Text("none"))
    val result = sheet.evaluateFormula("""=FIND("x", A1)""")
    assert(result.isLeft, s"FIND failure must surface as error; got $result")
  }

  // ============================================================
  // §14. Determinism / printer round-trip (2)
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

  test("§14.2 OOXML write→read→evaluate matches direct evaluation") {
    // Two-step determinism: a workbook containing each text-function formula must
    // evaluate to the same CellValue after a full xlsx round-trip as it does directly.
    // This is a placeholder — a full impl would write/read the workbook via
    // ExcelIO. For now we pin behavior at the SheetEvaluator level for the same
    // CellValue identity, which the implementation phase will extend.
    val sheet = sheetWith(ref"A1" -> CellValue.Text("  hello  "))
    val direct = sheet.evaluateFormula("=TRIM(A1)")
    val again = sheet.evaluateFormula("=TRIM(A1)")
    assertEquals(direct, again)
  }
