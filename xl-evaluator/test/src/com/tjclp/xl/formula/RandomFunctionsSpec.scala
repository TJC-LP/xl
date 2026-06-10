package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.formula.functions.FunctionRegistry
import com.tjclp.xl.formula.parser.FormulaParser
import munit.ScalaCheckSuite
import org.scalacheck.Prop.*
import org.scalacheck.Gen

/**
 * GH-115: RAND() and RANDBETWEEN(bottom, top).
 *
 * Purity-preserving design: randomness is an explicit Rng capability (Clock pattern). Production
 * uses Rng.system (ThreadLocalRandom); tests and reproducible pipelines use Rng.seeded(seed),
 * which is deterministic.
 */
class RandomFunctionsSpec extends ScalaCheckSuite:

  private val sheet = Sheet(SheetName.unsafe("S"))

  private def evalNumWith(formula: String, rng: Rng): BigDecimal =
    sheet.evaluateFormula(formula, Clock.system, rng) match
      case Right(CellValue.Number(n)) => n
      case other => fail(s"expected Number for $formula, got $other")

  // ===== Rng capability =====

  test("Rng.seeded is deterministic: same seed, same sequence") {
    val a = Rng.seeded(42L)
    val b = Rng.seeded(42L)
    val seqA = List.fill(10)(a.nextDouble())
    val seqB = List.fill(10)(b.nextDouble())
    assertEquals(seqA, seqB)
  }

  test("Rng.seeded with different seeds diverges") {
    val a = Rng.seeded(1L)
    val b = Rng.seeded(2L)
    assert(List.fill(5)(a.nextDouble()) != List.fill(5)(b.nextDouble()))
  }

  test("Rng.system produces values in [0, 1)") {
    val rng = Rng.system
    (1 to 100).foreach { _ =>
      val d = rng.nextDouble()
      assert(d >= 0.0 && d < 1.0, s"out of range: $d")
    }
  }

  // ===== Registry =====

  test("RAND and RANDBETWEEN are registered (104 -> 106 functions)") {
    val functions = FunctionRegistry.allNames
    assert(functions.contains("RAND"))
    assert(functions.contains("RANDBETWEEN"))
    assertEquals(functions.length, 106)
  }

  // ===== RAND =====

  test("RAND() with the same seed evaluates to the same value") {
    val first = evalNumWith("=RAND()", Rng.seeded(42L))
    val second = evalNumWith("=RAND()", Rng.seeded(42L))
    assertEquals(first, second)
  }

  property("RAND() is always in [0, 1)") {
    forAll(Gen.long) { (seed: Long) =>
      val n = evalNumWith("=RAND()", Rng.seeded(seed))
      n >= 0 && n < 1
    }
  }

  test("RAND() composes in expressions: RAND()*0 = 0") {
    assertEquals(evalNumWith("=RAND()*0", Rng.seeded(7L)), BigDecimal(0))
  }

  test("RAND() with system rng works via the default evaluateFormula path") {
    sheet.evaluateFormula("=RAND()") match
      case Right(CellValue.Number(n)) => assert(n >= 0 && n < 1, s"out of range: $n")
      case other => fail(s"expected Number, got $other")
  }

  // ===== RANDBETWEEN =====

  property("RANDBETWEEN(lo, hi) is an integer in [lo, hi] inclusive") {
    val boundsGen = for
      lo <- Gen.choose(-1000, 1000)
      hi <- Gen.choose(lo, 1000)
    yield (lo, hi)
    forAll(Gen.long, boundsGen) { (seed: Long, bounds: (Int, Int)) =>
      val (lo, hi) = bounds
      val n = evalNumWith(s"=RANDBETWEEN($lo, $hi)", Rng.seeded(seed))
      n.isWhole && n >= lo && n <= hi
    }
  }

  test("RANDBETWEEN(5, 5) = 5") {
    assertEquals(evalNumWith("=RANDBETWEEN(5, 5)", Rng.seeded(1L)), BigDecimal(5))
  }

  test("RANDBETWEEN with negative range") {
    val n = evalNumWith("=RANDBETWEEN(-10, -5)", Rng.seeded(3L))
    assert(n.isWhole && n >= -10 && n <= -5, s"out of range: $n")
  }

  test("RANDBETWEEN(10, 1) with bottom > top is an evaluation error") {
    val result = sheet.evaluateFormula("=RANDBETWEEN(10, 1)", Clock.system, Rng.seeded(1L))
    assert(result.isLeft, s"expected error, got $result")
  }

  test("RANDBETWEEN is deterministic under a fixed seed") {
    val first = evalNumWith("=RANDBETWEEN(1, 1000000)", Rng.seeded(99L))
    val second = evalNumWith("=RANDBETWEEN(1, 1000000)", Rng.seeded(99L))
    assertEquals(first, second)
  }

  test("RANDBETWEEN inside LET binds once per evaluation") {
    // x is computed once and reused: x - x = 0 regardless of the draw
    assertEquals(evalNumWith("=LET(x, RANDBETWEEN(1, 100), x-x)", Rng.seeded(11L)), BigDecimal(0))
  }

  // ===== Round-trip =====

  test("parse . print = id for RAND and RANDBETWEEN") {
    List("=RAND()", "=RANDBETWEEN(1, 10)").foreach { f =>
      val parsed = FormulaParser.parse(f)
      assert(parsed.isRight, s"parse failed for $f: $parsed")
      parsed.foreach { expr =>
        assertEquals(FormulaParser.parse(FormulaPrinter.print(expr)), Right(expr))
      }
    }
  }

  // ===== Workbook-level plumbing =====

  test("recalculate(clock, rng) caches RAND deterministically") {
    val s = Sheet(SheetName.unsafe("S")).put(ref"A1", CellValue.Formula("=RAND()", None))
    val r1 = Workbook(Vector(s)).recalculate(Clock.system, Rng.seeded(5L))
    val r2 = Workbook(Vector(s)).recalculate(Clock.system, Rng.seeded(5L))
    assert(r1.isClean && r2.isClean, s"errors: ${r1.errors} ${r2.errors}")
    val v1 = r1.evaluated(SheetName.unsafe("S")).get(ref"A1")
    val v2 = r2.evaluated(SheetName.unsafe("S")).get(ref"A1")
    assert(v1.isDefined, "A1 not evaluated")
    assertEquals(v1, v2)
  }

  test("wb.evaluateFormula(formula, onSheet, clock, rng) seeded overload") {
    val s = Sheet(SheetName.unsafe("S"))
    val wb = Workbook(Vector(s))
    val a = wb.evaluateFormula("=RANDBETWEEN(1, 10)", "S", Clock.system, Rng.seeded(8L))
    val b = wb.evaluateFormula("=RANDBETWEEN(1, 10)", "S", Clock.system, Rng.seeded(8L))
    assert(a.isRight, s"$a")
    assertEquals(a, b)
  }
