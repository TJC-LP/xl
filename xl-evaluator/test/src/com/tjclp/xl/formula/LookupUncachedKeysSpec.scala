package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellError, CellValue}
import munit.FunSuite

/**
 * GH-715: MATCH, VLOOKUP, HLOOKUP and XLOOKUP read their keys the way a reference reads its cell,
 * lazily and in search order. An uncached formula key evaluates; a key whose evaluation fails is a
 * non-match, so a broken or circular cell after the match changes nothing; a whole column costs the
 * data, not a million reads.
 */
class LookupUncachedKeysSpec extends FunSuite:
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  /** A1:A3 = 1, `=3+4` uncached, 9; B1:B3 = 10, 20, 30. */
  private val column = Sheet("T")
    .put(ref"A1", num(1))
    .put(ref"A2", CellValue.Formula("=3+4"))
    .put(ref"A3", num(9))
    .put(ref"B1", num(10))
    .put(ref"B2", num(20))
    .put(ref"B3", num(30))

  /** The transpose: A1:C1 keys, A2:C2 results. */
  private val row = Sheet("T")
    .put(ref"A1", num(1))
    .put(ref"B1", CellValue.Formula("=3+4"))
    .put(ref"C1", num(9))
    .put(ref"A2", num(10))
    .put(ref"B2", num(20))
    .put(ref"C2", num(30))

  private def eval(sheet: Sheet, formula: String): CellValue =
    sheet.evaluateFormula(formula).fold(e => fail(s"$formula: $e"), identity)

  test("an uncached formula key evaluates and matches, with no placement involved") {
    assertEquals(eval(column, "=VLOOKUP(7,A1:B3,2,FALSE)"), num(20))
    assertEquals(eval(row, "=HLOOKUP(7,A1:C2,2,FALSE)"), num(20))
    assertEquals(eval(column, "=MATCH(7,A1:A3,0)"), num(2))
    assertEquals(eval(column, "=XLOOKUP(7,A1:A3,B1:B3)"), num(20))
    assertEquals(eval(column, "=XLOOKUP(7,A1:A3,B1:B3,,0,-1)"), num(20))
  }

  // A4 fails to evaluate: an unknown function, or a formula that reads itself (no positioned
  // evaluation, so it is the recursion guard's failure, not a CircularRef).
  List("=NOSUCHFN(1)", "=A4+1").foreach { broken =>
    val withBroken = column.put(ref"A4", CellValue.Formula(broken)).put(ref"B4", num(40))

    test(s"a key after the match that fails to evaluate ($broken) changes nothing") {
      assertEquals(eval(withBroken, "=VLOOKUP(7,A1:B4,2,FALSE)"), num(20))
      assertEquals(eval(withBroken, "=MATCH(7,A1:A4,0)"), num(2))
      assertEquals(eval(withBroken, "=XLOOKUP(7,A1:A4,B1:B4)"), num(20))
      assertEquals(
        eval(
          row.put(ref"D1", CellValue.Formula(broken.replace("A4", "D1"))),
          "=HLOOKUP(7,A1:D2,2,FALSE)"
        ),
        num(20)
      )
    }

    test(s"a failing key ($broken) is a non-match, so a miss is #N/A, not a failure") {
      assertEquals(eval(withBroken, "=IFNA(MATCH(99,A1:A4,0),-1)"), num(-1))
      assertEquals(eval(withBroken, "=IFNA(VLOOKUP(99,A1:B4,2,FALSE),-1)"), num(-1))
      assertEquals(eval(withBroken, "=XLOOKUP(99,A1:A4,B1:B4,-1)"), num(-1))
    }
  }

  /** A1:A100 = 1..100 ascending, B = twice that. */
  private val tall = (0 until 100).foldLeft(Sheet("T")) { (s, i) =>
    s.put(ARef.from0(0, i), num(i + 1)).put(ARef.from0(1, i), num(2 * (i + 1)))
  }

  test("a whole column searches the used range and still answers every mode") {
    assertEquals(eval(tall, "=MATCH(7,A:A,0)"), num(7))
    assertEquals(eval(tall, "=MATCH(7.5,A:A,1)"), num(7))
    assertEquals(eval(tall, "=XLOOKUP(42,A:A,B:B)"), num(84))
    assertEquals(eval(tall, "=XLOOKUP(42,A:A,B:B,,0,-1)"), num(84))
    assertEquals(eval(tall, "=XLOOKUP(41.5,A:A,B:B,,1)"), num(84))
    assertEquals(eval(tall, "=VLOOKUP(42,A:B,2,FALSE)"), num(84))
    assertEquals(eval(tall, "=VLOOKUP(42.5,A:B,2,TRUE)"), num(84))
    assertEquals(eval(tall, "=IFNA(MATCH(500,A:A,0),-1)"), num(-1))
    assertEquals(eval(tall, "=XLOOKUP(500,A:A,B:B)"), CellValue.Error(CellError.NA))
  }
