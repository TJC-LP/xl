package com.tjclp.xl.ooxml

import munit.FunSuite

/**
 * GH-604: the storage form of the implicit-intersection operator. Excel 365 stores `@x` as
 * `_xlfn.SINGLE(x)`; the model keeps the formula-bar spelling. `toStored` wraps every `@` operand,
 * `fromStored` unwraps every one-argument `SINGLE(...)`, and the lint sees a bare `@` as the SINGLE
 * call the writer would spell out. GH-603: omitted-argument commas pass through untouched.
 */
class FormulaStorageIntersectionSpec extends FunSuite:

  private val toStored: String => String = FormulaStorage.toStored
  private val fromStored: String => String = FormulaStorage.fromStored
  private val bareFutureCalls: String => Vector[String] = FormulaStorage.bareFutureCalls

  /**
   * GH-577's invariant: the lint reports something exactly when the writer would change the text.
   */
  private def pinLintInvariant(text: String): Unit =
    assertEquals(
      bareFutureCalls(text).isEmpty,
      toStored(text) == text.stripPrefix("="),
      s"bareFutureCalls and toStored disagree on: $text"
    )

  private def roundTrip(model: String, stored: String): Unit =
    pinLintInvariant(model)
    pinLintInvariant(stored)
    assertEquals(toStored(model), stored, s"toStored($model)")
    assertEquals(fromStored(stored), model, s"fromStored($stored)")
    assertEquals(toStored(stored), stored, s"toStored is idempotent on $stored")
    assertEquals(toStored(fromStored(stored)), stored, s"toStored ∘ fromStored on $stored")

  test("@name ↔ _xlfn.SINGLE(name)") {
    roundTrip("@acq", "_xlfn.SINGLE(acq)")
    assertEquals(toStored("=@acq"), "_xlfn.SINGLE(acq)")
    roundTrip("IF(@acq=1,'M&A'!I12,0)", "IF(_xlfn.SINGLE(acq)=1,'M&A'!I12,0)")
    roundTrip("@acq+@case*2", "_xlfn.SINGLE(acq)+_xlfn.SINGLE(case)*2")
  }

  test("qualified and anchored operands wrap whole") {
    roundTrip("@'My Sheet'!A1:A10", "_xlfn.SINGLE('My Sheet'!A1:A10)")
    roundTrip("@Sheet1!A1:A10", "_xlfn.SINGLE(Sheet1!A1:A10)")
    roundTrip("@$A$1:$A$10", "_xlfn.SINGLE($A$1:$A$10)")
    roundTrip("@[1]Book!A1", "_xlfn.SINGLE([1]Book!A1)")
    roundTrip("@Sales.Total", "_xlfn.SINGLE(Sales.Total)")
  }

  test("a call operand wraps whole, and its own future-function prefix is applied inside") {
    roundTrip("@INDEX(A1:B2,1,)", "_xlfn.SINGLE(INDEX(A1:B2,1,))")
    roundTrip("@XLOOKUP(1,A:A,B:B)", "_xlfn.SINGLE(_xlfn.XLOOKUP(1,A:A,B:B))")
    roundTrip("@Table1[Col]", "_xlfn.SINGLE(Table1[Col])")
  }

  test("a parenthesized operand lends its parens to the call") {
    roundTrip("@(A1:A3*2)", "_xlfn.SINGLE(A1:A3*2)")
    roundTrip("SUM(@(A1:A3*2),1)", "SUM(_xlfn.SINGLE(A1:A3*2),1)")
  }

  test("@ inside a structured reference or a string is copied verbatim") {
    assertEquals(toStored("SUM(Table1[@Col])"), "SUM(Table1[@Col])")
    assertEquals(toStored("\"a@b\"&A1"), "\"a@b\"&A1")
    assertEquals(toStored("'a@b'!A1"), "'a@b'!A1")
    assertEquals(bareFutureCalls("SUM(Table1[@Col])"), Vector.empty)
    // an '@' with nothing operand-shaped after it is left alone
    assertEquals(toStored("A1&\"@\""), "A1&\"@\"")
    assertEquals(toStored("@"), "@")
    assertEquals(toStored("@+1"), "@+1")
  }

  test("the lint reports a bare @ as SINGLE; the stored form is clean") {
    assertEquals(bareFutureCalls("@acq"), Vector("SINGLE"))
    assertEquals(bareFutureCalls("_xlfn.SINGLE(acq)"), Vector.empty)
    // a bare-writing producer's SINGLE(x) is prefixed like any other future function
    assertEquals(toStored("SINGLE(acq)"), "_xlfn.SINGLE(acq)")
    assertEquals(bareFutureCalls("SINGLE(acq)"), Vector("SINGLE"))
    assertEquals(bareFutureCalls("@acq+IFS(A1,1)"), Vector("SINGLE", "IFS"))
  }

  test("only a one-argument SINGLE unwraps to @; the model form SINGLE(x) reads back as @x") {
    assertEquals(fromStored("_xlfn.SINGLE(A1,B1)"), "SINGLE(A1,B1)")
    assertEquals(fromStored("_xlfn.SINGLE(acq)"), "@acq")
    assertEquals(fromStored("_xlfn.single(acq)"), "@acq")
    assertEquals(fromStored("_xlfn.SINGLE(A1:A3*2)"), "@(A1:A3*2)")
    // a different call with a matching suffix is not SINGLE
    assertEquals(fromStored("_xlfn.XLOOKUP(1,A:A,B:B)"), "XLOOKUP(1,A:A,B:B)")
    assertEquals(fromStored("MYSINGLE(A1)+_xlfn.IFS(1,2)"), "MYSINGLE(A1)+IFS(1,2)")
  }

  test("GH-603: omitted-argument commas and GH-605: RRI pass through the storage boundary") {
    assertEquals(toStored("RATE($G$8-$D$8,,-D9,G9,)"), "RATE($G$8-$D$8,,-D9,G9,)")
    assertEquals(fromStored("RATE($G$8-$D$8,,-D9,G9,)"), "RATE($G$8-$D$8,,-D9,G9,)")
    roundTrip("RRI(10,100,150)", "_xlfn.RRI(10,100,150)")
    roundTrip("IFERROR(RRI($G$8-$D$8,,G9),0)", "IFERROR(_xlfn.RRI($G$8-$D$8,,G9),0)")
  }
