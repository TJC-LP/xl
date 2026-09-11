package com.tjclp.xl.ooxml

import munit.FunSuite

/**
 * GH-655: the storage form of the spill reference. Excel 365 stores `x#` as `_xlfn.ANCHORARRAY(x)`;
 * the model keeps the formula-bar spelling. `toStored` wraps every `reference#`, `fromStored`
 * unwraps every `ANCHORARRAY(reference)`, and the lint records a bare `#` as the token the file
 * holds. Nested under `@` both directions rescan (`@A1#` ↔ `_xlfn.SINGLE(_xlfn.ANCHORARRAY(A1))`).
 */
class FormulaStorageSpillSpec extends FunSuite:

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

  test("reference# ↔ _xlfn.ANCHORARRAY(reference) for every reference shape") {
    roundTrip("SUM(A1#)", "SUM(_xlfn.ANCHORARRAY(A1))")
    roundTrip("A1#*2", "_xlfn.ANCHORARRAY(A1)*2")
    roundTrip("$A$1#", "_xlfn.ANCHORARRAY($A$1)")
    roundTrip("Sheet1!A1#", "_xlfn.ANCHORARRAY(Sheet1!A1)")
    roundTrip("'My Sheet'!$A$1#", "_xlfn.ANCHORARRAY('My Sheet'!$A$1)")
    roundTrip("[1]Book!A1#", "_xlfn.ANCHORARRAY([1]Book!A1)")
    roundTrip("spill#", "_xlfn.ANCHORARRAY(spill)")
    roundTrip("Sales.Total#", "_xlfn.ANCHORARRAY(Sales.Total)")
    roundTrip("IF(A1#>1,1,0)", "IF(_xlfn.ANCHORARRAY(A1)>1,1,0)")
    roundTrip("SUM(A1#,B1#)", "SUM(_xlfn.ANCHORARRAY(A1),_xlfn.ANCHORARRAY(B1))")
    assertEquals(toStored("=SUM(A1#)"), "SUM(_xlfn.ANCHORARRAY(A1))")
  }

  test("under @ both directions rescan: @A1# ↔ _xlfn.SINGLE(_xlfn.ANCHORARRAY(A1))") {
    roundTrip("@A1#", "_xlfn.SINGLE(_xlfn.ANCHORARRAY(A1))")
    roundTrip("@Sheet1!A1#+1", "_xlfn.SINGLE(_xlfn.ANCHORARRAY(Sheet1!A1))+1")
    roundTrip(
      "@XLOOKUP(1,A1#,B1#)",
      "_xlfn.SINGLE(_xlfn.XLOOKUP(1,_xlfn.ANCHORARRAY(A1),_xlfn.ANCHORARRAY(B1)))"
    )
  }

  test("a # inside a structured reference, a string or an error literal is copied verbatim") {
    assertEquals(toStored("SUM(Table1[#All])"), "SUM(Table1[#All])")
    assertEquals(toStored("SUM(Table1[[#Headers],[Col]])"), "SUM(Table1[[#Headers],[Col]])")
    assertEquals(toStored("\"a#\"&A1"), "\"a#\"&A1")
    assertEquals(toStored("IF(A1,#REF!,1)"), "IF(A1,#REF!,1)")
    assertEquals(toStored("#N/A"), "#N/A")
    assertEquals(toStored("'a#b'!A1"), "'a#b'!A1")
    assertEquals(bareFutureCalls("SUM(Table1[#All])"), Vector.empty)
    assertEquals(bareFutureCalls("IF(A1,#REF!,1)"), Vector.empty)
    // a '#' with nothing reference-shaped before it is left alone
    assertEquals(toStored("1+#"), "1+#")
    assertEquals(toStored("SUM(A1:A3)#"), "SUM(A1:A3)#")
  }

  test("the lint reports a bare x# as the token #; the stored form is clean") {
    assertEquals(bareFutureCalls("SUM(A1#)"), Vector("#"))
    assertEquals(bareFutureCalls("@A1#"), Vector("@", "#"))
    assertEquals(bareFutureCalls("A1#+IFS(A1,1)"), Vector("#", "IFS"))
    assertEquals(bareFutureCalls("_xlfn.ANCHORARRAY(A1)"), Vector.empty)
    // a bare-writing producer's ANCHORARRAY(x) is prefixed like any other future function
    assertEquals(toStored("ANCHORARRAY(A1)"), "_xlfn.ANCHORARRAY(A1)")
    assertEquals(bareFutureCalls("ANCHORARRAY(A1)"), Vector("ANCHORARRAY"))
  }

  test("only ANCHORARRAY over one reference unwraps to #; any other argument keeps the call") {
    assertEquals(fromStored("_xlfn.ANCHORARRAY(A1:A3)"), "ANCHORARRAY(A1:A3)")
    assertEquals(fromStored("_xlfn.ANCHORARRAY(A1,B1)"), "ANCHORARRAY(A1,B1)")
    assertEquals(fromStored("_xlfn.ANCHORARRAY(INDEX(A1:A3,1))"), "ANCHORARRAY(INDEX(A1:A3,1))")
    assertEquals(fromStored("_xlfn.anchorarray(a1)"), "a1#")
    // a different call with a matching suffix is not ANCHORARRAY
    assertEquals(fromStored("MYANCHORARRAY(A1)+_xlfn.IFS(1,2)"), "MYANCHORARRAY(A1)+IFS(1,2)")
  }
