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
  private val corruptSpillQualifiers: String => Vector[String] =
    FormulaStorage.corruptSpillQualifiers

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

  // ===== GH-687: `X!#REF!` — a sheet qualifier with no cell reference is never a spill operand =====

  /** Unchanged through both directions, and the lint reports nothing. */
  private def untouched(text: String): Unit =
    pinLintInvariant(text)
    assertEquals(toStored(text), text.stripPrefix("="), s"toStored($text)")
    assertEquals(fromStored(text), text, s"fromStored($text)")
    assertEquals(bareFutureCalls(text), Vector.empty, s"bareFutureCalls($text)")

  test("GH-687: Excel's Sheet!#REF! (a deleted target) is an error literal, not a spill") {
    untouched("Sheet1!#REF!")
    untouched("'My Sheet'!#REF!")
    untouched("'It''s'!#REF!")
    untouched("[1]Sheet1!#REF!")
    untouched("[1]'My Sheet'!#REF!")
    untouched("Support!#REF!")
    untouched("Sheet1!$A$1,Sheet1!#REF!")
    untouched("#REF!")
    untouched("=SUM(Sheet1!#REF!,A1)")
    untouched("=Sheet1!#REF!+1")
    untouched("IF(A1,Sheet1!#REF!,1)")
    untouched("Sheet_2.x!#REF!")
    assertEquals(toStored("=SUM(Sheet1!#REF!,A1)"), "SUM(Sheet1!#REF!,A1)")
    assertEquals(toStored("=Sheet1!#REF!+1"), "Sheet1!#REF!+1")
  }

  test("GH-687: every error literal after ! stays verbatim (a # right after ! has no operand)") {
    for err <- Vector("#N/A", "#NAME?", "#DIV/0!", "#NULL!", "#NUM!", "#VALUE!", "#SPILL!") do
      untouched(s"Sheet1!$err")
      untouched(s"'My Sheet'!$err")
      untouched(s"[1]Sheet1!$err")
    // a lone '#' after '!' is not a spill of the qualifier either
    untouched("Sheet1!#")
  }

  test("GH-687: a real spill still wraps beside a Sheet!#REF! in the same formula") {
    roundTrip("SUM(A1#,Sheet1!#REF!)", "SUM(_xlfn.ANCHORARRAY(A1),Sheet1!#REF!)")
    roundTrip("Sheet1!#REF!+Sheet1!A1#", "Sheet1!#REF!+_xlfn.ANCHORARRAY(Sheet1!A1)")
    roundTrip("'My Sheet'!$B$2#", "_xlfn.ANCHORARRAY('My Sheet'!$B$2)")
    roundTrip("SUM(A1#)", "SUM(_xlfn.ANCHORARRAY(A1))")
    assertEquals(bareFutureCalls("SUM(A1#,Sheet1!#REF!)"), Vector("#"))
  }

  test("GH-687: an @ operand of Sheet!#REF! is the whole error literal") {
    roundTrip("@Sheet1!#REF!", "_xlfn.SINGLE(Sheet1!#REF!)")
    roundTrip("@'My Sheet'!#REF!+1", "_xlfn.SINGLE('My Sheet'!#REF!)+1")
    assertEquals(bareFutureCalls("@Sheet1!#REF!"), Vector("@"))
  }

  // ===== GH-687 heal-on-read: the 0.23.0–0.23.1 corruption reads back as Excel's spelling =====

  /**
   * `stored` (what xl 0.23.x wrote) reads back as `model` (what Excel wrote before it), and the
   * next write stores Excel's spelling again — the write side stays strict, so the heal is a
   * read-side property only.
   */
  private def heals(stored: String, model: String, restored: String): Unit =
    assertEquals(fromStored(stored), model, s"fromStored($stored)")
    assertEquals(toStored(fromStored(stored)), restored, s"toStored ∘ fromStored on $stored")
    // the restored text is a fixpoint: it reads back as the same model
    assertEquals(fromStored(restored), model, s"fromStored($restored)")

  test("GH-687: _xlfn.ANCHORARRAY(Sheet!)REF! (xl 0.23.x's corruption) heals to Sheet!#REF!") {
    heals("_xlfn.ANCHORARRAY(Support!)REF!", "Support!#REF!", "Support!#REF!")
    heals("_xlfn.ANCHORARRAY(Sheet1!)REF!", "Sheet1!#REF!", "Sheet1!#REF!")
    heals("_xlfn.ANCHORARRAY('My Sheet'!)REF!", "'My Sheet'!#REF!", "'My Sheet'!#REF!")
    heals("_xlfn.ANCHORARRAY('It''s'!)REF!", "'It''s'!#REF!", "'It''s'!#REF!")
    heals("_xlfn.ANCHORARRAY([1]Sheet1!)REF!", "[1]Sheet1!#REF!", "[1]Sheet1!#REF!")
    heals("_xlfn.ANCHORARRAY([1]'My Sheet'!)REF!", "[1]'My Sheet'!#REF!", "[1]'My Sheet'!#REF!")
    heals("_xlfn.ANCHORARRAY('[1]My Sheet'!)REF!", "'[1]My Sheet'!#REF!", "'[1]My Sheet'!#REF!")
    heals("_xlfn.ANCHORARRAY(Sheet_2.x!)REF!", "Sheet_2.x!#REF!", "Sheet_2.x!#REF!")
    // the prefix is matched case-insensitively, like every other storage prefix
    heals("_xlfn.anchorarray(Sheet1!)REF!", "Sheet1!#REF!", "Sheet1!#REF!")
    // a bare qualifier with nothing after it: `Sheet1!#`, the text 0.23.x started from
    heals("_xlfn.ANCHORARRAY(Sheet1!)", "Sheet1!#", "Sheet1!#")
  }

  test("GH-687: every error code after the corrupted qualifier heals") {
    for err <- Vector("N/A", "NAME?", "DIV/0!", "NULL!", "NUM!", "VALUE!", "SPILL!", "REF!") do
      heals(s"_xlfn.ANCHORARRAY(Sheet1!)$err", s"Sheet1!#$err", s"Sheet1!#$err")
      heals(s"_xlfn.ANCHORARRAY('My Sheet'!)$err", s"'My Sheet'!#$err", s"'My Sheet'!#$err")
      heals(s"_xlfn.ANCHORARRAY([1]Sheet1!)$err", s"[1]Sheet1!#$err", s"[1]Sheet1!#$err")
  }

  test("GH-687: the corruption heals inside larger formulas, beside real spills and calls") {
    heals("SUM(_xlfn.ANCHORARRAY(Sheet1!)REF!,A1)", "SUM(Sheet1!#REF!,A1)", "SUM(Sheet1!#REF!,A1)")
    heals("_xlfn.ANCHORARRAY(Sheet1!)REF!+1", "Sheet1!#REF!+1", "Sheet1!#REF!+1")
    heals(
      "IF(A1,_xlfn.ANCHORARRAY(Sheet1!)REF!,_xlfn.ANCHORARRAY('My Sheet'!)N/A)",
      "IF(A1,Sheet1!#REF!,'My Sheet'!#N/A)",
      "IF(A1,Sheet1!#REF!,'My Sheet'!#N/A)"
    )
    heals(
      "Sheet1!$A$1,_xlfn.ANCHORARRAY(Sheet1!)REF!",
      "Sheet1!$A$1,Sheet1!#REF!",
      "Sheet1!$A$1,Sheet1!#REF!"
    )
    // a 3-D reference: only the trailing sheet was a token start for the 0.23.x wrap
    heals("Sheet1:_xlfn.ANCHORARRAY(Sheet3!)REF!", "Sheet1:Sheet3!#REF!", "Sheet1:Sheet3!#REF!")
    // a genuine spill and a future function beside the healed literal keep their own storage
    heals(
      "SUM(_xlfn.ANCHORARRAY(A1),_xlfn.ANCHORARRAY(Sheet1!)REF!)",
      "SUM(A1#,Sheet1!#REF!)",
      "SUM(_xlfn.ANCHORARRAY(A1),Sheet1!#REF!)"
    )
    heals(
      "_xlfn.XLOOKUP(1,_xlfn.ANCHORARRAY(Sheet1!)REF!,B1:B3)",
      "XLOOKUP(1,Sheet1!#REF!,B1:B3)",
      "_xlfn.XLOOKUP(1,Sheet1!#REF!,B1:B3)"
    )
  }

  test("GH-687: @Sheet1!#REF! stored as SINGLE(ANCHORARRAY(Sheet1!))REF! heals to the @ form") {
    heals(
      "_xlfn.SINGLE(_xlfn.ANCHORARRAY(Sheet1!))REF!",
      "@Sheet1!#REF!",
      "_xlfn.SINGLE(Sheet1!#REF!)"
    )
    heals(
      "_xlfn.SINGLE(_xlfn.ANCHORARRAY('My Sheet'!))REF!+1",
      "@'My Sheet'!#REF!+1",
      "_xlfn.SINGLE('My Sheet'!#REF!)+1"
    )
  }

  test(
    "GH-687: a genuine ANCHORARRAY still reads as a spill; a string or non-qualifier never heals"
  ) {
    assertEquals(fromStored("_xlfn.ANCHORARRAY(A1)"), "A1#")
    assertEquals(fromStored("_xlfn.ANCHORARRAY(Sheet1!A1)"), "Sheet1!A1#")
    assertEquals(fromStored("_xlfn.ANCHORARRAY('My Sheet'!$A$1)"), "'My Sheet'!$A$1#")
    assertEquals(fromStored("_xlfn.ANCHORARRAY([1]Sheet1!A1)"), "[1]Sheet1!A1#")
    // not a bare qualifier: the call keeps its spelling
    assertEquals(fromStored("_xlfn.ANCHORARRAY(Sheet1!,A1)"), "ANCHORARRAY(Sheet1!,A1)")
    assertEquals(fromStored("_xlfn.ANCHORARRAY(!)"), "ANCHORARRAY(!)")
    assertEquals(fromStored("_xlfn.ANCHORARRAY('My Sheet')"), "ANCHORARRAY('My Sheet')")
    assertEquals(fromStored("_xlfn.ANCHORARRAY(1+Sheet1!)"), "ANCHORARRAY(1+Sheet1!)")
    // inside a string literal nothing is rewritten
    assertEquals(
      fromStored("\"_xlfn.ANCHORARRAY(Sheet1!)REF!\"&_xlfn.IFS(1,2)"),
      "\"_xlfn.ANCHORARRAY(Sheet1!)REF!\"&IFS(1,2)"
    )
  }

  test("GH-687: corruptSpillQualifiers names exactly what fromStored heals") {
    assertEquals(corruptSpillQualifiers("_xlfn.ANCHORARRAY(Support!)REF!"), Vector("Support!"))
    assertEquals(
      corruptSpillQualifiers(
        "IF(A1,_xlfn.ANCHORARRAY(Sheet1!)REF!,_xlfn.ANCHORARRAY('My Sheet'!)N/A)"
      ),
      Vector("Sheet1!", "'My Sheet'!")
    )
    assertEquals(
      corruptSpillQualifiers("_xlfn.SINGLE(_xlfn.ANCHORARRAY([1]Sheet1!))REF!"),
      Vector("[1]Sheet1!")
    )
    // Excel's own spellings, genuine spills and strings report nothing
    Vector(
      "Sheet1!#REF!",
      "_xlfn.SINGLE(Sheet1!#REF!)",
      "_xlfn.ANCHORARRAY(A1)",
      "_xlfn.ANCHORARRAY(Sheet1!A1)",
      "SUM(_xlfn.ANCHORARRAY(A1),Sheet1!#REF!)",
      "\"_xlfn.ANCHORARRAY(Sheet1!)REF!\"",
      "_xlfn.IFS(1,2)",
      ""
    ).foreach(t => assertEquals(corruptSpillQualifiers(t), Vector.empty, t))
  }

  test("GH-687: storageNeedsHealing = bare future call OR corrupt ANCHORARRAY qualifier") {
    assert(FormulaStorage.storageNeedsHealing("_xlfn.ANCHORARRAY(Support!)REF!"))
    assert(FormulaStorage.storageNeedsHealing("IFS(1,2)"))
    assert(!FormulaStorage.storageNeedsHealing("Support!#REF!"))
    assert(!FormulaStorage.storageNeedsHealing("_xlfn.ANCHORARRAY(A1)"))
    // the corruption is not a bare future call: xlfn-missing stays silent on it
    assertEquals(bareFutureCalls("_xlfn.ANCHORARRAY(Support!)REF!"), Vector.empty)
  }
