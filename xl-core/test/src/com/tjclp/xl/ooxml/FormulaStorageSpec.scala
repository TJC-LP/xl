package com.tjclp.xl.ooxml

import munit.FunSuite

/**
 * GH-556: the `_xlfn.` storage form of post-2007 functions. `toStored` is the writer's rule (model
 * → `<f>`), `fromStored` the reader's (`<f>` → model), `bareFunctionName` the parser's.
 */
class FormulaStorageSpec extends FunSuite:

  test("toStored: strips the leading '=' and prefixes future functions") {
    assertEquals(
      FormulaStorage.toStored("=MAXIFS(B1:B3,A1:A3,\"a\")"),
      "_xlfn.MAXIFS(B1:B3,A1:A3,\"a\")"
    )
    assertEquals(
      FormulaStorage.toStored("XLOOKUP(C1,A1:A3,B1:B3)"),
      "_xlfn.XLOOKUP(C1,A1:A3,B1:B3)"
    )
    assertEquals(FormulaStorage.toStored("IFS(A1>0,1,TRUE,0)"), "_xlfn.IFS(A1>0,1,TRUE,0)")
    assertEquals(FormulaStorage.toStored("LET(x,1,x+1)"), "_xlfn.LET(x,1,x+1)")
  }

  test("toStored: FILTER and SORT take the worksheet-scoped _xlfn._xlws. prefix") {
    assertEquals(FormulaStorage.toStored("FILTER(A1:A3,B1:B3)"), "_xlfn._xlws.FILTER(A1:A3,B1:B3)")
    assertEquals(FormulaStorage.toStored("SORT(A1:A3)"), "_xlfn._xlws.SORT(A1:A3)")
    // SORTBY is a plain _xlfn. function
    assertEquals(FormulaStorage.toStored("SORTBY(A1:A3,B1:B3)"), "_xlfn.SORTBY(A1:A3,B1:B3)")
  }

  test("toStored: Excel 2007 functions and everything else are untouched") {
    val plain = "IF(SUM(A1:A3)>0,VLOOKUP(A1,B1:C3,2,FALSE),SUMIFS(C1:C3,A1:A3,\">1\"))"
    assertEquals(FormulaStorage.toStored(plain), plain)
    assertEquals(FormulaStorage.toStored("=A1*2"), "A1*2")
  }

  test("toStored: nested and repeated calls are each prefixed") {
    assertEquals(
      FormulaStorage.toStored("IFERROR(XLOOKUP(A1,B:B,C:C),MAXIFS(C:C,B:B,A1))"),
      "IFERROR(_xlfn.XLOOKUP(A1,B:B,C:C),_xlfn.MAXIFS(C:C,B:B,A1))"
    )
  }

  test("toStored: idempotent and never double-prefixes (any case)") {
    val once = FormulaStorage.toStored("MAXIFS(A1:A3,B1:B3,1)")
    assertEquals(FormulaStorage.toStored(once), once)
    assertEquals(FormulaStorage.toStored("_XLFN.MAXIFS(A1)"), "_XLFN.MAXIFS(A1)")
    assertEquals(
      FormulaStorage.toStored("_xlfn._xlws.FILTER(A1:A3,B1:B3)"),
      "_xlfn._xlws.FILTER(A1:A3,B1:B3)"
    )
  }

  test("toStored: the author's case is kept; whitespace before '(' is allowed") {
    assertEquals(FormulaStorage.toStored("maxifs(A1:A3,B1:B3,1)"), "_xlfn.maxifs(A1:A3,B1:B3,1)")
    assertEquals(FormulaStorage.toStored("MAXIFS (A1:A3,B1:B3,1)"), "_xlfn.MAXIFS (A1:A3,B1:B3,1)")
  }

  test("toStored: string literals, quoted sheet names and bracketed references are opaque") {
    assertEquals(
      FormulaStorage.toStored("IF(A1=\"MAXIFS(\",\"\"\"IFS(x)\",1)"),
      "IF(A1=\"MAXIFS(\",\"\"\"IFS(x)\",1)"
    )
    assertEquals(FormulaStorage.toStored("'IFS (2024)'!A1+1"), "'IFS (2024)'!A1+1")
    assertEquals(FormulaStorage.toStored("'It''s IFS (x)'!A1"), "'It''s IFS (x)'!A1")
    assertEquals(FormulaStorage.toStored("SUM(Table1[IFS (x)])"), "SUM(Table1[IFS (x)])")
    assertEquals(FormulaStorage.toStored("[1]Sheet1!A1"), "[1]Sheet1!A1")
  }

  test("toStored: a future-function NAME that is not a call is not a function") {
    // a defined name or LET binding may legally be spelled like a function
    assertEquals(FormulaStorage.toStored("IFS+1"), "IFS+1")
    assertEquals(FormulaStorage.toStored("LET(IFS,1,IFS+1)"), "_xlfn.LET(IFS,1,IFS+1)")
  }

  test("fromStored: strips the prefixes of known future functions") {
    assertEquals(
      FormulaStorage.fromStored("_xlfn.MAXIFS(B1:B3,A1:A3,\"a\")"),
      "MAXIFS(B1:B3,A1:A3,\"a\")"
    )
    assertEquals(
      FormulaStorage.fromStored("_xlfn._xlws.FILTER(A1:A3,B1:B3)"),
      "FILTER(A1:A3,B1:B3)"
    )
    assertEquals(FormulaStorage.fromStored("_XLFN.XLOOKUP(A1,B:B,C:C)"), "XLOOKUP(A1,B:B,C:C)")
    assertEquals(
      FormulaStorage.fromStored("IFERROR(_xlfn.XLOOKUP(A1,B:B,C:C),_xlfn.MAXIFS(C:C,B:B,A1))"),
      "IFERROR(XLOOKUP(A1,B:B,C:C),MAXIFS(C:C,B:B,A1))"
    )
  }

  test("fromStored: a prefix on an unknown function is kept so the writer cannot lose it") {
    assertEquals(FormulaStorage.fromStored("_xlfn.NOSUCHFN(A1)"), "_xlfn.NOSUCHFN(A1)")
    assertEquals(
      FormulaStorage.toStored(FormulaStorage.fromStored("_xlfn.NOSUCHFN(A1)")),
      "_xlfn.NOSUCHFN(A1)"
    )
  }

  test("fromStored ∘ toStored = id on bare model text; toStored ∘ fromStored = id on file text") {
    val model = List(
      "MAXIFS(B1:B3,A1:A3,\"a\")",
      "IFERROR(XLOOKUP(A1,B:B,C:C),0)",
      "FILTER(A1:A3,B1:B3)",
      "SUM(A1:A3)",
      "LET(x,1,x+1)"
    )
    model.foreach(m => assertEquals(FormulaStorage.fromStored(FormulaStorage.toStored(m)), m, m))
    val file = model.map(FormulaStorage.toStored) :+ "_xlfn.NOSUCHFN(A1)"
    file.foreach(f => assertEquals(FormulaStorage.toStored(FormulaStorage.fromStored(f)), f, f))
  }

  test("bareFunctionName: drops every _xlfn. / _xlws. prefix, case-insensitively") {
    assertEquals(FormulaStorage.bareFunctionName("_xlfn.XLOOKUP"), "XLOOKUP")
    assertEquals(FormulaStorage.bareFunctionName("_XLFN._XLWS.FILTER"), "FILTER")
    assertEquals(FormulaStorage.bareFunctionName("_xlws.SORT"), "SORT")
    assertEquals(FormulaStorage.bareFunctionName("SUM"), "SUM")
  }

  test("the registry-known future functions are all in the list") {
    val known =
      List("MAXIFS", "MINIFS", "IFS", "SWITCH", "IFNA", "XLOOKUP", "SEQUENCE", "UNIQUE", "LET")
    known.foreach(n => assert(FormulaStorage.FutureFunctions.contains(n), n))
    assert(FormulaStorage.WorksheetScoped.subsetOf(FormulaStorage.FutureFunctions))
    // Excel 2007 functions must NOT be prefixed
    List("SUMIFS", "COUNTIFS", "AVERAGEIFS", "IFERROR", "SUM", "VLOOKUP").foreach { n =>
      assert(!FormulaStorage.FutureFunctions.contains(n), n)
    }
  }
