package com.tjclp.xl.ooxml

import java.util.Locale

import munit.FunSuite

/**
 * GH-556: the `_xlfn.` storage form of post-2007 functions. `toStored` is the writer's rule (model
 * → `<f>`), `fromStored` the reader's (`<f>` → model), `bareFunctionName` the parser's.
 */
class FormulaStorageSpec extends FunSuite:

  private val toStored: String => String = FormulaStorage.toStored
  private val fromStored: String => String = FormulaStorage.fromStored
  private val bareFutureCalls: String => Vector[String] = FormulaStorage.bareFutureCalls

  /**
   * GH-577: `bareFutureCalls` reports something exactly when `toStored` would change the text — the
   * lint and the writer share one scanner. Every `toStored` / `fromStored` input in this suite goes
   * through [[stored]] / [[restored]], so each of them pins the invariant.
   */
  private def pinLintInvariant(text: String): Unit =
    assertEquals(
      bareFutureCalls(text).isEmpty,
      toStored(text) == text.stripPrefix("="),
      s"bareFutureCalls and toStored disagree on: $text"
    )

  private def stored(expr: String): String =
    pinLintInvariant(expr)
    toStored(expr)

  private def restored(text: String): String =
    pinLintInvariant(text)
    fromStored(text)

  test("toStored: strips the leading '=' and prefixes future functions") {
    assertEquals(
      stored("=MAXIFS(B1:B3,A1:A3,\"a\")"),
      "_xlfn.MAXIFS(B1:B3,A1:A3,\"a\")"
    )
    assertEquals(
      stored("XLOOKUP(C1,A1:A3,B1:B3)"),
      "_xlfn.XLOOKUP(C1,A1:A3,B1:B3)"
    )
    assertEquals(stored("IFS(A1>0,1,TRUE,0)"), "_xlfn.IFS(A1>0,1,TRUE,0)")
  }

  test("toStored: FILTER and SORT take the worksheet-scoped _xlfn._xlws. prefix") {
    assertEquals(stored("FILTER(A1:A3,B1:B3)"), "_xlfn._xlws.FILTER(A1:A3,B1:B3)")
    assertEquals(stored("SORT(A1:A3)"), "_xlfn._xlws.SORT(A1:A3)")
    // SORTBY is a plain _xlfn. function
    assertEquals(stored("SORTBY(A1:A3,B1:B3)"), "_xlfn.SORTBY(A1:A3,B1:B3)")
  }

  test("toStored: Excel 2007 functions and everything else are untouched") {
    val plain = "IF(SUM(A1:A3)>0,VLOOKUP(A1,B1:C3,2,FALSE),SUMIFS(C1:C3,A1:A3,\">1\"))"
    assertEquals(stored(plain), plain)
    assertEquals(stored("=A1*2"), "A1*2")
  }

  test("toStored: nested and repeated calls are each prefixed") {
    assertEquals(
      stored("IFERROR(XLOOKUP(A1,B:B,C:C),MAXIFS(C:C,B:B,A1))"),
      "IFERROR(_xlfn.XLOOKUP(A1,B:B,C:C),_xlfn.MAXIFS(C:C,B:B,A1))"
    )
  }

  test("toStored: idempotent and never double-prefixes (any case)") {
    val once = stored("MAXIFS(A1:A3,B1:B3,1)")
    assertEquals(stored(once), once)
    assertEquals(stored("_XLFN.MAXIFS(A1)"), "_XLFN.MAXIFS(A1)")
    assertEquals(
      stored("_xlfn._xlws.FILTER(A1:A3,B1:B3)"),
      "_xlfn._xlws.FILTER(A1:A3,B1:B3)"
    )
  }

  test("toStored: the author's case is kept; whitespace before '(' is allowed") {
    assertEquals(stored("maxifs(A1:A3,B1:B3,1)"), "_xlfn.maxifs(A1:A3,B1:B3,1)")
    assertEquals(stored("MAXIFS (A1:A3,B1:B3,1)"), "_xlfn.MAXIFS (A1:A3,B1:B3,1)")
  }

  test("toStored/fromStored: lower-case names are recognized under a Turkish default locale") {
    // Locale-sensitive upper-casing turns "ifs" into "İFS" and misses the set — the classic
    // dotted-I trap. Both directions must use the root locale.
    val saved = Locale.getDefault
    try
      Locale.setDefault(Locale.forLanguageTag("tr-TR"))
      assertEquals(stored("ifs(A1>0,1,TRUE,0)"), "_xlfn.ifs(A1>0,1,TRUE,0)")
      assertEquals(stored("minifs(A1:A3,B1:B3,1)"), "_xlfn.minifs(A1:A3,B1:B3,1)")
      assertEquals(restored("_xlfn.ifs(A1>0,1,TRUE,0)"), "ifs(A1>0,1,TRUE,0)")
    finally Locale.setDefault(saved)
  }

  test("toStored: string literals, quoted sheet names and bracketed references are opaque") {
    assertEquals(
      stored("IF(A1=\"MAXIFS(\",\"\"\"IFS(x)\",1)"),
      "IF(A1=\"MAXIFS(\",\"\"\"IFS(x)\",1)"
    )
    assertEquals(stored("'IFS (2024)'!A1+1"), "'IFS (2024)'!A1+1")
    assertEquals(stored("'It''s IFS (x)'!A1"), "'It''s IFS (x)'!A1")
    assertEquals(stored("SUM(Table1[IFS (x)])"), "SUM(Table1[IFS (x)])")
    assertEquals(stored("[1]Sheet1!A1"), "[1]Sheet1!A1")
  }

  test("toStored: bracketed references are opaque even when they carry quote escapes") {
    // A structured-reference column name escapes its specials with a single quote; the scanner
    // must not enter quote mode there, or the rest of the formula would be copied unrewritten.
    assertEquals(
      stored("SUM(Table1['#Sales])+MAXIFS(A1:A3,B1:B3,1)"),
      "SUM(Table1['#Sales])+_xlfn.MAXIFS(A1:A3,B1:B3,1)"
    )
    assertEquals(
      stored("SUM(Table1[a\"b])+IFS(A1>0,1,TRUE,0)"),
      "SUM(Table1[a\"b])+_xlfn.IFS(A1>0,1,TRUE,0)"
    )
  }

  test("toStored: an escaped bracket inside a structured reference keeps the depth honest") {
    // `'[` and `']` are one unit each; counting the escaped bracket would leave the scanner
    // "inside brackets" for the rest of the formula and MAXIFS would land bare (#NAME? in Excel).
    assertEquals(
      stored("SUM(Table1['[Total])+MAXIFS(A1:A3,B1:B3,1)"),
      "SUM(Table1['[Total])+_xlfn.MAXIFS(A1:A3,B1:B3,1)"
    )
    assertEquals(
      stored("SUM(Table1[']Total])+MAXIFS(A1:A3,B1:B3,1)"),
      "SUM(Table1[']Total])+_xlfn.MAXIFS(A1:A3,B1:B3,1)"
    )
    assertEquals(
      stored("SUM(Table1[[#Totals],[Sales]])+MAXIFS(A1:A3,B1:B3,1)"),
      "SUM(Table1[[#Totals],[Sales]])+_xlfn.MAXIFS(A1:A3,B1:B3,1)"
    )
  }

  test("toStored: a lone _xlws. gains the _xlfn. Excel requires in front of it") {
    assertEquals(
      stored("_xlws.FILTER(A1:A3,B1:B3)"),
      "_xlfn._xlws.FILTER(A1:A3,B1:B3)"
    )
    // a third-party `_xlfn.FILTER` reads back bare and is re-written in Excel's own form
    assertEquals(restored("_xlfn.FILTER(A1:A3,B1:B3)"), "FILTER(A1:A3,B1:B3)")
  }

  test("toStored: a future-function NAME that is not a call is not a function") {
    // a defined name may legally be spelled like a function
    assertEquals(stored("IFS+1"), "IFS+1")
    assertEquals(stored("SUM(IFS)+IFS(1,2)"), "SUM(IFS)+_xlfn.IFS(1,2)")
  }

  // ===== LET / LAMBDA parameters: Excel's _xlpm. prefix =====

  test("toStored: LET parameters and their references take _xlpm. (Excel's storage form)") {
    assertEquals(stored("LET(x,1,x+1)"), "_xlfn.LET(_xlpm.x,1,_xlpm.x+1)")
    assertEquals(
      stored("LET(x, 1, y, x*2, x+y)"),
      "_xlfn.LET(_xlpm.x, 1, _xlpm.y, _xlpm.x*2, _xlpm.x+_xlpm.y)"
    )
    // references inside nested calls, and a nested future function, are both rewritten
    assertEquals(
      stored("LET(x,SUM(A1:A3),IF(x>0,MAXIFS(B:B,C:C,x),0))"),
      "_xlfn.LET(_xlpm.x,SUM(A1:A3),IF(_xlpm.x>0,_xlfn.MAXIFS(B:B,C:C,_xlpm.x),0))"
    )
    // names are case-insensitive: X refers to x
    assertEquals(stored("let(x,1,X+1)"), "_xlfn.let(_xlpm.x,1,_xlpm.X+1)")
  }

  test("toStored: LET values and the calculation that are defined names are NOT parameters") {
    assertEquals(stored("LET(x,Rate,x*2)"), "_xlfn.LET(_xlpm.x,Rate,_xlpm.x*2)")
    assertEquals(stored("LET(x,1,Total)"), "_xlfn.LET(_xlpm.x,1,Total)")
    // a LET name shadowing a future-function name is a name, and the call is still a call
    assertEquals(
      stored("LET(IFS,1,IFS+IFS(1,2))"),
      "_xlfn.LET(_xlpm.IFS,1,_xlpm.IFS+_xlfn.IFS(1,2))"
    )
  }

  test("toStored: commas inside strings and array constants do not shift LET positions") {
    assertEquals(
      stored("LET(x,\"a,b\",x)"),
      "_xlfn.LET(_xlpm.x,\"a,b\",_xlpm.x)"
    )
    assertEquals(
      stored("LET(x,{1,2,3},SUM(x))"),
      "_xlfn.LET(_xlpm.x,{1,2,3},SUM(_xlpm.x))"
    )
  }

  test("toStored: a parameter name never swallows a reference or an error literal spelled alike") {
    // `A:A` and `$A$1` are references even beside a parameter named A; #N/A is not the name N
    assertEquals(
      stored("LET(A,1,SUM(A:A)+$A$1+A)"),
      "_xlfn.LET(_xlpm.A,1,SUM(A:A)+$A$1+_xlpm.A)"
    )
    assertEquals(
      stored("LET(N,1,IFERROR(#N/A,N))"),
      "_xlfn.LET(_xlpm.N,1,IFERROR(#N/A,_xlpm.N))"
    )
  }

  test("toStored: LAMBDA parameters (optional ones in brackets) take _xlpm.; the body is not one") {
    assertEquals(
      stored("LAMBDA(a,b,a+b)(1,2)"),
      "_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)(1,2)"
    )
    assertEquals(
      stored("BYROW(A1:B2,LAMBDA(r,SUM(r)))"),
      "_xlfn.BYROW(A1:B2,_xlfn.LAMBDA(_xlpm.r,SUM(_xlpm.r)))"
    )
    assertEquals(
      stored("LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))"),
      "_xlfn.LAMBDA(_xlpm.x,[_xlpm.y],IF(_xlfn.ISOMITTED(_xlpm.y),_xlpm.x,_xlpm.x+_xlpm.y))"
    )
    // a body that is a bare defined name is the calculation, not a parameter
    assertEquals(stored("LAMBDA(x,Total)"), "_xlfn.LAMBDA(_xlpm.x,Total)")
  }

  test("toStored: _xlpm. is never doubled; an openpyxl-style bare-parameter LET is completed") {
    val excel = "_xlfn.LET(_xlpm.x,1,_xlpm.x+1)"
    assertEquals(stored(excel), excel)
    assertEquals(stored("_xlfn.LET(x,1,x+1)"), excel)
    assertEquals(stored("LET(_XLPM.x,1,x+1)"), "_xlfn.LET(_XLPM.x,1,_xlpm.x+1)")
  }

  test("fromStored: strips _xlpm. from LET/LAMBDA parameters, keeps it elsewhere") {
    assertEquals(restored("_xlfn.LET(_xlpm.x,1,_xlpm.x+1)"), "LET(x,1,x+1)")
    assertEquals(
      restored(
        "_xlfn.LAMBDA(_xlpm.x,[_xlpm.y],IF(_xlfn.ISOMITTED(_xlpm.y),_xlpm.x,_xlpm.x+_xlpm.y))"
      ),
      "LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))"
    )
    // outside a recognized parameter position the writer would not restore it, so it stays
    assertEquals(restored("SUM(_xlpm.x)"), "SUM(_xlpm.x)")
    assertEquals(stored("SUM(_xlpm.x)"), "SUM(_xlpm.x)")
  }

  test("fromStored: strips the prefixes of known future functions") {
    assertEquals(
      restored("_xlfn.MAXIFS(B1:B3,A1:A3,\"a\")"),
      "MAXIFS(B1:B3,A1:A3,\"a\")"
    )
    assertEquals(
      restored("_xlfn._xlws.FILTER(A1:A3,B1:B3)"),
      "FILTER(A1:A3,B1:B3)"
    )
    assertEquals(restored("_XLFN.XLOOKUP(A1,B:B,C:C)"), "XLOOKUP(A1,B:B,C:C)")
    assertEquals(
      restored("IFERROR(_xlfn.XLOOKUP(A1,B:B,C:C),_xlfn.MAXIFS(C:C,B:B,A1))"),
      "IFERROR(XLOOKUP(A1,B:B,C:C),MAXIFS(C:C,B:B,A1))"
    )
  }

  test("fromStored: a prefix on an unknown function is kept so the writer cannot lose it") {
    assertEquals(restored("_xlfn.NOSUCHFN(A1)"), "_xlfn.NOSUCHFN(A1)")
    assertEquals(
      stored(restored("_xlfn.NOSUCHFN(A1)")),
      "_xlfn.NOSUCHFN(A1)"
    )
  }

  test("fast paths: no call and no _xl prefix return the input itself, unscanned") {
    val noCall = "A1*2+Sheet2!B3"
    assert(stored(noCall) eq noCall)
    val bare = "IF(SUM(A1:A3)>0,VLOOKUP(A1,B1:C3,2,FALSE),0)"
    assert(restored(bare) eq bare)
    // the stem check is case-insensitive and position-independent
    assertEquals(restored("1+_XLFN.IFS(A1>0,1,TRUE,0)"), "1+IFS(A1>0,1,TRUE,0)")
  }

  test("fromStored ∘ toStored = id on bare model text; toStored ∘ fromStored = id on file text") {
    val model = List(
      "MAXIFS(B1:B3,A1:A3,\"a\")",
      "IFERROR(XLOOKUP(A1,B:B,C:C),0)",
      "FILTER(A1:A3,B1:B3)",
      "SUM(A1:A3)",
      "LET(x,1,x+1)",
      "LET(x, SUM(A1:A3), y, x*2, IF(x>y, MAXIFS(B:B,C:C,x), \"a,b\"))",
      "LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))",
      "BYROW(A1:B2,LAMBDA(r,SUM(r)))",
      "SUM(Table1['[Total])+MAXIFS(A1:A3,B1:B3,1)"
    )
    model.foreach(m => assertEquals(restored(stored(m)), m, m))
    val file = model.map(stored) ++ List("_xlfn.NOSUCHFN(A1)", "SUM(_xlpm.x)")
    file.foreach(f => assertEquals(stored(restored(f)), f, f))
    file.foreach(f => assertEquals(stored(f), f, f))
  }

  test("bareFutureCalls: names every call toStored would prefix, and nothing else") {
    assertEquals(bareFutureCalls("IFS(A1=1,1,TRUE,0)"), Vector("IFS"))
    assertEquals(bareFutureCalls("=IFS(A1=1,1,TRUE,0)"), Vector("IFS"))
    assertEquals(bareFutureCalls("_xlfn.IFS(A1=1,1,TRUE,0)"), Vector.empty)
    assertEquals(bareFutureCalls("_XLFN.ifs(A1)"), Vector.empty)
    assertEquals(bareFutureCalls("_xlfn._xlws.FILTER(A1:A3,B1:B3)"), Vector.empty)
    // `_xlws.` alone is not a prefix Excel resolves — the `_xlfn.` in front of it is missing
    assertEquals(bareFutureCalls("_xlws.FILTER(A1:A3,B1:B3)"), Vector("FILTER"))
    assertEquals(bareFutureCalls("xlookup(A1,A1:A3,B1:B3)"), Vector("XLOOKUP"))
    assertEquals(bareFutureCalls("SUMIFS(A1:A3,B1:B3,1)"), Vector.empty)
    assertEquals(bareFutureCalls("\"IFS(\"&SUM(A1)"), Vector.empty)
    assertEquals(bareFutureCalls("'IFS('!A1+'it''s IFS('!B1"), Vector.empty)
    assertEquals(bareFutureCalls("SUM(Table1[IFS(])"), Vector.empty)
    assertEquals(bareFutureCalls("XIFS(1)+MY.IFS(1)+IFS2(1)"), Vector.empty)
    assertEquals(
      bareFutureCalls("IFS(1,XLOOKUP(1,A:A,B:B),IFS(2,CONCAT(1)))"),
      Vector("IFS", "XLOOKUP", "CONCAT")
    )
    // distinct, in order of first appearance
    assertEquals(bareFutureCalls("IFS(1,IFS(2,CONCAT(1)),CONCAT(2))"), Vector("IFS", "CONCAT"))
    assertEquals(bareFutureCalls(""), Vector.empty)
    assertEquals(bareFutureCalls("A1*2"), Vector.empty)
    assertEquals(bareFutureCalls("IFS"), Vector.empty) // a name that is not a call
    // structured-reference quote escapes keep the bracket depth honest — a hand-rolled scanner
    // that treats the escaped '[' as real swallows the rest of the formula and misses the call
    assertEquals(bareFutureCalls("SUM(Table1['[Total])+IFS(1,1)"), Vector("IFS"))
    // array constants and error literals are opaque, but do not hide a call beside them
    assertEquals(bareFutureCalls("SUM({1,2})+#DIV/0!"), Vector.empty)
    assertEquals(bareFutureCalls("IFS({1,2},#N/A)"), Vector("IFS"))
  }

  test("bareFutureCalls: a LET / LAMBDA whose parameters lack _xlpm. is reported under its name") {
    assertEquals(bareFutureCalls("LET(x,1,x+1)"), Vector("LET"))
    // the openpyxl class: the function is prefixed, its parameters are not — Excel reports the
    // book as unreadable content on open, a harder failure than a bare function's #NAME?
    assertEquals(bareFutureCalls("_xlfn.LET(x,1,x+1)"), Vector("LET"))
    assertEquals(bareFutureCalls("_xlfn.LET(_xlpm.x,1,x+1)"), Vector("LET"))
    assertEquals(bareFutureCalls("_xlfn.LET(_xlpm.x,1,_xlpm.x+1)"), Vector.empty)
    assertEquals(bareFutureCalls("_xlfn.LAMBDA(x,[y],x+y)"), Vector("LAMBDA"))
    assertEquals(bareFutureCalls("_xlfn.LAMBDA(_xlpm.x,[_xlpm.y],_xlpm.x+_xlpm.y)"), Vector.empty)
    assertEquals(bareFutureCalls("_xlfn.BYROW(A1:B2,_xlfn.LAMBDA(r,SUM(r)))"), Vector("LAMBDA"))
    // a bare reference resolves to the innermost scope declaring it
    assertEquals(bareFutureCalls("_xlfn.LAMBDA(_xlpm.a,_xlfn.LET(b,1,_xlpm.a+b))"), Vector("LET"))
    assertEquals(
      bareFutureCalls("_xlfn.LAMBDA(a,_xlfn.LET(_xlpm.b,1,a+_xlpm.b))"),
      Vector("LAMBDA")
    )
    // values and the calculation that are defined names are not parameters
    assertEquals(bareFutureCalls("_xlfn.LET(_xlpm.x,Rate,_xlpm.x*Rate)"), Vector.empty)
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
    assert(FormulaStorage.ParameterScoped.subsetOf(FormulaStorage.FutureFunctions))
    // Excel 2007 functions must NOT be prefixed
    List("SUMIFS", "COUNTIFS", "AVERAGEIFS", "IFERROR", "SUM", "VLOOKUP").foreach { n =>
      assert(!FormulaStorage.FutureFunctions.contains(n), n)
    }
  }
