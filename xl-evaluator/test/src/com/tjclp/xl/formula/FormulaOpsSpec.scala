package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.formula.printer.{CutReferences, FormulaOps}
import munit.FunSuite

/**
 * ADR-017 §2.9: `FormulaOps` is the string-in/string-out face of the formula engine — parse,
 * rewrite, reprint — that the CLI's parse→shift→print copies collapse onto. It keeps the caller's
 * leading-'=' convention, prints in Excel's file form (bare ',' separators), and never touches text
 * that does not need rewriting.
 */
class FormulaOpsSpec extends FunSuite:

  private val Sheet1 = SheetName.unsafe("Sheet1")
  private val Data = SheetName.unsafe("Data")
  private val Q1Data = SheetName.unsafe("Q1 Data")

  // ===== shift =====

  test("shift drags relative references and keeps absolute ones") {
    assertEquals(FormulaOps.shift("=A1+$B$1", 1, 1), Right("=B2+$B$1"): XLResult[String])
    assertEquals(
      FormulaOps.shift("=SUM(A1:B2)*$C1", 0, 2),
      Right("=SUM(A3:B4)*$C3"): XLResult[String]
    )
    assertEquals(FormulaOps.shift("A1", 0, 1), Right("A2"): XLResult[String])
  }

  test("GH-612: shift never clamps — a reference pushed before A1 becomes #REF!") {
    assertEquals(FormulaOps.shift("=A1", -5, -5), Right("=#REF!"): XLResult[String])
    assertEquals(FormulaOps.shift("=B2", -1, -5), Right("=#REF!"): XLResult[String])
  }

  test("shift(text, 0, 0) canonicalises but never changes meaning") {
    assertEquals(FormulaOps.shift("=SUM(A1, B1)", 0, 0), Right("=SUM(A1,B1)"): XLResult[String])
  }

  test("shift refuses unparsable text with a FormulaError naming the text") {
    FormulaOps.shift("=A1+", 1, 1) match
      case Left(XLError.FormulaError(formula, _)) => assertEquals(formula, "=A1+")
      case other => fail(s"expected FormulaError, got $other")
  }

  // ===== mentionsSheet =====

  test("mentionsSheet recognises bare, quoted and case-varied qualifiers") {
    assert(FormulaOps.mentionsSheet("Sheet1!A1", Sheet1))
    assert(FormulaOps.mentionsSheet("='Sheet1'!A1", Sheet1))
    assert(FormulaOps.mentionsSheet("sheet1!A1", Sheet1))
    assert(FormulaOps.mentionsSheet("SUM(Sheet1 !A1:A9)", Sheet1))
    assert(FormulaOps.mentionsSheet("Sheet1!rate", Sheet1))
    assert(FormulaOps.mentionsSheet("'Q1 Data'!A1", Q1Data))
    assert(FormulaOps.mentionsSheet("'It''s'!A1", SheetName.unsafe("It's")))
  }

  test("mentionsSheet ignores string literals") {
    assert(!FormulaOps.mentionsSheet("\"Sheet1!A1\"", Sheet1))
    assert(!FormulaOps.mentionsSheet("IF(A1=\"Sheet1!A1\",1,0)", Sheet1))
    assert(FormulaOps.mentionsSheet("IF(A1=\"x\",Sheet1!A1,0)", Sheet1))
  }

  test("GH-653: a drag keeps `NOT (A1)^2` the call — reprinted NOT(A2)^2, never NOT(A2^2)") {
    assertEquals(FormulaOps.shift("=NOT (A1)^2", 0, 1), Right("=NOT(A2)^2"): XLResult[String])
    assertEquals(FormulaOps.shift("=NOT (A1)%", 0, 1), Right("=NOT(A2)%"): XLResult[String])
    assertEquals(FormulaOps.shift("=NOT(A1)^2", 0, 1), Right("=NOT(A2)^2"): XLResult[String])
  }

  test("GH-653: a double quote inside a quoted sheet name is not a string-literal boundary") {
    val final1 = SheetName.unsafe("Q1 \"Final\"")
    assert(FormulaOps.mentionsSheet("='Q1 \"Final\"'!A1", final1))
    assert(FormulaOps.mentionsSheet("=SUM('Q1 \"Final\"'!A1:A9)+\"x\"", final1))
    assert(FormulaOps.mentionsSheet("=\"it's\"&Sheet1!A1", Sheet1))
    assert(!FormulaOps.mentionsSheet("=\"'Q1 \"\"Final\"\"'!A1\"", final1))
    // the rename that used to leave the reference stale
    assertEquals(
      FormulaOps.renameSheet("='Q1 \"Final\"'!A1", final1, SheetName.unsafe("Final")),
      Right("=Final!A1"): XLResult[String]
    )
    // and the 3-D span that used to slip past the gate
    assert(FormulaOps.renameSheet("='Q1 \"Final\":Q3'!A1", final1, Sheet1).isLeft)
  }

  test("mentionsSheet is not fooled by longer names, external refs or defined names") {
    assert(!FormulaOps.mentionsSheet("MySheet1!A1", Sheet1))
    assert(!FormulaOps.mentionsSheet("Sheet10!A1", Sheet1))
    assert(!FormulaOps.mentionsSheet("[2]Sheet1!A1", Sheet1))
    assert(!FormulaOps.mentionsSheet("'[2]Sheet1'!A1", Sheet1))
    assert(!FormulaOps.mentionsSheet("Sheet1", Sheet1))
    assert(!FormulaOps.mentionsSheet("A1*2", Sheet1))
  }

  // ===== renameSheet =====

  test("renameSheet rewrites the qualifier and keeps the caller's leading-'=' convention") {
    assertEquals(
      FormulaOps.renameSheet("=Sheet1!A1*2", Sheet1, Data),
      Right("=Data!A1*2"): XLResult[String]
    )
    assertEquals(
      FormulaOps.renameSheet("Sheet1!A1*2", Sheet1, Data),
      Right("Data!A1*2"): XLResult[String]
    )
    assertEquals(
      FormulaOps.renameSheet("SUM(Sheet1!$A$1:A9)", Sheet1, Data),
      Right("SUM(Data!$A$1:A9)"): XLResult[String]
    )
  }

  test("renameSheet quotes a new name that needs it and unquotes one that no longer does") {
    assertEquals(
      FormulaOps.renameSheet("Sheet1!A1", Sheet1, Q1Data),
      Right("'Q1 Data'!A1"): XLResult[String]
    )
    assertEquals(
      FormulaOps.renameSheet("'Q1 Data'!A1", Q1Data, Sheet1),
      Right("Sheet1!A1"): XLResult[String]
    )
  }

  test("renameSheet leaves text that does not reference the sheet byte-identical") {
    assertEquals(
      FormulaOps.renameSheet("SUM(A1, B1)", Sheet1, Data),
      Right("SUM(A1, B1)"): XLResult[String]
    )
    assertEquals(
      FormulaOps.renameSheet("\"Sheet1\"", Sheet1, Data),
      Right("\"Sheet1\""): XLResult[String]
    )
    assertEquals(
      FormulaOps.renameSheet("[2]Sheet1!A1", Sheet1, Data),
      Right("[2]Sheet1!A1"): XLResult[String]
    )
    // unparsable text that never mentions the sheet is not this operation's problem
    assertEquals(FormulaOps.renameSheet("A1+", Sheet1, Data), Right("A1+"): XLResult[String])
  }

  test("renameSheet to the same name is the identity") {
    assertEquals(
      FormulaOps.renameSheet("SUM(Sheet1!A1, B1)", Sheet1, Sheet1),
      Right("SUM(Sheet1!A1, B1)"): XLResult[String]
    )
  }

  test("renameSheet refuses unparsable text that mentions the sheet") {
    FormulaOps.renameSheet("Sheet1!A1+", Sheet1, Data) match
      case Left(XLError.FormulaError(formula, reason)) =>
        assertEquals(formula, "Sheet1!A1+")
        assert(reason.contains("Sheet1"), reason)
      case other => fail(s"expected FormulaError, got $other")
  }

  test("refusals carry the parser's diagnostic, never its constructor text (GH-608)") {
    FormulaOps.renameSheet("IF(ZZZNOTAFUNC(1)=1,Sheet1!A1,0)", Sheet1, Data) match
      case Left(XLError.FormulaError(formula, reason)) =>
        assertEquals(formula, "IF(ZZZNOTAFUNC(1)=1,Sheet1!A1,0)")
        assert(reason.contains("Unknown function 'ZZZNOTAFUNC' at position 3"), reason)
        assert(!reason.contains("UnknownFunction("), reason)
        assert(!reason.contains("List("), reason)
      case other => fail(s"expected FormulaError, got $other")
    FormulaOps.shift("=ZZZNOTAFUNC(A1)", 1, 1) match
      case Left(XLError.FormulaError(_, reason)) =>
        assert(reason.contains("Unknown function 'ZZZNOTAFUNC'"), reason)
        assert(!reason.contains("UnknownFunction("), reason)
      case other => fail(s"expected FormulaError, got $other")
  }

  // ===== 3-D ranges: both ends are mentions, so both ends refuse (the parser has no 3-D form) =====

  private val Sheet3 = SheetName.unsafe("Sheet3")

  test("mentionsSheet sees both ends of a 3-D range, bare or quoted") {
    assert(FormulaOps.mentionsSheet("SUM(Sheet1:Sheet3!A1)", Sheet1))
    assert(FormulaOps.mentionsSheet("SUM(Sheet1:Sheet3!A1)", Sheet3))
    assert(FormulaOps.mentionsSheet("SUM(Sheet1 : Sheet3 !A1)", Sheet1))
    assert(!FormulaOps.mentionsSheet("SUM(Sheet1:Sheet3!A1)", SheetName.unsafe("Sheet2")))
    assert(!FormulaOps.mentionsSheet("SUM(Sheet10:Sheet13!A1)", Sheet1))
    assert(!FormulaOps.mentionsSheet("SUM(Sheet10:Sheet13!A1)", Sheet3))
    assert(FormulaOps.mentionsSheet("SUM('Q1 Data':'Q3 Data'!A1)", Q1Data))
    assert(FormulaOps.mentionsSheet("SUM('Q1 Data':'Q3 Data'!A1)", SheetName.unsafe("Q3 Data")))
    assert(FormulaOps.mentionsSheet("SUM('Sheet1:Sheet 3'!A1)", Sheet1))
    assert(FormulaOps.mentionsSheet("SUM('Sheet1:Sheet 3'!A1)", SheetName.unsafe("Sheet 3")))
    assert(FormulaOps.mentionsSheet("SUM('Sheet 1:Sheet3'!A1)", SheetName.unsafe("Sheet 1")))
    assert(FormulaOps.mentionsSheet("SUM('Sheet 1:Sheet3'!A1)", Sheet3))
    // a range between a defined NAME and a cell is not a 3-D range
    assert(!FormulaOps.mentionsSheet("SUM(Sheet1:A5)", Sheet1))
    assert(!FormulaOps.mentionsSheet("\"Sheet1:Sheet3!A1\"", Sheet1))
  }

  test("renameSheet refuses a 3-D range for EITHER end rather than leaving one end stale") {
    for sheet <- Vector(Sheet1, Sheet3) do
      FormulaOps.renameSheet("SUM(Sheet1:Sheet3!A1)", sheet, Data) match
        case Left(XLError.FormulaError(formula, _)) =>
          assertEquals(formula, "SUM(Sheet1:Sheet3!A1)")
        case other => fail(s"renaming ${sheet.value}: expected FormulaError, got $other")
  }

  // ===== GH-612: full-column / full-row forms and off-grid references =====

  private def shifted(text: String, colDelta: Int, rowDelta: Int): String =
    FormulaOps.shift(text, colDelta, rowDelta).fold(e => fail(e.message), identity)

  test("GH-612: a full-column reference drags along columns only") {
    assertEquals(shifted("=COUNTIF($A:$A,B1)", 0, 1), "=COUNTIF($A:$A,B2)")
    assertEquals(shifted("=COUNTIF($A:$A,B1)", 0, 2), "=COUNTIF($A:$A,B3)")
    assertEquals(shifted("=SUM(E:E)", 1, 0), "=SUM(F:F)")
    assertEquals(shifted("=SUM(E:E)", 0, 5), "=SUM(E:E)")
    assertEquals(shifted("=SUM(E:E)", 0, 1048575), "=SUM(E:E)")
    assertEquals(shifted("=SUM($E:$E)", 3, 3), "=SUM($E:$E)")
    assertEquals(shifted("=SUM($A:C)", 1, 1), "=SUM($A:D)")
    assertEquals(shifted("=SUM(Sheet1!A:A)", 2, 9), "=SUM(Sheet1!C:C)")
    assertEquals(shifted("=SUM([2]Book1!A:A)", 1, 4), "=SUM([2]Book1!B:B)")
    assertEquals(shifted("=[2]Book1!A:A", 1, 0), "=[2]Book1!B:B")
    assertEquals(shifted("=A:A", 1, 1), "=B:B")
  }

  test("GH-612: a full-row reference drags along rows only") {
    assertEquals(shifted("=SUM(1:1)", 1, 0), "=SUM(1:1)")
    assertEquals(shifted("=SUM(1:1)", 16383, 0), "=SUM(1:1)")
    assertEquals(shifted("=SUM(1:1)", 0, 1), "=SUM(2:2)")
    assertEquals(shifted("=SUM($3:$10)", 4, 4), "=SUM($3:$10)")
    assertEquals(shifted("=SUM($3:10)", 0, 1), "=SUM($3:11)")
    assertEquals(shifted("=1:1", 1, 1), "=2:2")
  }

  test("GH-612: a zero shift keeps the whole-column form (the first dragged cell)") {
    assertEquals(shifted("=COUNTIF($A:$A,B1)", 0, 0), "=COUNTIF($A:$A,B1)")
    assertEquals(shifted("=SUM(1:1)", 0, 0), "=SUM(1:1)")
  }

  test("GH-612: the mixed formula dragged down two and right one") {
    assertEquals(
      shifted("=COUNTIF($A:$A,B1)+SUM(1:1)+A1", 1, 2),
      "=COUNTIF($A:$A,C3)+SUM(3:3)+B3"
    )
  }

  test("GH-612: a corner range over every row is the whole-column form and drags like one") {
    // Excel canonicalises A1:A1048576 to A:A at entry; the corner spelling only reaches xl from
    // other producers or a putf, and drags exactly as A:A would
    assertEquals(shifted("=SUM(A1:A1048576)", 1, 0), "=SUM(B:B)")
    assertEquals(shifted("=SUM(A1:A1048576)", 0, 1), "=SUM(A:A)")
    assertEquals(shifted("=SUM(A1:XFD1)", 0, 2), "=SUM(3:3)")
    assertEquals(shifted("=SUM(A2:A1048576)", 0, 1), "=SUM(#REF!)")
  }

  test("GH-612: a shift that makes a relative corner overtake an anchored one prints normalised") {
    assertEquals(shifted("=SUM(E:$E)", 1, 0), "=SUM($E:F)")
    assertEquals(shifted("=SUM($E:E)", -1, 0), "=SUM(D:$E)")
    assertEquals(shifted("=SUM(3:$3)", 0, 2), "=SUM($3:5)")
    assertEquals(shifted("=SUM(A1:$B$1)", 2, 0), "=SUM($B$1:C1)")
  }

  test("GH-612: a reference pushed off the grid becomes #REF!, never a non-existent address") {
    // before row 1 / column A
    assertEquals(shifted("=A1+B3", 0, -1), "=#REF!+B2")
    assertEquals(shifted("=Sheet1!A1", 0, -1), "=#REF!")
    // a range in a range-typed slot becomes #REF! inside the call, as Excel writes it
    assertEquals(shifted("=SUM(A1:A2)+C5", 0, -1), "=SUM(#REF!)+C4")
    assertEquals(shifted("=COUNTIF(A1:A2,B5)", 0, -1), "=COUNTIF(#REF!,B4)")
    assertEquals(shifted("=SUM(A:A)", -1, 0), "=SUM(#REF!)")
    assertEquals(shifted("=SUM(1:1)", 0, -1), "=SUM(#REF!)")
    assertEquals(shifted("=SUM(Sheet1!A:A)", -1, 0), "=SUM(#REF!)")
    assertEquals(shifted("=A:A", -1, 0), "=#REF!")
    // past row 1048576 / column XFD
    assertEquals(shifted("=A1048576", 0, 1), "=#REF!")
    assertEquals(shifted("=XFD1", 1, 0), "=#REF!")
    assertEquals(shifted("=SUM(XFC1:XFD1)", 1, 0), "=SUM(#REF!)")
    assertEquals(shifted("=SUM(XFD:XFD)", 1, 0), "=SUM(#REF!)")
    assertEquals(shifted("=SUM(1048576:1048576)", 0, 1), "=SUM(#REF!)")
    // the written #REF! re-parses and shifts as an error, staying put
    assertEquals(shifted("=SUM(#REF!)+C4", 1, 1), "=SUM(#REF!)+D5")
    // anchors pin an axis, so it cannot fall off
    assertEquals(shifted("=A$1", 0, -1), "=A$1")
    assertEquals(shifted("=$A1", -1, 0), "=$A1")
    assertEquals(shifted("=$A:$A", -1, 0), "=$A:$A")
    assertEquals(shifted("=$1:$1", 0, -1), "=$1:$1")
    // the #REF! is per reference: the rest of the formula survives and re-parses
    assertEquals(shifted("=IF(A1>0,ABS(A1),B5)", 0, -1), "=IF(#REF!>0,ABS(#REF!),B4)")
    assertEquals(shifted("=#REF!+B2", 1, 1), "=#REF!+C3")
  }

  // ===== canonicalSheets (GH-936) =====

  private val tabs = Vector(Sheet1, Data, Q1Data, SheetName.unsafe("O'Brien"))
  private def canon(text: String): String = FormulaOps.canonicalSheets(text, tabs)

  test("GH-936: a qualifier typed in another case is stored as the sheet's own name") {
    assertEquals(canon("=sheet1!A1"), "=Sheet1!A1")
    assertEquals(canon("=SHEET1!A1+1"), "=Sheet1!A1+1")
    assertEquals(canon("=SUM(data!B1:B3)"), "=SUM(Data!B1:B3)")
    assertEquals(canon("=SUM(DATA!A:A)"), "=SUM(Data!A:A)")
    // the bare (equals-free) cell text keeps its form
    assertEquals(canon("sheet1!A1"), "Sheet1!A1")
  }

  test("GH-936: quoted names stay quoted, apostrophes doubled") {
    assertEquals(canon("='q1 data'!A1"), "='Q1 Data'!A1")
    assertEquals(canon("='o''brien'!B2"), "='O''Brien'!B2")
  }

  test("GH-936: sheet-qualified defined names follow too") {
    assertEquals(canon("=data!rate*2"), "=Data!rate*2")
    assertEquals(canon("=SUMIF(data!crit,\">1\")"), "=SUMIF(Data!crit,\">1\")")
  }

  test("GH-936: unknown sheets, external refs and string literals stay as typed") {
    assertEquals(canon("=nosuch!A1"), "=nosuch!A1")
    assertEquals(canon("=[1]sheet1!A1"), "=[1]sheet1!A1")
    assertEquals(canon("=INDIRECT(\"sheet1!A1\")"), "=INDIRECT(\"sheet1!A1\")")
    // a mix: only the known qualifier changes
    assertEquals(canon("=sheet1!A1+nosuch!B1"), "=Sheet1!A1+nosuch!B1")
  }

  test("GH-936: canonical or unqualified text comes back byte-identical, spacing and case kept") {
    assertEquals(canon("=sum( Sheet1!A1 , 2 )"), "=sum( Sheet1!A1 , 2 )")
    assertEquals(canon("=sum( a1 , 2 )"), "=sum( a1 , 2 )")
    assertEquals(canon("=SUM(("), "=SUM((")
    assertEquals(canon("=sheet1!"), "=sheet1!")
  }

  test("GH-936 review: only the qualifier changes; every other byte is kept as typed") {
    assertEquals(canon("=(data!A1) + 1"), "=(Data!A1) + 1")
    assertEquals(canon("=-(-data!A1)"), "=-(-Data!A1)")
    assertEquals(canon("=1E3+data!A1"), "=1E3+Data!A1")
    assertEquals(canon("=sum( data!A1 ,\"x!y\")"), "=sum( Data!A1 ,\"x!y\")")
    assertEquals(canon("=data!$A$1:$B2 + 'q1 data'!C3"), "=Data!$A$1:$B2 + 'Q1 Data'!C3")
    assertEquals(canon("=data!#REF!+1"), "=Data!#REF!+1")
    // a 3-D reference is one opaque span, kept as typed; the other qualifiers are still respelled
    assertEquals(canon("=SUM(sheet1:data!A1)"), "=SUM(sheet1:data!A1)")
    assertEquals(canon("=SUM(sheet1:data!A1)+data!B2"), "=SUM(sheet1:data!A1)+Data!B2")
    assertEquals(
      canon("=SUM('q1 data':data!A1:B2)+sheet1!C3"),
      "=SUM('q1 data':data!A1:B2)+Sheet1!C3"
    )
  }

  test("GH-935 review: relocateCell rewrites only the reference, keeping every other byte") {
    def rel(t: String) = FormulaOps.relocateCell(
      t,
      CutReferences.Host.Cell(Sheet1, ref"Z9", array = false),
      Sheet1,
      ref"B1",
      ref"A1"
    )
    assertEquals(rel("=( B1 ) + 1E3"), "=( A1 ) + 1E3")
    assertEquals(rel("=-(-$B$1)"), "=-(-$A$1)")
    assertEquals(rel("=sum(b1:b1, B10, Sheet1!B1)"), "=sum(A1:A1, B10, Sheet1!A1)")
    assertEquals(rel("=SUM(B1:B3)"), "=SUM(B1:B3)")
    assertEquals(rel("=\"B1\"&B1"), "=\"B1\"&A1")
  }

  test("GH-936 review 2: a structured reference does not hide the next qualifier") {
    assertEquals(canon("=SUM(Table1[Col] data!A1)"), "=SUM(Table1[Col] Data!A1)")
    assertEquals(canon("=SUM(A1:A2 data!A1)"), "=SUM(A1:A2 Data!A1)")
    // an external prefix still guards its own qualifier
    assertEquals(canon("=[1]sheet1!A1+data!A1"), "=[1]sheet1!A1+Data!A1")
  }

  test("GH-935 review 2: relocateCell reaches intersection operands") {
    def rel(t: String) = FormulaOps.relocateCell(
      t,
      CutReferences.Host.Cell(Sheet1, ref"Z9", array = false),
      Sheet1,
      ref"B1",
      ref"A1"
    )
    // an intersection depends on every cell of every operand area (Excel's conservative rule,
    // checked in Excel for Mac), wherever the formula sits (here Z9, off the cut cell's row and
    // column): even an empty one such as B1 C1 is rewritten
    assertEquals(rel("=B1 C1"), "=A1 C1")
    assertEquals(rel("=SUM(B1:C1 C1:C3)"), "=SUM(A1:C1 C1:C3)")
    assertEquals(rel("=B1 B1:C1"), "=A1 A1:C1")
    // a union depends on every cell of every area, the cut cell only in its second area
    assertEquals(rel("=SUM((D5,B1:C1))"), "=SUM((D5,A1:C1))")
    // an intersection without the cut cell is byte-identical
    assertEquals(rel("=SUM(C1:D1  C1:C3)"), "=SUM(C1:D1  C1:C3)")
    // a one-row range starting at the cut cell extends to the paste (Excel oracle rule c)
    assertEquals(rel("=SUM(B1 B1:C1)"), "=SUM(A1 A1:C1)")
    // the rewrite is textual, so a function xl does not register follows the cut too
    assertEquals(rel("=ISREF(B1)"), "=ISREF(A1)")
  }

  test(
    "GH-935 review 4: a 3-D reference stays as typed while the other references follow the cut"
  ) {
    def rel(t: String, host: SheetName) = FormulaOps.relocateCell(
      t,
      CutReferences.Host.Cell(host, ref"Z9", array = false),
      Sheet1,
      ref"B1",
      ref"A1"
    )
    val sheet3 = SheetName.unsafe("Sheet3")
    // checked in Excel for Mac (fixture f10-01)
    assertEquals(
      rel("=SUM(Sheet1:Sheet2!B1)+Sheet1!B1", sheet3),
      "=SUM(Sheet1:Sheet2!B1)+Sheet1!A1"
    )
    // not checked in Excel: on the merge sheet, the 3-D span's own B1 never leaks as a local ref
    assertEquals(rel("=SUM(Sheet1:Sheet2!B1)+B1", Sheet1), "=SUM(Sheet1:Sheet2!B1)+A1")
    assertEquals(rel("=SUM('Sheet1:Sheet 3'!B1)+B1", Sheet1), "=SUM('Sheet1:Sheet 3'!B1)+A1")
    assertEquals(rel("=SUM(Sheet1:'Sheet 3'!B1:C1)+B1", Sheet1), "=SUM(Sheet1:'Sheet 3'!B1:C1)+A1")
  }

  test("GH-935 review 4: array-typed parameters depend on every cell of the ranges they hold") {
    def rel(t: String) = FormulaOps.relocateCell(
      t,
      CutReferences.Host.Cell(Sheet1, ref"Z9", array = false),
      Sheet1,
      ref"B1",
      ref"A1"
    )
    // checked in Excel for Mac (fixture f9-03): SUMPRODUCT's, INDEX's first, ROWS' parameter
    assertEquals(rel("=SUMPRODUCT((B1:C1>0)*B1:C1)"), "=SUMPRODUCT((A1:C1>0)*A1:C1)")
    assertEquals(rel("=INDEX(B1:C1*1,1,1)"), "=INDEX(A1:C1*1,1,1)")
    assertEquals(rel("=ROWS(B1:C1*1)"), "=ROWS(A1:C1*1)")
    // value-typed parameters keep the implicit-intersection rule (Excel leaves these alone)
    assertEquals(rel("=SUM(B1:C1*1)"), "=SUM(B1:C1*1)")
    assertEquals(rel("=MAX(B1:C1*2)"), "=MAX(B1:C1*2)")
    // not checked in Excel: COLUMNS' parameter is array-typed too
    assertEquals(rel("=COLUMNS(B1:C1*1)"), "=COLUMNS(A1:C1*1)")
  }

  test("GH-936: tabs differing only by case keep the typed spelling (ambiguous)") {
    val twins = Vector(SheetName.unsafe("Data"), SheetName.unsafe("DATA"))
    assertEquals(FormulaOps.canonicalSheets("=data!A1", twins), "=data!A1")
    assertEquals(FormulaOps.canonicalSheets("=DATA!A1", twins), "=DATA!A1")
  }
