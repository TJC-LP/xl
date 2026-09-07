package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.formula.printer.FormulaOps
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

  test("shift keeps today's clamp at the sheet edge") {
    assertEquals(FormulaOps.shift("=A1", -5, -5), Right("=A1"): XLResult[String])
    assertEquals(FormulaOps.shift("=B2", -1, -5), Right("=A1"): XLResult[String])
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
