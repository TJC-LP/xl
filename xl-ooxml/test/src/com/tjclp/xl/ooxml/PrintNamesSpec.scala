package com.tjclp.xl.ooxml

import com.tjclp.xl.addressing.{CellRange, SheetName}
import munit.FunSuite

/**
 * GH-263: sheet names that parse as cell references (Q1, A1, R1C1, ...) must be single-quoted in
 * defined-name formulas, or the generated workbook.xml is spec-invalid (Excel writes
 * `'Q1'!$A$1:$B$2`, never `Q1!$A$1:$B$2`). Quoting delegates to the shared
 * `SheetName.needsQuoting` predicate.
 */
class PrintNamesSpec extends FunSuite:

  private def range(s: String): CellRange =
    CellRange.parse(s).fold(e => fail(s"bad range: $e"), identity)

  test("GH-263: quoteSheetName quotes cell-ref-shaped names") {
    assertEquals(PrintNames.quoteSheetName("Q1"), "'Q1'")
    assertEquals(PrintNames.quoteSheetName("A1"), "'A1'")
    assertEquals(PrintNames.quoteSheetName("XFD1048576"), "'XFD1048576'")
    assertEquals(PrintNames.quoteSheetName("R1C1"), "'R1C1'")
    assertEquals(PrintNames.quoteSheetName("RC"), "'RC'")
  }

  test("GH-263: quoteSheetName keeps safe names bare") {
    assertEquals(PrintNames.quoteSheetName("Sheet1"), "Sheet1")
    assertEquals(PrintNames.quoteSheetName("Sales_2024"), "Sales_2024")
    assertEquals(PrintNames.quoteSheetName("XFE1"), "XFE1") // beyond max column: not a cell ref
  }

  test("GH-263: quoteSheetName quotes leading digits, spaces, and escapes quotes") {
    assertEquals(PrintNames.quoteSheetName("2024Q1"), "'2024Q1'")
    assertEquals(PrintNames.quoteSheetName("Q1 Report"), "'Q1 Report'")
    assertEquals(PrintNames.quoteSheetName("It's Q1"), "'It''s Q1'")
  }

  test("GH-263: print-area defined-name formula for sheet Q1 is quoted") {
    assertEquals(
      PrintNames.printAreaFormula(SheetName.unsafe("Q1"), range("A1:B2")),
      "'Q1'!$A$1:$B$2"
    )
  }

  test("GH-263: print-titles defined-name formula for sheet Q1 is quoted") {
    assertEquals(PrintNames.printTitlesFormula(SheetName.unsafe("Q1"), (1, 3)), "'Q1'!$1:$3")
  }

  test("GH-263: quoted print formulas parse back to the model (round-trip)") {
    val sheet = SheetName.unsafe("Q1")
    val area = range("A1:B2")
    val areaFormula = PrintNames.printAreaFormula(sheet, area)
    assertEquals(PrintNames.parsePrintArea(areaFormula, sheet), Some(area))
    val titlesFormula = PrintNames.printTitlesFormula(sheet, (1, 3))
    assertEquals(PrintNames.parsePrintTitles(titlesFormula, sheet), Some((1, 3)))
  }

end PrintNamesSpec
