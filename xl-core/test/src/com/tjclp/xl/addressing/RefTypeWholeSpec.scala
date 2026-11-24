package com.tjclp.xl.addressing

import com.tjclp.xl.cells.{Cell, CellValue, Hyperlink}
import com.tjclp.xl.sheets.Sheet
import munit.FunSuite

/** #932: whole-column and whole-row refs, as the batch hyperlink op reads them. */
class RefTypeWholeSpec extends FunSuite:

  private def rng(s: String): CellRange = CellRange.parse(s).fold(e => fail(e), identity)

  test("parseWithWhole reads whole columns and rows, bare and qualified") {
    assertEquals(RefType.parseWithWhole("A:A"), Right(RefType.Range(rng("A:A"))))
    assertEquals(RefType.parseWithWhole("B:D"), Right(RefType.Range(rng("B:D"))))
    assertEquals(RefType.parseWithWhole("3:3"), Right(RefType.Range(rng("3:3"))))
    assertEquals(
      RefType.parseWithWhole("'Q1 Data'!$A:$A"),
      Right(RefType.QualifiedRange(SheetName.unsafe("Q1 Data"), rng("A:A")))
    )
    assertEquals(RefType.parseWithWhole("A1:B2"), RefType.parse("A1:B2"))
    assertEquals(RefType.parseWithWhole("C7"), RefType.parse("C7"))
    assert(RefType.parseWithWhole("A:3").isLeft)
    assert(RefType.parseWithWhole("A1:B").isLeft)
    assert(RefType.parseWithWhole("Bad[Name!A:A").isLeft)
  }

  test("a whole-column link is one entry and reports for every row without cells") {
    val sheet = Sheet("S")
      .put(Cell(ARef.from0(0, 0), CellValue.Text("Top")))
      .setHyperlink(rng("A:A"), Hyperlink("https://col/"))
    assertEquals(sheet.rangeHyperlinks.size, 1)
    assertEquals(sheet.cells.size, 1)
    assertEquals(sheet(ARef.from0(0, 1048575)).hyperlink, Some(Hyperlink("https://col/")))
  }
