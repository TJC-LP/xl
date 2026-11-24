package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{Cell, CellValue, Hyperlink}
import com.tjclp.xl.ops.ClearWhat
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/** #931 / #932: the range hyperlink model and the sheet operations over it. */
class RangeHyperlinkSpec extends ScalaCheckSuite:

  private def rng(s: String): CellRange = CellRange.parse(s).fold(e => fail(e), identity)
  private def at(s: String): ARef = ARef.parse(s).fold(e => fail(e), identity)

  private val link = Hyperlink("https://example.net/")
  private val other = Hyperlink("https://example.org/")
  private val d2e3 = RangeHyperlink(rng("D2:E3"), link, Some("Docs"), None)
  private val sheet = Sheet("S")
    .put(Cell(at("D2"), CellValue.Text("Range link")))
    .copy(rangeHyperlinks = com.tjclp.xl.sheets.RangeHyperlinks.from(Vector(d2e3)))

  // ===== geometry =====

  private val genSmallRange: Gen[CellRange] =
    for
      c1 <- Gen.choose(0, 7); c2 <- Gen.choose(0, 7)
      r1 <- Gen.choose(0, 7); r2 <- Gen.choose(0, 7)
    yield CellRange(ARef.from0(c1, r1), ARef.from0(c2, r2))

  private def cellSet(r: CellRange): Set[ARef] = r.cellsRowMajor.toSet

  property("remaining covers exactly the range minus the holes, in disjoint parts") {
    Prop.forAll(genSmallRange, Gen.listOfN(3, genSmallRange)) { (range, holes) =>
      val parts = RangeHyperlink.remaining(range, holes.toVector)
      val covered = parts.toList.flatMap(_.cellsRowMajor)
      assertEquals(covered.size, covered.distinct.size, "parts overlap")
      assertEquals(covered.toSet, cellSet(range) -- holes.flatMap(cellSet))
    }
  }

  test("remaining never visits the cells of a whole-column range") {
    val parts = RangeHyperlink.remaining(rng("A1:A1048576"), Vector(rng("A5:A5")))
    assertEquals(parts, Vector(rng("A1:A4"), rng("A6:A1048576")))
  }

  test("covers excludes the cut parts; without drops a link with nothing left") {
    val cut = d2e3.without(rng("E2:E2")).getOrElse(fail("D2:E3 minus E2 is not empty"))
    assert(cut.covers(at("D2")) && cut.covers(at("E3")))
    assert(!cut.covers(at("E2")))
    assert(cut.sameLinkAs(d2e3))
    assertEquals(d2e3.without(rng("A1:Z9")), None)
    assertEquals(cut.without(rng("D2:D3")).flatMap(_.without(rng("E3:E3"))), None)
    assertEquals(d2e3.without(rng("H8:H9")), Some(d2e3))
  }

  // ===== reads =====

  test("#931: apply and hyperlinkAt report the range link on covered cells only") {
    for a1 <- Seq("D2", "E2", "D3", "E3") do
      assertEquals(sheet(at(a1)).hyperlink, Some(link), a1)
      assertEquals(sheet.hyperlinkAt(at(a1)), Some(link), a1)
    assertEquals(sheet(at("F2")).hyperlink, None)
    assertEquals(sheet.rangeHyperlinkAt(at("E3")), Some(d2e3))
    assertEquals(sheet.cells.hyperlinkAt(at("D2")), None, "D2 must not own the range link")
    assert(!sheet.contains(at("E2")))
  }

  test("#931: a cell's own link wins over the range link at that cell") {
    val owned = sheet.put(Cell(at("E2")).withHyperlink(other))
    assertEquals(owned(at("E2")).hyperlink, Some(other))
    assertEquals(owned.hyperlinkAt(at("E2")), Some(other))
    assertEquals(owned(at("E3")).hyperlink, Some(link))
  }

  // ===== writes =====

  test("#931: putting back what apply reported keeps the link inherited") {
    val typed = sheet.put(sheet(at("E2")).withValue(CellValue.Text("x")))
    assertEquals(typed.cells.hyperlinkAt(at("E2")), None)
    assertEquals(typed(at("E2")).hyperlink, Some(link))
    assertEquals(typed.rangeHyperlinks.toVector, Vector(d2e3))
  }

  test("#931: styling covered cells, one by one or as a range, keeps the link inherited") {
    import com.tjclp.xl.sheets.styleSyntax.{withCellStyle, withRangeStyle}
    val bold = com.tjclp.xl.styles.CellStyle.default
      .withNumFmt(com.tjclp.xl.styles.numfmt.NumFmt.Percent)
    for styled <- Seq(sheet.withCellStyle(at("E2"), bold), sheet.withRangeStyle(rng("D2:E3"), bold))
    do
      for a1 <- Seq("D2", "E2", "D3", "E3") do
        assertEquals(styled.cells.hyperlinkAt(at(a1)), None, a1)
        assertEquals(styled(at(a1)).hyperlink, Some(link), a1)
      assertEquals(styled.rangeHyperlinks.toVector, Vector(d2e3))
  }

  test("#931: removeHyperlinks cuts the range and keeps the other cells linked") {
    val removed = sheet.removeHyperlinks(rng("E2:E2"))
    assertEquals(removed(at("E2")).hyperlink, None)
    assertEquals(removed(at("D2")).hyperlink, Some(link))
    assertEquals(removed.rangeHyperlinks.toVector.map(_.excluded), Vector(Vector(rng("E2:E2"))))
    assert(!removed.contains(at("E2")), "removal must not create cells")
    assertEquals(sheet.removeHyperlinks(rng("D2:E3")).rangeHyperlinks.toVector, Vector.empty)
  }

  test("#931: removeHyperlinks clears own links in the range too, keeping their cells") {
    val owned = sheet.put(Cell(at("B7"), CellValue.Text("x")).withHyperlink(other))
    val removed = owned.removeHyperlinks(rng("A1:E9"))
    assertEquals(removed(at("B7")), Cell(at("B7"), CellValue.Text("x")))
    assertEquals(removed.rangeHyperlinks.toVector, Vector.empty)
  }

  test("#931: clearing contents removes coverage; clearing styles does not") {
    val contents = sheet.clearRange(rng("D3:E3"), ClearWhat.contents)
    assertEquals(contents(at("D3")).hyperlink, None)
    assertEquals(contents(at("E2")).hyperlink, Some(link))
    assertEquals(
      sheet.clearRange(rng("D3:E3"), ClearWhat.styles).rangeHyperlinks.toVector,
      Vector(d2e3)
    )
  }

  test("#932: setHyperlink links one cell on the cell, a range as one range link") {
    val one = Sheet("S").setHyperlink(rng("A1:A1"), link)
    assertEquals(one.cells.hyperlinkAt(at("A1")), Some(link))
    assertEquals(one.rangeHyperlinks.toVector, Vector.empty)
    val many = Sheet("S").setHyperlink(rng("A1:A5"), link)
    assertEquals(many.rangeHyperlinks.toVector, Vector(RangeHyperlink(rng("A1:A5"), link)))
    assertEquals(many.cells.size, 0)
    assertEquals(many(at("A5")).hyperlink, Some(link))
  }

  test("#932: setHyperlink replaces what was linked in the range") {
    val owned = sheet.put(Cell(at("E3"), CellValue.Text("x")).withHyperlink(other))
    val relinked = owned.setHyperlink(rng("E2:F3"), other)
    assertEquals(relinked(at("E2")).hyperlink, Some(other))
    assertEquals(relinked(at("F3")).hyperlink, Some(other))
    assertEquals(relinked(at("D3")).hyperlink, Some(link))
    assertEquals(relinked.cells.hyperlinkAt(at("E3")), None, "E3's own link is replaced")
    assertEquals(relinked.rangeHyperlinkAt(at("E2")).map(_.range), Some(rng("E2:F3")))
  }

  test("#924 guard: structural edits leave range links where the file has them") {
    assertEquals(sheet.insertRows(0, 2).rangeHyperlinks.toVector, Vector(d2e3))
  }
