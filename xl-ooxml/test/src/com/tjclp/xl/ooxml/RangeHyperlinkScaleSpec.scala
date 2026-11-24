package com.tjclp.xl.ooxml

import scala.concurrent.duration.*
import scala.xml.{Elem, NamespaceBinding, Null, PrefixedAttribute, TopScope, UnprefixedAttribute}

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.Hyperlink
import com.tjclp.xl.ooxml.worksheet.{
  collectHyperlinks,
  sourceRangeHyperlinks,
  unmodeledHyperlinksElem
}
import com.tjclp.xl.sheets.{RangeHyperlink, Sheet}
import munit.FunSuite

/**
 * #931: matching a sheet's source range links against its model is linear. It was a scan of the
 * model per source element (and per model link in `collectHyperlinks`), so a write of an export
 * with one range link per row grew quadratically (40k links: 18.6 s against 4.6 s with none).
 *
 * 200k links is about 2·10^10 comparisons for the quadratic match and a second or two for the
 * linear one; the timeout is the backstop for that hang-class regression, not a speed threshold.
 */
class RangeHyperlinkScaleSpec extends FunSuite:

  // one large fixture: 200k links built, matched and collected twice
  override val munitTimeout: Duration = 120.seconds

  private val n = 200000

  private def rowRange(r: Int): CellRange = CellRange(ARef.from0(3, r), ARef.from0(4, r))

  /**
   * `<hyperlinks>` with one `<hyperlink ref="Dn:En" r:id="rLn"/>` per row, as the source had it.
   */
  private val preserved: Elem =
    val children = (0 until n).map { r =>
      val attrs = new UnprefixedAttribute(
        "ref",
        rowRange(r).toA1,
        new PrefixedAttribute("r", "id", s"rL$r", Null)
      )
      Elem(null, "hyperlink", attrs, TopScope, minimizeEmpty = true)
    }
    val scope = NamespaceBinding("r", XmlUtil.nsRelationships, TopScope)
    Elem(null, "hyperlinks", Null, scope, minimizeEmpty = false, children*)

  private val targets: Map[String, String] =
    (0 until n).iterator.map(r => s"rL$r" -> s"https://x.example/$r").toMap

  private val links: Vector[RangeHyperlink] = sourceRangeHyperlinks(preserved, targets)

  test("200k source range links match their model links in linear time") {
    assertEquals(links.size, n)
    val sheet = Sheet("S").copy(rangeHyperlinks = links)
    val kept = unmodeledHyperlinksElem(preserved, targets, sheet).getOrElse(fail("nothing kept"))
    assertEquals(kept.child.size, n)
    assertEquals(collectHyperlinks(sheet, Set.empty, links), Vector.empty)
  }

  test("with every other link removed, the kept and generated sets stay exact") {
    val sheet = Sheet("S").copy(rangeHyperlinks = links.zipWithIndex.collect {
      case (link, i) if i % 2 == 0 => link
    } :+ RangeHyperlink(rowRange(n), Hyperlink("https://new.example/")))
    val kept = unmodeledHyperlinksElem(preserved, targets, sheet).getOrElse(fail("nothing kept"))
    assertEquals(kept.child.size, n / 2)
    assertEquals(collectHyperlinks(sheet, Set.empty, links).map(_.ref), Vector(rowRange(n).toA1))
  }
