package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.Hyperlink
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/**
 * The range-hyperlink index (#931) returns exactly what the model-order scan it replaced returned,
 * and a lookup visits O(log n) tree nodes on the one-link-per-row shape, pinned by a visit count at
 * two sizes rather than by a clock.
 */
class RangeHyperlinkIndexSpec extends ScalaCheckSuite:

  private def rng(s: String): CellRange = CellRange.parse(s).fold(e => fail(e), identity)

  /** The pre-index implementation of Sheet.rangeHyperlinkAt, kept as the reference. */
  private def scan(links: Vector[RangeHyperlink], ref: ARef): Option[RangeHyperlink] =
    links.findLast(_.covers(ref))

  private val genRange: Gen[CellRange] =
    for
      c1 <- Gen.choose(0, 9); c2 <- Gen.choose(0, 9)
      r1 <- Gen.choose(0, 30); r2 <- Gen.choose(0, 30)
    yield CellRange(ARef.from0(c1, r1), ARef.from0(c2, r2))

  private val genLink: Gen[RangeHyperlink] =
    for
      range <- genRange
      target <- Gen.oneOf("https://a.example/", "https://b.example/", "Sheet1!A1")
      holes <- Gen.listOfN(2, genRange).map(_.toVector)
      cut <- Gen.oneOf(true, false)
    yield RangeHyperlink(range, Hyperlink(target), excluded = if cut then holes else Vector.empty)

  private val genRef: Gen[ARef] =
    for c <- Gen.choose(0, 10); r <- Gen.choose(0, 32) yield ARef.from0(c, r)

  property("index lookup ≡ the latest covering link in model order") {
    Prop.forAll(Gen.choose(0, 80).flatMap(Gen.listOfN(_, genLink)), Gen.listOfN(20, genRef)) {
      (links, refs) =>
        val index = RangeHyperlinkIndex(links.toVector)
        refs.foreach(ref => assertEquals(index.at(ref), scan(links.toVector, ref), ref.toA1))
    }
  }

  property("Sheet.rangeHyperlinkAt ≡ the scan, below and above the scan limit") {
    Prop.forAll(Gen.choose(0, 60).flatMap(Gen.listOfN(_, genLink)), genRef) { (links, ref) =>
      val sheet = Sheet("S").copy(rangeHyperlinks = links.toVector)
      assertEquals(sheet.rangeHyperlinkAt(ref), scan(links.toVector, ref))
    }
  }

  /** One link per row over D:E, the export shape that made the scan quadratic. */
  private def perRow(n: Int): Vector[RangeHyperlink] =
    Vector.tabulate(n)(r =>
      RangeHyperlink(CellRange(ARef.from0(3, r), ARef.from0(4, r)), Hyperlink(s"https://x/$r"))
    )

  test("a lookup visits O(log n) nodes on one link per row (1k vs 64k links)") {
    def worst(n: Int): Int =
      val index = RangeHyperlinkIndex(perRow(n))
      val probes = Seq(0, n / 3, n / 2, n - 1).flatMap(r => Seq(ARef.from0(3, r), ARef.from0(9, r)))
      probes.foreach(ref =>
        assertEquals(index.at(ref), scan(perRow(n), ref).filter(_ => ref.col.index0 <= 4))
      )
      probes.map(index.lookup(_)._2).max
    val small = worst(1 << 10)
    val large = worst(1 << 16)
    // log2 grows 10 -> 16; a scan would grow 64x. The descent touches two nodes per level.
    assert(small <= 4 * 12, s"1k links: $small nodes")
    assert(large <= 4 * 18, s"64k links: $large nodes")
  }

  test("a whole-column link and many row links: the column link is found for every row") {
    val column = RangeHyperlink(rng("A1:A1048576"), Hyperlink("https://col/"))
    val links = perRow(100) :+ column
    val index = RangeHyperlinkIndex(links)
    assertEquals(index.at(ARef.from0(0, 1048575)), Some(column))
    assertEquals(index.at(ARef.from0(3, 50)).map(_.link.target), Some("https://x/50"))
  }
