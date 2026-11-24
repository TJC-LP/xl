package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{ARef, CellRange}

/**
 * A sheet's range hyperlinks (#931), in model order: the value of `Sheet.rangeHyperlinks`.
 *
 * An immutable value compared by its links. It also answers "which link covers this cell" without
 * scanning every link: past [[RangeHyperlinkIndex.ScanLimit]] links, the first lookup builds a
 * [[RangeHyperlinkIndex]] held in a lazy field of this value. A sheet's copies share the value, so
 * a run of puts or a recalculation builds the index once; an edit of the links makes a new value
 * with a new index. A sheet can hold one link per row (80k on an 80k-row export), and `Sheet.apply`
 * looks one up for every read of a cell without its own link, so a scan made whole-sheet passes
 * quadratic.
 */
final class RangeHyperlinks private (val links: Vector[RangeHyperlink]) extends Serializable:

  /** Serialized as its links alone (the CellStore pattern): the index is rebuilt on demand. */
  private def writeReplace(): AnyRef = new RangeHyperlinks.Serialized(links)

  /** The index for lookups, built on the first lookup past the scan limit. */
  private[sheets] lazy val index: RangeHyperlinkIndex = RangeHyperlinkIndex(links)

  def isEmpty: Boolean = links.isEmpty
  def nonEmpty: Boolean = links.nonEmpty
  def size: Int = links.size
  def iterator: Iterator[RangeHyperlink] = links.iterator
  def toVector: Vector[RangeHyperlink] = links

  /** The latest link in model order that covers `ref`. */
  def at(ref: ARef): Option[RangeHyperlink] =
    if links.isEmpty then None
    else if links.sizeIs <= RangeHyperlinkIndex.ScanLimit then links.findLast(_.covers(ref))
    else index.at(ref)

  /** These links and `link` after them. */
  def :+(link: RangeHyperlink): RangeHyperlinks = new RangeHyperlinks(links :+ link)

  /** Every link with `cut` removed from its coverage; links left with nothing are dropped. */
  def without(cut: CellRange): RangeHyperlinks =
    if !links.exists(_.range.intersects(cut)) then this
    else new RangeHyperlinks(links.flatMap(_.without(cut)))

  override def equals(that: Any): Boolean = that match
    case other: RangeHyperlinks => (this eq other) || links == other.links
    case _ => false

  override def hashCode: Int = links.hashCode

  override def toString: String = links.mkString("RangeHyperlinks(", ", ", ")")

object RangeHyperlinks:
  val empty: RangeHyperlinks = new RangeHyperlinks(Vector.empty)

  def apply(links: RangeHyperlink*): RangeHyperlinks = from(links)

  def from(links: IterableOnce[RangeHyperlink]): RangeHyperlinks =
    val v = Vector.from(links)
    if v.isEmpty then empty else new RangeHyperlinks(v)

  /** The serial form: the links. `readResolve` rebuilds the value, with no index yet. */
  private final class Serialized(private val links: Vector[RangeHyperlink]) extends Serializable:
    private def readResolve(): AnyRef = from(links)
