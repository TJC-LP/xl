package com.tjclp.xl.ooxml.style

import scala.xml.*

import com.tjclp.xl.ooxml.XmlUtil
import com.tjclp.xl.styles.Dxf

/**
 * Dxf table planning for conditional formatting (GH-136): preserved-prefix merge with typed-
 * equality index reuse and dedup-append.
 *
 * APPEND-ONLY INVARIANT: existing `<dxfs>` children never move or change — dxfIds baked into
 * Preserved rules/blocks, byte-copied sheets, and tables' dataDxfId stay valid forever. Source
 * children are re-emitted as the SAME Elem objects (verbatim, never re-encoded); fresh entries are
 * appended at `preservedCount..`. When nothing is appended, the original element rides through
 * untouched, so the dxfs table is a fixpoint under repeated edit→write cycles.
 */
object DxfTable:

  /**
   * @param dxfIds
   *   index assignment for every requested Dxf (existing index reused where a source entry parses
   *   typed-equal — first index wins; unparseable entries occupy indices but never match)
   * @param merged
   *   the `<dxfs>` element to ship: the untouched source object when nothing was appended, a
   *   source-prefix + appended-suffix rebuild when something was, None when there is nothing
   */
  final case class Plan(dxfIds: Map[Dxf, Int], merged: Option[Elem]) derives CanEqual

  /**
   * Plan dxf emission: resolve each needed Dxf to an existing source index (typed equality) or
   * dedup-append after the preserved prefix, in `needed` encounter order.
   */
  def plan(preservedDxfs: Option[Elem], needed: Vector[Dxf]): Plan =
    val sourceChildren: Vector[Elem] =
      preservedDxfs.map(e => XmlUtil.getChildren(e, "dxf").toVector).getOrElse(Vector.empty)
    // First index wins; unparseable entries occupy indices, never matched.
    val existing: Map[Dxf, Int] =
      sourceChildren.zipWithIndex.foldLeft(Map.empty[Dxf, Int]) { case (acc, (child, idx)) =>
        DxfCodec.parse(child).filterNot(acc.contains).fold(acc)(dxf => acc + (dxf -> idx))
      }
    val (ids, appended) =
      needed.foldLeft((Map.empty[Dxf, Int], Vector.empty[Dxf])) { case ((acc, app), dxf) =>
        if acc.contains(dxf) then (acc, app)
        else
          existing.get(dxf) match
            case Some(idx) => (acc + (dxf -> idx), app)
            case None => (acc + (dxf -> (sourceChildren.size + app.size)), app :+ dxf)
      }
    val merged: Option[Elem] =
      if appended.isEmpty then preservedDxfs
      else
        val children: Seq[Node] = sourceChildren ++ appended.map(DxfCodec.toXml)
        val count = children.size.toString
        preservedDxfs match
          case Some(src) =>
            val withCount = src % new UnprefixedAttribute("count", count, Null)
            Some(withCount.copy(child = children))
          case None => Some(XmlUtil.elem("dxfs", "count" -> count)(children*))
    Plan(ids, merged)
