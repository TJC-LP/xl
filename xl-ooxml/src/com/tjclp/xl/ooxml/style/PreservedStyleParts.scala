package com.tjclp.xl.ooxml.style

import scala.xml.Elem

import com.tjclp.xl.ooxml.XmlUtil

/**
 * Sections of a source `styles.xml` that xl does not model, carried verbatim into the rewritten
 * part (GH-610): the named-style master records (`cellStyleXfs`), their names (`cellStyles`), the
 * source's own cell records (`cellXfs`, emitted verbatim for every slot the writer keeps
 * positional), the table-style defaults (`tableStyles`), the palette (`colors` —
 * `mruColors`/`indexedColors`) and the styles-level `extLst` (x14 slicer / x15 timeline style
 * defaults).
 *
 * Same opaque-passthrough contract as `dxfs` and a worksheet's `extLst`: the fragments are the
 * parsed source elements, re-emitted in CT_Stylesheet order (ECMA-376 Part 1, §18.8.39). The one
 * thing the serializer touches is the `fontId`/`fillId`/`borderId` references inside the `<xf>`
 * records (and a cellXf's `xfId`), which follow the component tables if those move (see
 * [[OoxmlStyles]]); on an Excel-authored source the tables stay positional and the fragments are
 * byte-identical.
 *
 * An xl-authored workbook has none of these (`empty`), and the serializer falls back to the single
 * `Normal` master it has always written.
 */
final case class PreservedStyleParts(
  cellStyleXfs: Option[Elem] = None,
  cellXfs: Option[Elem] = None,
  cellStyles: Option[Elem] = None,
  tableStyles: Option[Elem] = None,
  colors: Option[Elem] = None,
  extLst: Option[Elem] = None
) derives CanEqual:

  /** True when the source carried none of the passthrough sections. */
  def isEmpty: Boolean =
    cellStyleXfs.isEmpty && cellXfs.isEmpty && cellStyles.isEmpty && tableStyles.isEmpty &&
      colors.isEmpty && extLst.isEmpty

  /** Number of `<xf>` masters in the preserved `cellStyleXfs`; 0 when absent or empty. */
  def masterCount: Int = cellStyleXfs.map(XmlUtil.getChildren(_, "xf").size).getOrElse(0)

  /**
   * Whether the named-style tables ride through: only when `cellStyleXfs` has at least one master.
   * `cellStyles` is honoured only together with the masters its `xfId`s point at, and an empty
   * `<cellStyleXfs count="0"/>` (which every cellXf `xfId` would dangle from) falls back to the
   * single `Normal` pair — xl introduces no dangling reference and repairs a source's own.
   */
  def hasNamedStyles: Boolean = masterCount > 0

object PreservedStyleParts:
  val empty: PreservedStyleParts = PreservedStyleParts()

  /** The passthrough sections of a parsed `<styleSheet>` (direct children only). */
  def fromXml(root: Elem): PreservedStyleParts =
    def child(label: String): Option[Elem] =
      (root \ label).headOption.collect { case e: Elem => e }
    PreservedStyleParts(
      cellStyleXfs = child("cellStyleXfs"),
      cellXfs = child("cellXfs"),
      cellStyles = child("cellStyles"),
      tableStyles = child("tableStyles"),
      colors = child("colors"),
      extLst = child("extLst")
    )
