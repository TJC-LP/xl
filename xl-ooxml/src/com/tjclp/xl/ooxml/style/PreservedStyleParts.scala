package com.tjclp.xl.ooxml.style

import scala.xml.Elem

/**
 * Sections of a source `styles.xml` that xl does not model, carried verbatim into the rewritten
 * part (GH-610): the named-style master records (`cellStyleXfs`), their names (`cellStyles`), the
 * table-style defaults (`tableStyles`), the palette (`colors` — `mruColors`/`indexedColors`) and
 * the styles-level `extLst` (x14 slicer / x15 timeline style defaults).
 *
 * Same opaque-passthrough contract as `dxfs` and a worksheet's `extLst`: the fragments are the
 * parsed source elements, re-emitted in CT_Stylesheet order (ECMA-376 Part 1, §18.8.39). The one
 * thing the serializer touches is the `fontId`/`fillId`/`borderId` references inside
 * `cellStyleXfs`, which follow the component tables if those move (see [[OoxmlStyles]]); on an
 * Excel-authored source the tables stay positional and the fragment is byte-identical.
 *
 * An xl-authored workbook has none of these (`empty`), and the serializer falls back to the single
 * `Normal` master it has always written.
 */
final case class PreservedStyleParts(
  cellStyleXfs: Option[Elem] = None,
  cellStyles: Option[Elem] = None,
  tableStyles: Option[Elem] = None,
  colors: Option[Elem] = None,
  extLst: Option[Elem] = None
) derives CanEqual:

  /** True when the source carried none of the passthrough sections. */
  def isEmpty: Boolean =
    cellStyleXfs.isEmpty && cellStyles.isEmpty && tableStyles.isEmpty && colors.isEmpty &&
      extLst.isEmpty

  /**
   * Whether the named-style tables ride through. `cellStyles` is only honoured together with the
   * `cellStyleXfs` its `xfId`s point at — a source that has names without masters (or vice versa)
   * gets the single `Normal` pair, so no `xfId` can dangle in the output.
   */
  def hasNamedStyles: Boolean = cellStyleXfs.isDefined

object PreservedStyleParts:
  val empty: PreservedStyleParts = PreservedStyleParts()

  /** The passthrough sections of a parsed `<styleSheet>` (direct children only). */
  def fromXml(root: Elem): PreservedStyleParts =
    def child(label: String): Option[Elem] =
      (root \ label).headOption.collect { case e: Elem => e }
    PreservedStyleParts(
      cellStyleXfs = child("cellStyleXfs"),
      cellStyles = child("cellStyles"),
      tableStyles = child("tableStyles"),
      colors = child("colors"),
      extLst = child("extLst")
    )
