package com.tjclp.xl.ooxml

import scala.xml.*
import XmlUtil.*
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.error.XLError
import com.tjclp.xl.tables.{
  TableSpec,
  TableColumn as DomainTableColumn,
  TableAutoFilter,
  TableStyle
}

/**
 * OOXML table column in xl/tables/tableN.xml.
 *
 * Each column has a unique ID and display name.
 *
 * OOXML structure:
 * {{{
 * <tableColumn id="1" name="Product"/>
 * }}}
 *
 * @param id
 *   Column identifier (1-indexed)
 * @param name
 *   Column display name
 * @param otherAttrs
 *   Unknown attributes for forwards compatibility
 */
final case class OoxmlTableColumn(
  id: Long,
  name: String,
  otherAttrs: Map[String, String] = Map.empty
)

/**
 * OOXML table in xl/tables/tableN.xml.
 *
 * Maps to Excel's table model with columns, AutoFilter, and styling. Supports forwards
 * compatibility via otherAttrs/otherChildren for unknown properties.
 *
 * OOXML structure:
 * {{{
 * <table id="1" name="Table1" displayName="Table1" ref="A1:D10" headerRowCount="1" totalsRowCount="0">
 *   <autoFilter ref="A1:D10"/>
 *   <tableColumns count="4">
 *     <tableColumn id="1" name="Product"/>
 *     <tableColumn id="2" name="Price"/>
 *     <tableColumn id="3" name="Quantity"/>
 *     <tableColumn id="4" name="Total"/>
 *   </tableColumns>
 *   <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0"
 *                   showRowStripes="1" showColumnStripes="0"/>
 * </table>
 * }}}
 *
 * @param id
 *   Table identifier (1-indexed, unique within workbook)
 * @param name
 *   Internal table name (used in formulas)
 * @param displayName
 *   User-visible table name (shown in Name Manager)
 * @param ref
 *   Table range (includes header row)
 * @param headerRowCount
 *   Number of header rows (typically 1)
 * @param totalsRowCount
 *   Number of totals rows (0 = no totals)
 * @param columns
 *   Column definitions
 * @param autoFilter
 *   Optional AutoFilter range
 * @param styleInfo
 *   Optional table style information
 * @param otherAttrs
 *   Unknown attributes for forwards compatibility
 * @param otherChildren
 *   Unknown child elements for forwards compatibility
 */
final case class OoxmlTable(
  id: Long,
  name: String,
  displayName: String,
  ref: CellRange,
  headerRowCount: Int,
  totalsRowCount: Int,
  columns: Vector[OoxmlTableColumn],
  autoFilter: Option[CellRange],
  styleInfo: Option[OoxmlTableStyleInfo],
  otherAttrs: Map[String, String] = Map.empty,
  otherChildren: Seq[Elem] = Seq.empty
)

/**
 * OOXML table style information.
 *
 * OOXML structure:
 * {{{
 * <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0"
 *                 showRowStripes="1" showColumnStripes="0"/>
 * }}}
 *
 * @param name
 *   Style name (e.g., TableStyleMedium2)
 * @param showFirstColumn
 *   Whether to emphasize first column
 * @param showLastColumn
 *   Whether to emphasize last column
 * @param showRowStripes
 *   Whether to show row banding
 * @param showColumnStripes
 *   Whether to show column banding
 * @param otherAttrs
 *   Unknown attributes for forwards compatibility
 */
final case class OoxmlTableStyleInfo(
  name: String,
  showFirstColumn: Boolean = false,
  showLastColumn: Boolean = false,
  showRowStripes: Boolean = true,
  showColumnStripes: Boolean = false,
  otherAttrs: Map[String, String] = Map.empty
)

object OoxmlTable extends XmlReadable[OoxmlTable] with XmlWritable:

  /**
   * Parse table from XML.
   *
   * REQUIRES: elem is <table> element from xl/tables/tableN.xml ENSURES:
   *   - Returns OoxmlTable with all columns and optional AutoFilter
   *   - Preserves unknown attributes/children for forwards compatibility
   *   - Returns error if structure is invalid (missing required attributes)
   * DETERMINISTIC: Yes (stable iteration order)
   *
   * @param elem
   *   The <table> root element
   * @return
   *   Either[String, OoxmlTable] with error if parsing fails
   */
  def fromXml(elem: Elem): Either[String, OoxmlTable] =
    if elem.label != "table" then Left(s"Expected <table> but found <${elem.label}>")
    else
      for
        idStr <- getAttr(elem, "id")
        id <- idStr.toLongOption.toRight(s"Invalid table id: $idStr")
        name <- getAttr(elem, "name")
        displayName <- getAttr(elem, "displayName")
        refStr <- getAttr(elem, "ref")
        range <- CellRange.parse(refStr).left.map(err => s"Invalid table ref '$refStr': $err")
      yield
        val headerCount = getAttrOpt(elem, "headerRowCount").flatMap(_.toIntOption).getOrElse(1)
        val totalsCount = getAttrOpt(elem, "totalsRowCount").flatMap(_.toIntOption).getOrElse(0)

        // Parse columns
        val columnsElem = (elem \ "tableColumns").headOption
        val columns = columnsElem.toList.flatMap { colsElem =>
          (colsElem \ "tableColumn").collect { case c: Elem => c }.map(decodeColumn)
        }.toVector

        // Parse AutoFilter
        val autoFilterRange =
          (elem \ "autoFilter").headOption.collect { case af: Elem => af }.flatMap { afElem =>
            getAttrOpt(afElem, "ref").flatMap(CellRange.parse(_).toOption)
          }

        // Parse table style info
        val styleInfo = (elem \ "tableStyleInfo").headOption
          .collect { case si: Elem => si }
          .map(decodeStyleInfo)

        // Preserve unknown attributes/children
        val knownAttrs = Set(
          "id",
          "name",
          "displayName",
          "ref",
          "headerRowCount",
          "totalsRowCount"
        )
        val attrs = elem.attributes.asAttrMap.filterNot { case (k, _) => knownAttrs.contains(k) }

        val knownChildren =
          (elem \ "tableColumns") ++ (elem \ "autoFilter") ++ (elem \ "tableStyleInfo")
        val others = elem.child.collect {
          case el: Elem if !knownChildren.contains(el) => el
        }.toVector

        OoxmlTable(
          id = id,
          name = name,
          displayName = displayName,
          ref = range,
          headerRowCount = headerCount,
          totalsRowCount = totalsCount,
          columns = columns,
          autoFilter = autoFilterRange,
          styleInfo = styleInfo,
          otherAttrs = attrs,
          otherChildren = others
        )

  /**
   * Parse table column from XML.
   *
   * @param elem
   *   The <tableColumn> element
   * @return
   *   OoxmlTableColumn
   */
  private def decodeColumn(elem: Elem): OoxmlTableColumn =
    val id = getAttrOpt(elem, "id").flatMap(_.toLongOption).getOrElse(0L)
    val name = getAttrOpt(elem, "name").getOrElse("")

    val known = Set("id", "name")
    val attrs = elem.attributes.asAttrMap.filterNot { case (k, _) => known.contains(k) }

    OoxmlTableColumn(id, name, attrs)

  /**
   * Parse table style info from XML.
   *
   * @param elem
   *   The <tableStyleInfo> element
   * @return
   *   OoxmlTableStyleInfo
   */
  private def decodeStyleInfo(elem: Elem): OoxmlTableStyleInfo =
    val name = getAttrOpt(elem, "name").getOrElse("TableStyleMedium2")
    val showFirstColumn = getAttrOpt(elem, "showFirstColumn").exists(v => v == "1" || v == "true")
    val showLastColumn = getAttrOpt(elem, "showLastColumn").exists(v => v == "1" || v == "true")
    val showRowStripes = getAttrOpt(elem, "showRowStripes").forall(v => v == "1" || v == "true")
    val showColumnStripes =
      getAttrOpt(elem, "showColumnStripes").exists(v => v == "1" || v == "true")

    val known = Set(
      "name",
      "showFirstColumn",
      "showLastColumn",
      "showRowStripes",
      "showColumnStripes"
    )
    val attrs = elem.attributes.asAttrMap.filterNot { case (k, _) => known.contains(k) }

    OoxmlTableStyleInfo(
      name = name,
      showFirstColumn = showFirstColumn,
      showLastColumn = showLastColumn,
      showRowStripes = showRowStripes,
      showColumnStripes = showColumnStripes,
      otherAttrs = attrs
    )

  /**
   * Serialize table to XML with deterministic attribute ordering.
   *
   * REQUIRES: table is valid OoxmlTable ENSURES:
   *   - Returns <table> element with all required attributes
   *   - Attributes are sorted for deterministic output
   *   - Includes all children (tableColumns, autoFilter, tableStyleInfo)
   *   - Preserves unknown attributes/children
   * DETERMINISTIC: Yes (sorted attributes)
   *
   * @return
   *   XML element for xl/tables/tableN.xml
   */
  def toXml(table: OoxmlTable): Elem =
    val baseAttrs = Seq(
      "id" -> table.id.toString,
      "name" -> table.name,
      "displayName" -> table.displayName,
      "ref" -> table.ref.toA1,
      "headerRowCount" -> table.headerRowCount.toString,
      "totalsRowCount" -> table.totalsRowCount.toString
    ) ++ table.otherAttrs.toSeq

    val children = Seq(
      // tableColumns
      elem(
        "tableColumns",
        "count" -> table.columns.size.toString
      )(table.columns.map(encodeColumn)*),
      // autoFilter (optional)
      table.autoFilter.map(range => elem("autoFilter", "ref" -> range.toA1)()).toSeq,
      // tableStyleInfo (optional)
      table.styleInfo.map(encodeStyleInfo).toSeq,
      // Other unknown children
      table.otherChildren
    ).flatten

    elem("table", baseAttrs*)(children*)

  /**
   * Serialize table column to XML.
   *
   * @param col
   *   The column to serialize
   * @return
   *   XML element
   */
  private def encodeColumn(col: OoxmlTableColumn): Elem =
    val attrs = Seq(
      "id" -> col.id.toString,
      "name" -> col.name
    ) ++ col.otherAttrs.toSeq

    elem("tableColumn", attrs*)()

  /**
   * Serialize table style info to XML.
   *
   * @param info
   *   The style info to serialize
   * @return
   *   XML element
   */
  private def encodeStyleInfo(info: OoxmlTableStyleInfo): Elem =
    val attrs = Seq(
      "name" -> info.name,
      "showFirstColumn" -> (if info.showFirstColumn then "1" else "0"),
      "showLastColumn" -> (if info.showLastColumn then "1" else "0"),
      "showRowStripes" -> (if info.showRowStripes then "1" else "0"),
      "showColumnStripes" -> (if info.showColumnStripes then "1" else "0")
    ) ++ info.otherAttrs.toSeq

    elem("tableStyleInfo", attrs*)()

  override def toXml: Elem =
    throw new UnsupportedOperationException("Use toXml(table: OoxmlTable) instead")

/**
 * Conversion between domain TableSpec and OOXML representation.
 */
object TableConversions:

  /**
   * Convert domain TableSpec to OOXML representation.
   *
   * @param spec
   *   Domain table specification
   * @param id
   *   Table ID (1-indexed, unique within workbook)
   * @return
   *   OOXML table
   */
  def toOoxml(spec: TableSpec, id: Long): OoxmlTable =
    val columns = spec.columns.map { col =>
      OoxmlTableColumn(col.id, col.name)
    }

    val autoFilterRange = spec.autoFilter.filter(_.enabled).map(_ => spec.range)

    val styleInfo = styleToStyleInfo(spec.style)

    val headerCount = if spec.showHeaderRow then 1 else 0
    val totalsCount = if spec.showTotalsRow then 1 else 0

    OoxmlTable(
      id = id,
      name = spec.name,
      displayName = spec.displayName,
      ref = spec.range,
      headerRowCount = headerCount,
      totalsRowCount = totalsCount,
      columns = columns,
      autoFilter = autoFilterRange,
      styleInfo = Some(styleInfo)
    )

  /**
   * Convert OOXML table to domain TableSpec.
   *
   * @param ooxml
   *   OOXML table
   * @return
   *   Domain table specification
   */
  def fromOoxml(ooxml: OoxmlTable): TableSpec =
    val columns = ooxml.columns.map { col =>
      DomainTableColumn(col.id, col.name)
    }

    val autoFilter = ooxml.autoFilter.map(_ => TableAutoFilter(enabled = true))

    val style = ooxml.styleInfo.map(styleInfoToStyle).getOrElse(TableStyle.default)

    TableSpec(
      name = ooxml.name,
      displayName = ooxml.displayName,
      range = ooxml.ref,
      columns = columns,
      showHeaderRow = ooxml.headerRowCount > 0,
      showTotalsRow = ooxml.totalsRowCount > 0,
      autoFilter = autoFilter,
      style = style
    )

  /**
   * Convert domain TableStyle to OOXML style info.
   *
   * @param style
   *   Domain table style
   * @return
   *   OOXML style info
   */
  private def styleToStyleInfo(style: TableStyle): OoxmlTableStyleInfo =
    val name = style match
      case TableStyle.None => ""
      case TableStyle.Light(n) => s"TableStyleLight$n"
      case TableStyle.Medium(n) => s"TableStyleMedium$n"
      case TableStyle.Dark(n) => s"TableStyleDark$n"

    OoxmlTableStyleInfo(
      name = name,
      showFirstColumn = false,
      showLastColumn = false,
      showRowStripes = true,
      showColumnStripes = false
    )

  /**
   * Parse OOXML style name to domain TableStyle.
   *
   * @param info
   *   OOXML style info
   * @return
   *   Domain table style
   */
  private def styleInfoToStyle(info: OoxmlTableStyleInfo): TableStyle =
    info.name match
      case s if s.startsWith("TableStyleLight") =>
        s.stripPrefix("TableStyleLight")
          .toIntOption
          .map(TableStyle.Light.apply)
          .getOrElse(TableStyle.default)
      case s if s.startsWith("TableStyleMedium") =>
        s.stripPrefix("TableStyleMedium")
          .toIntOption
          .map(TableStyle.Medium.apply)
          .getOrElse(TableStyle.default)
      case s if s.startsWith("TableStyleDark") =>
        s.stripPrefix("TableStyleDark")
          .toIntOption
          .map(TableStyle.Dark.apply)
          .getOrElse(TableStyle.default)
      case _ => TableStyle.default
