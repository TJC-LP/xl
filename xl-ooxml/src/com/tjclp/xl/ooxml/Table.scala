package com.tjclp.xl.ooxml

import scala.xml.*
import XmlUtil.*
import SaxSupport.*
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.error.XLError
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.tables.{
  TableSpec,
  TableColumn as DomainTableColumn,
  TableAutoFilter,
  TableStyle,
  TotalsRowFunction
}

/**
 * OOXML table column in xl/tables/tableN.xml.
 *
 * Each column has a unique ID and display name, and — when the table shows a totals row — what its
 * totals-row cell holds: a label, or an aggregate (`totalsRowFunction`, with `<totalsRowFormula>`
 * for `custom`).
 *
 * OOXML structure:
 * {{{
 * <tableColumn id="1" name="Product" totalsRowLabel="Total"/>
 * <tableColumn id="2" name="Amount" totalsRowFunction="sum"/>
 * <tableColumn id="3" name="Tax" totalsRowFunction="custom"><totalsRowFormula>SUM(Table1[Tax])*0.2</totalsRowFormula></tableColumn>
 * }}}
 *
 * @param id
 *   Column identifier (1-indexed)
 * @param name
 *   Column display name
 * @param otherAttrs
 *   Unknown attributes for forwards compatibility
 * @param totalsRowLabel
 *   `totalsRowLabel`: the totals-row cell's text
 * @param totalsRowFunction
 *   `totalsRowFunction`: the `ST_TotalsRowFunction` token of the totals-row aggregate
 * @param totalsRowFormula
 *   the `<totalsRowFormula>` child a `custom` aggregate carries (no leading `=`)
 */
final case class OoxmlTableColumn(
  id: Long,
  name: String,
  uid: Option[String] = None, // xr3:uid for Excel 2016+ revision tracking
  dataDxfId: Option[Int] = None, // Optional data formatting ID
  otherAttrs: Map[String, String] = Map.empty,
  totalsRowLabel: Option[String] = None,
  totalsRowFunction: Option[String] = None,
  totalsRowFormula: Option[String] = None
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
 *   Number of totals rows (0 = no totals) — the attribute that says whether the table SHOWS a
 *   totals row (`totalsRowCount`, default 0)
 * @param totalsRowShown
 *   `totalsRowShown` (default true per ECMA-376 §18.5.1.2): whether a totals row has EVER been
 *   shown for this table — Excel writes `totalsRowShown="0"` on a table that never had one and
 *   omits it otherwise; it does not say whether one is shown now (that is `totalsRowCount`)
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
  totalsRowShown: Boolean = false, // Excel attribute name differs from count
  columns: Vector[OoxmlTableColumn],
  autoFilter: Option[CellRange],
  autoFilterUid: Option[String] = None, // xr:uid for autoFilter element
  styleInfo: Option[OoxmlTableStyleInfo],
  tableUid: Option[String] = None, // xr:uid for table element
  otherAttrs: Map[String, String] = Map.empty,
  otherChildren: Seq[Elem] = Seq.empty
) extends XmlWritable,
      SaxSerializable:

  def toXml: Elem = OoxmlTable.toXml(this)

  def writeSax(writer: SaxWriter): Unit =
    writer.startDocument()
    writer.writeElem(toXml)
    writer.endDocument()
    writer.flush()

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

object OoxmlTable extends XmlReadable[OoxmlTable]:

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

        // `totalsRowCount` (default 0) is the totals row shown NOW; `totalsRowShown` (default
        // true) only records that one has ever been shown — a table Excel saved with its totals
        // row hidden carries totalsRowShown="1" and no count, and shows NO totals row. Reading
        // the flag as the count promoted the last data row of such a table to a totals row.
        val totalsCount = getAttrOpt(elem, "totalsRowCount").flatMap(_.toIntOption).getOrElse(0)
        val totalsShown =
          getAttrOpt(elem, "totalsRowShown").fold(true)(v => v == "1" || v == "true")

        // Parse table UID (use asAttrMap for prefixed attributes)
        val tableUid = elem.attributes.asAttrMap.get("xr:uid")

        // Parse columns
        val columnsElem = (elem \ "tableColumns").headOption
        val columns = columnsElem.toList.flatMap { colsElem =>
          (colsElem \ "tableColumn").collect { case c: Elem => c }.map(decodeColumn)
        }.toVector

        // Parse AutoFilter with UID
        val (autoFilterRange, autoFilterUid) =
          (elem \ "autoFilter").headOption
            .collect { case af: Elem => af }
            .map { afElem =>
              val range = getAttrOpt(afElem, "ref").flatMap(CellRange.parse(_).toOption)
              val uid =
                afElem.attributes.asAttrMap.get("xr:uid") // Use asAttrMap for prefixed attrs
              (range, uid)
            }
            .getOrElse((None, None))

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
          "totalsRowCount",
          "totalsRowShown",
          "xr:uid" // Don't preserve as otherAttr (handled explicitly)
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
          totalsRowShown = totalsShown,
          columns = columns,
          autoFilter = autoFilterRange,
          autoFilterUid = autoFilterUid,
          styleInfo = styleInfo,
          tableUid = tableUid,
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
    val uid = elem.attributes.asAttrMap.get("xr3:uid") // Use asAttrMap for prefixed attrs
    val dataDxfId = getAttrOpt(elem, "dataDxfId").flatMap(_.toIntOption)
    val totalsRowLabel = getAttrOpt(elem, "totalsRowLabel")
    val totalsRowFunction = getAttrOpt(elem, "totalsRowFunction")
    val totalsRowFormula =
      (elem \ "totalsRowFormula").headOption.map(_.text).filter(_.nonEmpty)

    val known = Set("id", "name", "xr3:uid", "dataDxfId", "totalsRowLabel", "totalsRowFunction")
    val attrs = elem.attributes.asAttrMap.filterNot { case (k, _) => known.contains(k) }

    OoxmlTableColumn(
      id = id,
      name = name,
      uid = uid,
      dataDxfId = dataDxfId,
      otherAttrs = attrs,
      totalsRowLabel = totalsRowLabel,
      totalsRowFunction = totalsRowFunction,
      totalsRowFormula = totalsRowFormula
    )

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
    // CRITICAL: Excel requires xmlns declarations FIRST in attribute list
    // Scala XML NamespaceBinding serializes namespaces LAST, so we must treat xmlns as regular attributes

    // Build attributes in REVERSE order (linked list serializes backwards)
    // Excel expects: xmlns, xmlns:mc, mc:Ignorable, xmlns:xr, xmlns:xr3, id, xr:uid, name, displayName, ref, totalsRowCount|totalsRowShown
    // So build in reverse: totals → ref → displayName → name → xr:uid → id → xmlns:xr3 → xmlns:xr → mc:Ignorable → xmlns:mc → xmlns

    // Start with other attributes
    val attrs1 = table.otherAttrs.foldLeft(scala.xml.Null: scala.xml.MetaData) {
      case (acc, (k, v)) =>
        new scala.xml.UnprefixedAttribute(k, v, acc)
    }

    // Regular attributes (in reverse of target order)
    // NOTE: displayName validation enforced in TableSpec.create smart constructor
    // Excel's shape: a shown totals row is `totalsRowCount="N"` (totalsRowShown then defaults to
    // true and is omitted); a table that never had one is `totalsRowShown="0"`; one that had and
    // hid it carries neither. Writing only totalsRowShown="1" declared NO totals row (count 0).
    val attrs2 =
      if table.totalsRowCount > 0 then
        new scala.xml.UnprefixedAttribute("totalsRowCount", table.totalsRowCount.toString, attrs1)
      else if !table.totalsRowShown then
        new scala.xml.UnprefixedAttribute("totalsRowShown", "0", attrs1)
      else attrs1
    val attrs3 = new scala.xml.UnprefixedAttribute("ref", table.ref.toA1, attrs2)
    val attrs4 = new scala.xml.UnprefixedAttribute("displayName", table.displayName, attrs3)
    val attrs5 = new scala.xml.UnprefixedAttribute("name", table.name, attrs4)
    val attrs6 = table.tableUid match
      case Some(uid) => new scala.xml.PrefixedAttribute("xr", "uid", uid, attrs5)
      case None => attrs5
    val attrs7 = new scala.xml.UnprefixedAttribute("id", table.id.toString, attrs6)

    // Namespace declarations (as regular attributes, in reverse order)
    val attrs8 = new scala.xml.UnprefixedAttribute(
      "xmlns:xr3",
      "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
      attrs7
    )
    val attrs9 = new scala.xml.UnprefixedAttribute(
      "xmlns:xr",
      "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
      attrs8
    )
    val attrs10 = new scala.xml.PrefixedAttribute("mc", "Ignorable", "xr xr3", attrs9)
    val attrs11 = new scala.xml.UnprefixedAttribute(
      "xmlns:mc",
      "http://schemas.openxmlformats.org/markup-compatibility/2006",
      attrs10
    )
    val finalAttrs = new scala.xml.UnprefixedAttribute("xmlns", nsSpreadsheetML, attrs11)

    // Build child elements: autoFilter → tableColumns → tableStyleInfo
    val autoFilterElem = table.autoFilter.map { range =>
      val afUidAttr = table.autoFilterUid match
        case Some(uid) => scala.xml.Attribute("xr", "uid", uid, scala.xml.Null)
        case None => scala.xml.Null

      scala.xml.Elem(
        null,
        "autoFilter",
        new scala.xml.UnprefixedAttribute("ref", range.toA1, afUidAttr),
        scala.xml.TopScope,
        minimizeEmpty = true
      )
    }

    val tableColumnsElem = elem("tableColumns", "count" -> table.columns.size.toString)(
      table.columns.map(encodeColumn)*
    )

    val children = Seq(
      autoFilterElem.toList,
      Seq(tableColumnsElem),
      table.styleInfo.map(encodeStyleInfo).toList,
      table.otherChildren
    ).flatten

    // Use TopScope since we're handling namespaces as regular attributes
    scala.xml.Elem(null, "table", finalAttrs, scala.xml.TopScope, minimizeEmpty = false, children*)

  /**
   * Serialize table column to XML.
   *
   * @param col
   *   The column to serialize
   * @return
   *   XML element
   */
  private def encodeColumn(col: OoxmlTableColumn): Elem =
    // Build attributes in reverse order (linked list serializes backwards)
    // Excel expects (schema order): id, xr3:uid, name, totalsRowFunction, totalsRowLabel, dataDxfId
    // Build order: dataDxfId, totalsRowLabel, totalsRowFunction, name, xr3:uid, id

    val attrs1 = col.otherAttrs.foldLeft(scala.xml.Null: scala.xml.MetaData) { case (acc, (k, v)) =>
      new scala.xml.UnprefixedAttribute(k, v, acc)
    }

    val attrs2 = col.dataDxfId match
      case Some(id) => new scala.xml.UnprefixedAttribute("dataDxfId", id.toString, attrs1)
      case None => attrs1

    val attrs2b = col.totalsRowLabel match
      case Some(label) => new scala.xml.UnprefixedAttribute("totalsRowLabel", label, attrs2)
      case None => attrs2

    val attrs2c = col.totalsRowFunction match
      case Some(fn) => new scala.xml.UnprefixedAttribute("totalsRowFunction", fn, attrs2b)
      case None => attrs2b

    val attrs3 = new scala.xml.UnprefixedAttribute("name", col.name, attrs2c)

    val attrs4 = col.uid match
      case Some(uid) => new scala.xml.PrefixedAttribute("xr3", "uid", uid, attrs3)
      case None => attrs3

    val finalAttrs = new scala.xml.UnprefixedAttribute("id", col.id.toString, attrs4)

    // a `custom` aggregate's formula is the column's one child (calculatedColumnFormula is not
    // modelled; the CLI has no verb that writes one)
    val children = col.totalsRowFormula.toList.map(formula =>
      scala.xml.Elem(
        null,
        "totalsRowFormula",
        scala.xml.Null,
        scala.xml.TopScope,
        minimizeEmpty = false,
        scala.xml.Text(formula)
      )
    )
    scala.xml.Elem(
      null,
      "tableColumn",
      finalAttrs,
      scala.xml.TopScope,
      minimizeEmpty = children.isEmpty,
      children*
    )

  /**
   * Serialize table style info to XML.
   *
   * @param info
   *   The style info to serialize
   * @return
   *   XML element
   */
  private def encodeStyleInfo(info: OoxmlTableStyleInfo): Elem =
    // Build attributes in reverse order (linked list serializes backwards)
    // Excel expects: name, showFirstColumn, showLastColumn, showRowStripes, showColumnStripes
    // Build order: showColumnStripes, showRowStripes, showLastColumn, showFirstColumn, name

    val attrs1 = info.otherAttrs.foldLeft(scala.xml.Null: scala.xml.MetaData) {
      case (acc, (k, v)) =>
        new scala.xml.UnprefixedAttribute(k, v, acc)
    }

    val attrs2 = new scala.xml.UnprefixedAttribute(
      "showColumnStripes",
      if info.showColumnStripes then "1" else "0",
      attrs1
    )
    val attrs3 = new scala.xml.UnprefixedAttribute(
      "showRowStripes",
      if info.showRowStripes then "1" else "0",
      attrs2
    )
    val attrs4 = new scala.xml.UnprefixedAttribute(
      "showLastColumn",
      if info.showLastColumn then "1" else "0",
      attrs3
    )
    val attrs5 = new scala.xml.UnprefixedAttribute(
      "showFirstColumn",
      if info.showFirstColumn then "1" else "0",
      attrs4
    )
    val finalAttrs = new scala.xml.UnprefixedAttribute("name", info.name, attrs5)

    scala.xml.Elem(null, "tableStyleInfo", finalAttrs, scala.xml.TopScope, minimizeEmpty = true)

/**
 * Conversion between domain TableSpec and OOXML representation.
 */
object TableConversions:

  /**
   * Convert domain TableSpec to OOXML representation.
   *
   * The revision uids (`xr:uid` on the table and its autoFilter, `xr3:uid` on each column) are
   * optional metadata Excel stamps on its own saves; the domain model does not carry them. They
   * ride through from `source` — the file's part for this table, columns matched by name — and are
   * OMITTED without one (openpyxl and LibreOffice write none; Excel accepts and re-stamps). A fresh
   * random UUID per write made two writes of one workbook differ in every table part (GH-595).
   *
   * The totals row is written the way Excel writes it: `totalsRowCount="1"` on the table, the
   * columns' `totalsRowLabel` / `totalsRowFunction` (+ `<totalsRowFormula>` for `custom`), and the
   * autoFilter over the header and data rows ONLY — Excel excludes the totals row from the filter
   * range (`TableSpec.dataRange` already does).
   *
   * @param spec
   *   Domain table specification
   * @param id
   *   Table ID (1-indexed, unique within workbook)
   * @param source
   *   The source file's OOXML table of the same name, if the write has one
   * @return
   *   OOXML table
   */
  def toOoxml(spec: TableSpec, id: Long, source: Option[OoxmlTable]): OoxmlTable =
    val autoFilterEnabled = spec.autoFilter.exists(_.enabled)
    val tableUid = source.flatMap(_.tableUid)
    val autoFilterUid = if autoFilterEnabled then source.flatMap(_.autoFilterUid) else None

    val columns = spec.columns.map { col =>
      OoxmlTableColumn(
        id = col.id,
        name = col.name,
        uid = source.flatMap(_.columns.find(_.name == col.name)).flatMap(_.uid),
        totalsRowLabel = col.totalsRowLabel,
        totalsRowFunction = col.totalsRowFunction.map(TotalsRowFunction.token),
        totalsRowFormula = col.totalsRowFunction.collect { case TotalsRowFunction.Custom(f) => f }
      )
    }

    val headerCount = if spec.showHeaderRow then 1 else 0
    val totalsCount = if spec.showTotalsRow then 1 else 0

    // header + data rows: the totals row is outside the filter range (a one-row table cannot
    // shrink; CellRange would swap the ends)
    val filterRange =
      if totalsCount > 0 && spec.range.height > totalsCount then
        CellRange(spec.range.start, ARef(spec.range.end.col, spec.range.end.row - totalsCount))
      else spec.range
    val autoFilterRange = Option.when(autoFilterEnabled)(filterRange)

    val styleInfo = styleToStyleInfo(spec.style)

    OoxmlTable(
      id = id,
      name = spec.name,
      displayName = spec.displayName,
      ref = spec.range,
      headerRowCount = headerCount,
      totalsRowCount = totalsCount,
      // "ever shown": now, or per the source part (a table whose totals Excel hid keeps its flag)
      totalsRowShown = spec.showTotalsRow || source.exists(_.totalsRowShown),
      columns = columns,
      autoFilter = autoFilterRange,
      autoFilterUid = autoFilterUid,
      styleInfo = Some(styleInfo),
      tableUid = tableUid
    )

  /**
   * [[toOoxml]] without a source part — an overload rather than a default argument on the
   * three-parameter form, so the two-parameter JVM method consumers compiled against xl-ooxml
   * 0.22.x keeps its erasure (`toOoxml(TableSpec, long)`; a defaulted parameter replaces it with
   * the three-parameter descriptor and a `NoSuchMethodError` at link time).
   */
  def toOoxml(spec: TableSpec, id: Long): OoxmlTable = toOoxml(spec, id, None)

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
      DomainTableColumn(
        col.id,
        col.name,
        totalsRowLabel = col.totalsRowLabel,
        totalsRowFunction =
          col.totalsRowFunction.flatMap(TotalsRowFunction.fromToken(_, col.totalsRowFormula))
      )
    }

    val autoFilter = ooxml.autoFilter.map(_ => TableAutoFilter(enabled = true))

    val style = ooxml.styleInfo.map(styleInfoToStyle).getOrElse(TableStyle.default)

    TableSpec(
      name = ooxml.name,
      displayName = ooxml.displayName,
      range = ooxml.ref,
      columns = columns,
      showHeaderRow = ooxml.headerRowCount > 0,
      // the count says whether a totals row is shown; totalsRowShown only that one ever was
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
