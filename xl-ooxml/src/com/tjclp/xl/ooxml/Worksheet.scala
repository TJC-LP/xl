package com.tjclp.xl.ooxml

import scala.xml.*
import XmlUtil.*
import com.tjclp.xl.addressing.* // For ARef, Column, Row types and extension methods
import com.tjclp.xl.cell.{Cell, CellValue}
import com.tjclp.xl.sheet.Sheet

// Default namespaces for generated worksheets. Real files capture the original scope/attributes to
// avoid redundant declarations and preserve mc/x14/xr bindings from the source sheet.
private val defaultWorksheetScope =
  NamespaceBinding(null, nsSpreadsheetML, NamespaceBinding("r", nsRelationships, TopScope))

private def cleanNamespaces(elem: Elem): Elem =
  val cleanedChildren = elem.child.map {
    case e: Elem => cleanNamespaces(e)
    case other => other
  }
  elem.copy(scope = TopScope, child = cleanedChildren)

/** Cell data for worksheet - maps domain Cell to XML representation */
case class OoxmlCell(
  ref: ARef,
  value: CellValue,
  styleIndex: Option[Int] = None,
  cellType: String = "inlineStr" // "s" for SST, "inlineStr" for inline, "n" for number, etc.
):
  def toA1: String = ref.toA1

  def toXml: Elem =
    // Excel expects attributes in specific order: r, s, t
    val attrs = Seq.newBuilder[(String, String)]
    attrs += ("r" -> toA1)
    styleIndex.foreach(s => attrs += ("s" -> s.toString))
    if cellType.nonEmpty then attrs += ("t" -> cellType)

    val finalAttrs = attrs.result()

    val valueElem = value match
      case CellValue.Empty => Seq.empty
      case CellValue.Text(text) if cellType == "inlineStr" =>
        // Add xml:space="preserve" for text with leading/trailing/multiple spaces
        val needsPreserve = needsXmlSpacePreserve(text)
        val tElem =
          if needsPreserve then
            Elem(
              null,
              "t",
              PrefixedAttribute("xml", "space", "preserve", Null),
              TopScope,
              true,
              Text(text)
            )
          else elem("t")(Text(text))
        Seq(elem("is")(tElem))
      case CellValue.Text(text) => // SST index
        Seq(elem("v")(Text(text))) // text here would be the SST index as string
      case CellValue.RichText(richText) =>
        // Rich text: <is> with multiple <r> (text run) elements
        val runElems = richText.runs.map { run =>
          // Use preserved raw <rPr> if available (byte-perfect), otherwise build from Font
          val rPrElems = run.rawRPrXml.flatMap { xmlString =>
            // Parse preserved XML string back to Elem with XXE protection
            XmlSecurity.parseSafe(xmlString, "worksheet richtext rPr").toOption.map { elem =>
              // Strip redundant xmlns recursively from entire tree (namespace already on parent)
              XmlUtil.stripNamespaces(elem)
            }
          }.toList match
            case preserved if preserved.nonEmpty => preserved
            case _ =>
              // Build from Font model if no raw XML or parse failed
              run.font.map { f =>
                val fontProps = Seq.newBuilder[Elem]

                // Font style properties (order matters for OOXML)
                if f.bold then fontProps += elem("b")()
                if f.italic then fontProps += elem("i")()
                if f.underline then fontProps += elem("u")()

                // Font color
                f.color.foreach { c =>
                  fontProps += elem("color", "rgb" -> c.toHex.drop(1))() // Attributes then children
                }

                // Font size and name
                fontProps += elem("sz", "val" -> f.sizePt.toString)()
                fontProps += elem("name", "val" -> f.name)()

                elem("rPr")(fontProps.result()*)
              }.toList

          // Text run: <r> with optional <rPr> and <t>
          // Add xml:space="preserve" to preserve leading/trailing/multiple spaces
          val textElem =
            if needsXmlSpacePreserve(run.text) then
              Elem(
                null,
                "t",
                PrefixedAttribute("xml", "space", "preserve", Null),
                TopScope,
                true,
                Text(run.text)
              )
            else elem("t")(Text(run.text))

          elem("r")(
            rPrElems ++ Seq(textElem)*
          )
        }

        Seq(elem("is")(runElems*))
      case CellValue.Number(num) =>
        Seq(elem("v")(Text(num.toString)))
      case CellValue.Bool(b) =>
        Seq(elem("v")(Text(if b then "1" else "0")))
      case CellValue.Formula(expr) =>
        Seq(elem("f")(Text(expr))) // Simplified - full formula support later
      case CellValue.Error(err) =>
        import com.tjclp.xl.cell.CellError.toExcel
        Seq(elem("v")(Text(err.toExcel)))
      case CellValue.DateTime(dt) =>
        // DateTime is serialized as number with Excel serial format
        val serial = CellValue.dateTimeToExcelSerial(dt)
        Seq(elem("v")(Text(serial.toString)))

    XmlUtil.elemOrdered("c", finalAttrs*)(valueElem*)

/**
 * Row in worksheet with full attribute preservation for surgical modification.
 *
 * Preserves all OOXML row attributes to maintain byte-level fidelity during regeneration.
 */
case class OoxmlRow(
  rowIndex: Int, // 1-based
  cells: Seq[OoxmlCell],
  // Row-level attributes (all optional)
  spans: Option[String] = None, // "2:16" (cell coverage optimization hint)
  style: Option[Int] = None, // s="7" (row-level style ID)
  height: Option[Double] = None, // ht="24.95" (custom row height in points)
  customHeight: Boolean = false, // customHeight="1"
  customFormat: Boolean = false, // customFormat="1"
  hidden: Boolean = false, // hidden="1"
  outlineLevel: Option[Int] = None, // outlineLevel="1" (grouping level)
  collapsed: Boolean = false, // collapsed="1" (outline collapsed)
  thickBot: Boolean = false, // thickBot="1" (thick bottom border)
  thickTop: Boolean = false, // thickTop="1" (thick top border)
  dyDescent: Option[Double] = None // x14ac:dyDescent="0.25" (font descent adjustment)
):
  def toXml: Elem =
    // Excel expects attributes in specific order (not alphabetical!)
    // Order: r, spans, s, customFormat, ht, customHeight, hidden, outlineLevel, collapsed, thickBot, thickTop, x14ac:dyDescent
    val attrs = Seq.newBuilder[(String, String)]

    attrs += ("r" -> rowIndex.toString)
    spans.foreach(s => attrs += ("spans" -> s))
    style.foreach(s => attrs += ("s" -> s.toString))
    if customFormat then attrs += ("customFormat" -> "1")
    height.foreach(h => attrs += ("ht" -> h.toString))
    if customHeight then attrs += ("customHeight" -> "1")
    if hidden then attrs += ("hidden" -> "1")
    outlineLevel.foreach(l => attrs += ("outlineLevel" -> l.toString))
    if collapsed then attrs += ("collapsed" -> "1")
    if thickBot then attrs += ("thickBot" -> "1")
    if thickTop then attrs += ("thickTop" -> "1")
    dyDescent.foreach(d => attrs += ("x14ac:dyDescent" -> d.toString))

    XmlUtil.elemOrdered("row", attrs.result()*)(
      cells.sortBy(_.ref.col.index0).map(_.toXml)*
    )

/**
 * Worksheet for xl/worksheets/sheet#.xml
 *
 * Contains cell data in <sheetData> and preserves all worksheet metadata for surgical modification.
 * Unparsed elements are preserved as raw XML to maintain Excel compatibility.
 *
 * Elements are emitted in OOXML Part 1 schema order (§18.3.1.99).
 *
 * @param rows
 *   Row data with cells
 * @param mergedRanges
 *   Merged cell ranges
 * @param sheetPr
 *   Sheet properties (pageSetUpPr, outlinePr, tabColor, etc.)
 * @param dimension
 *   Used range reference (e.g., ref="B1:U104")
 * @param sheetViews
 *   View settings (zoom, gridLines, selection, tabSelected, pane, etc.)
 * @param sheetFormatPr
 *   Default row/column sizes and outline levels
 * @param cols
 *   Column definitions (CRITICAL - widths, styles, hidden, outlineLevel, etc.)
 * @param pageMargins
 *   Page margins (left, right, top, bottom, header, footer)
 * @param pageSetup
 *   Page setup (orientation, paperSize, scale, fitToWidth, fitToHeight, etc.)
 * @param headerFooter
 *   Header and footer content
 * @param drawing
 *   Drawing reference for charts (r:id link to drawing.xml)
 * @param legacyDrawing
 *   Legacy VML drawing reference for comments (r:id link to vmlDrawing.vml)
 * @param picture
 *   Background picture reference
 * @param oleObjects
 *   OLE objects
 * @param controls
 *   ActiveX controls
 * @param extLst
 *   Extensions
 * @param otherElements
 *   Any other elements not explicitly handled
 */
case class OoxmlWorksheet(
  rows: Seq[OoxmlRow],
  mergedRanges: Set[CellRange] = Set.empty,
  // Worksheet metadata (emitted in OOXML schema order)
  sheetPr: Option[Elem] = None,
  dimension: Option[Elem] = None,
  sheetViews: Option[Elem] = None,
  sheetFormatPr: Option[Elem] = None,
  cols: Option[Elem] = None,
  conditionalFormatting: Seq[Elem] = Seq.empty,
  printOptions: Option[Elem] = None,
  rowBreaks: Option[Elem] = None,
  colBreaks: Option[Elem] = None,
  customPropertiesWs: Option[Elem] = None,
  // Page layout
  pageMargins: Option[Elem] = None,
  pageSetup: Option[Elem] = None,
  headerFooter: Option[Elem] = None,
  // Drawings and objects
  drawing: Option[Elem] = None,
  legacyDrawing: Option[Elem] = None,
  picture: Option[Elem] = None,
  oleObjects: Option[Elem] = None,
  controls: Option[Elem] = None,
  // Extensions
  extLst: Option[Elem] = None,
  otherElements: Seq[Elem] = Seq.empty,
  rootAttributes: MetaData = Null,
  rootScope: NamespaceBinding = defaultWorksheetScope
) extends XmlWritable:

  def toXml: Elem =
    val children = Seq.newBuilder[Node]

    // Emit in OOXML Part 1 schema order (§18.3.1.99) - critical for Excel compatibility
    sheetPr.foreach(e => children += cleanNamespaces(e))
    dimension.foreach(e => children += cleanNamespaces(e))
    sheetViews.foreach(e => children += cleanNamespaces(e))
    sheetFormatPr.foreach(e => children += cleanNamespaces(e))
    cols.foreach(e => children += cleanNamespaces(e)) // Column definitions BEFORE sheetData

    // sheetData (always regenerated with current cell values)
    val sheetDataElem = elem("sheetData")(
      rows.sortBy(_.rowIndex).map(_.toXml)*
    )
    children += sheetDataElem

    // mergeCells (regenerated if present)
    if mergedRanges.nonEmpty then
      val mergeCellElems = mergedRanges.toSeq
        .sortBy(r => (r.start.row.index0, r.start.col.index0))
        .map(range => elem("mergeCell", "ref" -> range.toA1)())
      children += elem("mergeCells", "count" -> mergedRanges.size.toString)(mergeCellElems*)

    // Conditional formatting (multiple allowed)
    conditionalFormatting.foreach(e => children += cleanNamespaces(e))

    // Page layout (after sheetData/mergeCells)
    printOptions.foreach(e => children += cleanNamespaces(e))
    pageMargins.foreach(e => children += cleanNamespaces(e))
    pageSetup.foreach(e => children += cleanNamespaces(e))
    headerFooter.foreach(e => children += cleanNamespaces(e))

    rowBreaks.foreach(e => children += cleanNamespaces(e))
    colBreaks.foreach(e => children += cleanNamespaces(e))
    customPropertiesWs.foreach(e => children += cleanNamespaces(e))

    // Drawings and objects
    drawing.foreach(e => children += cleanNamespaces(e))
    legacyDrawing.foreach(e => children += cleanNamespaces(e))
    picture.foreach(e => children += cleanNamespaces(e))
    oleObjects.foreach(e => children += cleanNamespaces(e))
    controls.foreach(e => children += cleanNamespaces(e))

    // Extensions
    extLst.foreach(e => children += cleanNamespaces(e))

    // Any other elements
    otherElements.foreach(e => children += cleanNamespaces(e))

    val scope = Option(rootScope).getOrElse(defaultWorksheetScope)
    val attrs = Option(rootAttributes).getOrElse(Null)

    Elem(null, "worksheet", attrs, scope, minimizeEmpty = false, children.result()*)

object OoxmlWorksheet extends XmlReadable[OoxmlWorksheet]:
  /** Create minimal empty worksheet */
  def empty: OoxmlWorksheet = OoxmlWorksheet(Seq.empty)

  /** Create worksheet from domain Sheet (inline strings only) */
  def fromDomain(sheet: Sheet, styleRemapping: Map[Int, Int] = Map.empty): OoxmlWorksheet =
    fromDomainWithSST(sheet, None, styleRemapping)

  /**
   * Create worksheet from domain Sheet with optional SST and style remapping.
   *
   * @param sheet
   *   The domain Sheet to serialize
   * @param sst
   *   Optional SharedStrings table for string deduplication
   * @param styleRemapping
   *   Map from sheet-local styleId to workbook-level styleId
   */
  def fromDomainWithSST(
    sheet: Sheet,
    sst: Option[SharedStrings],
    styleRemapping: Map[Int, Int] = Map.empty
  ): OoxmlWorksheet =
    fromDomainWithMetadata(sheet, sst, styleRemapping, None)

  /**
   * Create worksheet from domain Sheet, preserving metadata from original worksheet XML.
   *
   * Used during surgical modification to regenerate sheetData while preserving all worksheet
   * metadata (column widths, view settings, page setup, etc.).
   *
   * @param sheet
   *   The domain Sheet with updated cell values
   * @param sst
   *   Optional SharedStrings table
   * @param styleRemapping
   *   Map from sheet-local styleId to workbook-level styleId
   * @param preservedMetadata
   *   Optional original worksheet to extract metadata from
   */
  def fromDomainWithMetadata(
    sheet: Sheet,
    sst: Option[SharedStrings],
    styleRemapping: Map[Int, Int] = Map.empty,
    preservedMetadata: Option[OoxmlWorksheet] = None
  ): OoxmlWorksheet =
    // Build a map of row indices to preserved row attributes
    val preservedRowAttrs = preservedMetadata
      .map { preserved =>
        preserved.rows.map(r => r.rowIndex -> r).toMap
      }
      .getOrElse(Map.empty)

    // Group cells by row
    val cellsByRow = sheet.cells.values
      .groupBy(_.ref.row.index1) // 1-based row index
      .toSeq
      .sortBy(_._1)

    // Create rows with cells (preserving attributes from original)
    val rowsWithCells = cellsByRow.map { case (rowIdx, cells) =>
      val ooxmlCells = cells.map { cell =>
        // Remap sheet-local styleId to workbook-level index
        val globalStyleIdx = cell.styleId.flatMap { localId =>
          // Look up in remapping table, fall back to 0 (default) if not found
          styleRemapping.get(localId.value).orElse(Some(0))
        }

        // Determine cell type and value based on CellValue type and SST availability
        val (cellType, value) = cell.value match
          case com.tjclp.xl.cell.CellValue.Text(s) =>
            sst.flatMap(_.indexOf(s)) match
              case Some(idx) => ("s", com.tjclp.xl.cell.CellValue.Text(idx.toString))
              case None => ("inlineStr", cell.value)
          case com.tjclp.xl.cell.CellValue.RichText(rt) =>
            // Check if RichText exists in SST (it can be shared!)
            sst.flatMap(_.indexOf(rt)) match
              case Some(idx) => ("s", com.tjclp.xl.cell.CellValue.Text(idx.toString))
              case None => ("inlineStr", cell.value)
          case com.tjclp.xl.cell.CellValue.Number(_) => ("n", cell.value)
          case com.tjclp.xl.cell.CellValue.Bool(_) => ("b", cell.value)
          case com.tjclp.xl.cell.CellValue.DateTime(dt) =>
            // Convert to Excel serial number
            val serial = com.tjclp.xl.cell.CellValue.dateTimeToExcelSerial(dt)
            ("n", com.tjclp.xl.cell.CellValue.Number(BigDecimal(serial)))
          case com.tjclp.xl.cell.CellValue.Formula(_) => ("str", cell.value) // Formula result
          case com.tjclp.xl.cell.CellValue.Error(_) => ("e", cell.value)
          case com.tjclp.xl.cell.CellValue.Empty => ("", cell.value)

        OoxmlCell(cell.ref, value, globalStyleIdx, cellType)
      }.toSeq

      // Preserve row attributes from original if available
      val baseRow = preservedRowAttrs.get(rowIdx) match
        case Some(original) =>
          // Merge: use original's attributes, replace cells with new data
          original.copy(cells = ooxmlCells)
        case None =>
          // New row - create with defaults
          OoxmlRow(rowIdx, ooxmlCells)

      // Preserve row-level style exactly as-is from original (even if "invalid" per spec)
      // Excel expects these preserved, removing them causes corruption warnings
      baseRow
    }

    // Preserve empty rows from original (critical for Row 1!)
    val emptyRowsFromOriginal = preservedMetadata.toList.flatMap { preserved =>
      preserved.rows.filter(_.cells.isEmpty)
    }

    // Combine rows with cells + empty rows, sort by index
    val allRows = (rowsWithCells ++ emptyRowsFromOriginal).sortBy(_.rowIndex)

    // If preservedMetadata is provided, use its metadata fields; otherwise use defaults (None)
    preservedMetadata match
      case Some(preserved) =>
        OoxmlWorksheet(
          allRows, // Use merged rows (with cells + empty rows)
          sheet.mergedRanges,
          // Preserve all metadata from original
          preserved.sheetPr,
          preserved.dimension,
          preserved.sheetViews,
          preserved.sheetFormatPr,
          preserved.cols,
          preserved.conditionalFormatting,
          preserved.printOptions,
          preserved.rowBreaks,
          preserved.colBreaks,
          preserved.customPropertiesWs,
          preserved.pageMargins,
          preserved.pageSetup,
          preserved.headerFooter,
          preserved.drawing,
          preserved.legacyDrawing,
          preserved.picture,
          preserved.oleObjects,
          preserved.controls,
          preserved.extLst,
          preserved.otherElements,
          preserved.rootAttributes,
          preserved.rootScope
        )
      case None =>
        // No preserved metadata - create minimal worksheet
        OoxmlWorksheet(rowsWithCells, sheet.mergedRanges)

  /** Parse worksheet from XML (XmlReadable trait compatibility) */
  def fromXml(elem: Elem): Either[String, OoxmlWorksheet] =
    fromXmlWithSST(elem, None)

  /** Parse worksheet from XML with optional SharedStrings table */
  def fromXmlWithSST(elem: Elem, sst: Option[SharedStrings]): Either[String, OoxmlWorksheet] =
    for
      // Parse sheetData (required)
      sheetDataElem <- getChild(elem, "sheetData")
      rowElems = getChildren(sheetDataElem, "row")
      rows <- parseRows(rowElems, sst)

      // Parse mergeCells (optional)
      mergedRanges <- parseMergeCells(elem)

      // Extract ALL preserved metadata elements (all optional)
      sheetPr = (elem \ "sheetPr").headOption.collect { case e: Elem => cleanNamespaces(e) }
      dimension = (elem \ "dimension").headOption.collect { case e: Elem => cleanNamespaces(e) }
      sheetViews = (elem \ "sheetViews").headOption.collect { case e: Elem => cleanNamespaces(e) }
      sheetFormatPr = (elem \ "sheetFormatPr").headOption.collect { case e: Elem =>
        cleanNamespaces(e)
      }
      cols = (elem \ "cols").headOption.collect { case e: Elem => cleanNamespaces(e) }

      conditionalFormatting = elem.child.collect {
        case e: Elem if e.label == "conditionalFormatting" => cleanNamespaces(e)
      }
      printOptions = (elem \ "printOptions").headOption.collect { case e: Elem =>
        cleanNamespaces(e)
      }
      rowBreaks = (elem \ "rowBreaks").headOption.collect { case e: Elem => cleanNamespaces(e) }
      colBreaks = (elem \ "colBreaks").headOption.collect { case e: Elem => cleanNamespaces(e) }
      customPropertiesWs = (elem \ "customProperties").headOption.collect { case e: Elem =>
        cleanNamespaces(e)
      }

      pageMargins = (elem \ "pageMargins").headOption.collect { case e: Elem => cleanNamespaces(e) }
      pageSetup = (elem \ "pageSetup").headOption.collect { case e: Elem => cleanNamespaces(e) }
      headerFooter = (elem \ "headerFooter").headOption.collect { case e: Elem =>
        cleanNamespaces(e)
      }

      drawing = (elem \ "drawing").headOption.collect { case e: Elem => cleanNamespaces(e) }
      legacyDrawing = (elem \ "legacyDrawing").headOption.collect { case e: Elem =>
        cleanNamespaces(e)
      }
      picture = (elem \ "picture").headOption.collect { case e: Elem => cleanNamespaces(e) }
      oleObjects = (elem \ "oleObjects").headOption.collect { case e: Elem => cleanNamespaces(e) }
      controls = (elem \ "controls").headOption.collect { case e: Elem => cleanNamespaces(e) }

      extLst = (elem \ "extLst").headOption.collect { case e: Elem => cleanNamespaces(e) }

      // Collect any other elements we don't explicitly handle
      knownElements = Set(
        "sheetPr",
        "dimension",
        "sheetViews",
        "sheetFormatPr",
        "cols",
        "sheetData",
        "mergeCells",
        "pageMargins",
        "pageSetup",
        "headerFooter",
        "drawing",
        "legacyDrawing",
        "picture",
        "oleObjects",
        "controls",
        "extLst",
        // Additional elements from OOXML spec
        "sheetCalcPr",
        "sheetProtection",
        "protectedRanges",
        "scenarios",
        "autoFilter",
        "sortState",
        "dataConsolidate",
        "customSheetViews",
        "phoneticPr",
        "conditionalFormatting",
        "dataValidations",
        "hyperlinks",
        "printOptions",
        "rowBreaks",
        "colBreaks",
        "customProperties",
        "cellWatches",
        "ignoredErrors",
        "smartTags",
        "legacyDrawingHF",
        "webPublishItems",
        "tableParts"
      )
      otherElements = elem.child.collect {
        case e: Elem if !knownElements.contains(e.label) => e
      }
    yield OoxmlWorksheet(
      rows,
      mergedRanges,
      sheetPr,
      dimension,
      sheetViews,
      sheetFormatPr,
      cols,
      conditionalFormatting,
      printOptions,
      rowBreaks,
      colBreaks,
      customPropertiesWs,
      pageMargins,
      pageSetup,
      headerFooter,
      drawing,
      legacyDrawing,
      picture,
      oleObjects,
      controls,
      extLst,
      otherElements.toSeq,
      rootAttributes = elem.attributes,
      rootScope = Option(elem.scope).getOrElse(defaultWorksheetScope)
    )

  private def parseMergeCells(worksheetElem: Elem): Either[String, Set[CellRange]] =
    // mergeCells is optional
    (worksheetElem \ "mergeCells").headOption match
      case None => Right(Set.empty)
      case Some(mergeCellsElem: Elem) =>
        val mergeCellElems = getChildren(mergeCellsElem, "mergeCell")
        val parsed = mergeCellElems.map { elem =>
          for
            refStr <- getAttr(elem, "ref")
            range <- CellRange.parse(refStr)
          yield range
        }
        val errors = parsed.collect { case Left(err) => err }
        if errors.nonEmpty then Left(s"MergeCell parse errors: ${errors.mkString(", ")}")
        else Right(parsed.collect { case Right(range) => range }.toSet)
      case _ => Right(Set.empty) // Non-Elem node, ignore

  private def parseRows(
    elems: Seq[Elem],
    sst: Option[SharedStrings]
  ): Either[String, Seq[OoxmlRow]] =
    val parsed = elems.map { e =>
      for
        rStr <- getAttr(e, "r")
        rowIdx <- rStr.toIntOption.toRight(s"Invalid row index: $rStr")

        // Extract ALL row attributes for byte-perfect preservation
        spans = getAttrOpt(e, "spans")
        style = getAttrOpt(e, "s").flatMap(_.toIntOption)
        height = getAttrOpt(e, "ht").flatMap(_.toDoubleOption)
        customHeight = getAttrOpt(e, "customHeight").contains("1")
        customFormat = getAttrOpt(e, "customFormat").contains("1")
        hidden = getAttrOpt(e, "hidden").contains("1")
        outlineLevel = getAttrOpt(e, "outlineLevel").flatMap(_.toIntOption)
        collapsed = getAttrOpt(e, "collapsed").contains("1")
        thickBot = getAttrOpt(e, "thickBot").contains("1")
        thickTop = getAttrOpt(e, "thickTop").contains("1")
        dyDescent = XmlUtil.getNamespacedAttrOpt(e, "x14ac:dyDescent").flatMap(_.toDoubleOption)

        cellElems = getChildren(e, "c")
        cells <- parseCells(cellElems, sst)
      yield OoxmlRow(
        rowIdx,
        cells,
        spans,
        style,
        height,
        customHeight,
        customFormat,
        hidden,
        outlineLevel,
        collapsed,
        thickBot,
        thickTop,
        dyDescent
      )
    }

    val errors = parsed.collect { case Left(err) => err }
    if errors.nonEmpty then Left(s"Row parse errors: ${errors.mkString(", ")}")
    else Right(parsed.collect { case Right(row) => row })

  private def parseCells(
    elems: Seq[Elem],
    sst: Option[SharedStrings]
  ): Either[String, Seq[OoxmlCell]] =
    val parsed = elems.map { e =>
      for
        refStr <- getAttr(e, "r")
        ref <- ARef.parse(refStr)
        cellType = getAttrOpt(e, "t").getOrElse("")
        styleIdx = getAttrOpt(e, "s").flatMap(_.toIntOption)
        value <- parseCellValue(e, cellType, sst)
      yield OoxmlCell(ref, value, styleIdx, cellType)
    }

    val errors = parsed.collect { case Left(err) => err }
    if errors.nonEmpty then Left(s"Cell parse errors: ${errors.mkString(", ")}")
    else Right(parsed.collect { case Right(cell) => cell })

  private def parseCellValue(
    elem: Elem,
    cellType: String,
    sst: Option[SharedStrings]
  ): Either[String, CellValue] =
    cellType match
      case "inlineStr" | "str" =>
        // Both "inlineStr" and "str" cell types use <is> for inline strings
        // Check for rich text (<is><r>) vs simple text (<is><t>)
        (elem \ "is").headOption match
          case None =>
            // Fallback: "str" type may have text in <v> element (preserving whitespace)
            (elem \ "v").headOption
              .collect { case e: Elem => e }
              .map(getTextPreservingWhitespace) match
              case Some(text) => Right(CellValue.Text(text))
              case None => Left(s"$cellType cell missing <is> element and <v> element")
          case Some(isElem: Elem) =>
            val rElems = getChildren(isElem, "r")

            if rElems.nonEmpty then
              // Rich text: parse runs with formatting
              parseTextRuns(rElems).map(CellValue.RichText.apply)
            else
              // Simple text: extract from <t> (preserving whitespace)
              (isElem \ "t").headOption
                .collect { case e: Elem => e }
                .map(getTextPreservingWhitespace) match
                case Some(text) => Right(CellValue.Text(text))
                case None => Left(s"$cellType <is> missing <t> element and has no <r> runs")
          case _ => Left(s"$cellType <is> is not an Elem")

      case "s" =>
        // SST index - resolve using SharedStrings table
        (elem \ "v").headOption.map(_.text) match
          case Some(idxStr) =>
            idxStr.toIntOption match
              case Some(idx) =>
                (for {
                  sharedStrings <- sst
                  entry <- sharedStrings.apply(idx)
                } yield sharedStrings.toCellValue(entry)) match
                  case Some(cellValue) => Right(cellValue)
                  case None =>
                    // SST index out of bounds → CellError.Ref (not parse failure)
                    Right(CellValue.Error(com.tjclp.xl.cell.CellError.Ref))
              case None =>
                // Invalid SST index format → CellError.Value
                Right(CellValue.Error(com.tjclp.xl.cell.CellError.Value))
          case None => Left("SST cell missing <v>")

      case "n" | "" =>
        // Number
        (elem \ "v").headOption.map(_.text) match
          case Some(numStr) =>
            try Right(CellValue.Number(BigDecimal(numStr)))
            catch case _: NumberFormatException => Left(s"Invalid number: $numStr")
          case None => Right(CellValue.Empty) // Empty numeric cell

      case "b" =>
        // Boolean
        (elem \ "v").headOption.map(_.text) match
          case Some("1") | Some("true") => Right(CellValue.Bool(true))
          case Some("0") | Some("false") => Right(CellValue.Bool(false))
          case other => Left(s"Invalid boolean value: $other")

      case "e" =>
        // Error
        (elem \ "v").headOption.map(_.text) match
          case Some(errStr) =>
            import com.tjclp.xl.cell.CellError
            CellError.parse(errStr).map(CellValue.Error.apply)
          case None => Left("Error cell missing <v>")

      case other =>
        Left(s"Unsupported cell type: $other")
