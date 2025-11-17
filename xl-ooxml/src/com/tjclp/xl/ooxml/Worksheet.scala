package com.tjclp.xl.ooxml

import scala.xml.*
import XmlUtil.*
import com.tjclp.xl.addressing.* // For ARef, Column, Row types and extension methods
import com.tjclp.xl.cell.{Cell, CellValue}
import com.tjclp.xl.sheet.Sheet

/** Cell data for worksheet - maps domain Cell to XML representation */
case class OoxmlCell(
  ref: ARef,
  value: CellValue,
  styleIndex: Option[Int] = None,
  cellType: String = "inlineStr" // "s" for SST, "inlineStr" for inline, "n" for number, etc.
):
  def toA1: String = ref.toA1

  def toXml: Elem =
    val baseAttrs = Seq("r" -> toA1)
    val typeAttr = if cellType.nonEmpty then Seq("t" -> cellType) else Seq.empty
    val styleAttr = styleIndex.map(s => Seq("s" -> s.toString)).getOrElse(Seq.empty)
    val attrs = baseAttrs ++ typeAttr ++ styleAttr

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
          val rPrElems = run.font.map { f =>
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

    elem("c", attrs*)(valueElem*)

/** Row in worksheet */
case class OoxmlRow(
  rowIndex: Int, // 1-based
  cells: Seq[OoxmlCell]
):
  def toXml: Elem =
    elem("row", "r" -> rowIndex.toString)(
      cells.sortBy(_.ref.col.index0).map(_.toXml)*
    )

/**
 * Worksheet for xl/worksheets/sheet#.xml
 *
 * Contains the actual cell data in <sheetData> and merged cell ranges in <mergeCells>.
 */
case class OoxmlWorksheet(
  rows: Seq[OoxmlRow],
  mergedRanges: Set[CellRange] = Set.empty
) extends XmlWritable:

  def toXml: Elem =
    val sheetDataElem = elem("sheetData")(
      rows.sortBy(_.rowIndex).map(_.toXml)*
    )

    // Add mergeCells element if there are merged ranges
    val mergeCellsElem = if mergedRanges.nonEmpty then
      // Sort by (row, col) for deterministic, natural ordering (A1, A2, ..., A10, ..., B1, ...)
      // This gives row-major order and avoids lexicographic issues with string sort
      val mergeCellElems = mergedRanges.toSeq
        .sortBy(r => (r.start.row.index0, r.start.col.index0))
        .map { range =>
          elem("mergeCell", "ref" -> range.toA1)()
        }
      Some(elem("mergeCells", "count" -> mergedRanges.size.toString)(mergeCellElems*))
    else None

    val children = sheetDataElem +: mergeCellsElem.toList

    Elem(
      null,
      "worksheet",
      new UnprefixedAttribute(
        "xmlns",
        nsSpreadsheetML,
        new PrefixedAttribute("xmlns", "r", nsRelationships, Null)
      ),
      TopScope,
      minimizeEmpty = false,
      children*
    )

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
    // Group cells by row
    val cellsByRow = sheet.cells.values
      .groupBy(_.ref.row.index1) // 1-based row index
      .toSeq
      .sortBy(_._1)

    val rows = cellsByRow.map { case (rowIdx, cells) =>
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
          case com.tjclp.xl.cell.CellValue.RichText(_) =>
            // Rich text is always inline (cannot be shared)
            ("inlineStr", cell.value)
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
      OoxmlRow(rowIdx, ooxmlCells)
    }

    OoxmlWorksheet(rows, sheet.mergedRanges)

  /** Parse worksheet from XML (XmlReadable trait compatibility) */
  def fromXml(elem: Elem): Either[String, OoxmlWorksheet] =
    fromXmlWithSST(elem, None)

  /** Parse worksheet from XML with optional SharedStrings table */
  def fromXmlWithSST(elem: Elem, sst: Option[SharedStrings]): Either[String, OoxmlWorksheet] =
    for
      sheetDataElem <- getChild(elem, "sheetData")
      rowElems = getChildren(sheetDataElem, "row")
      rows <- parseRows(rowElems, sst)
      mergedRanges <- parseMergeCells(elem)
    yield OoxmlWorksheet(rows, mergedRanges)

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
        cellElems = getChildren(e, "c")
        cells <- parseCells(cellElems, sst)
      yield OoxmlRow(rowIdx, cells)
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
            // Fallback: "str" type may have text in <v> element
            (elem \ "v").headOption.map(_.text) match
              case Some(text) => Right(CellValue.Text(text))
              case None => Left(s"$cellType cell missing <is> element and <v> element")
          case Some(isElem: Elem) =>
            val rElems = getChildren(isElem, "r")

            if rElems.nonEmpty then
              // Rich text: parse runs with formatting
              parseTextRuns(rElems).map(CellValue.RichText.apply)
            else
              // Simple text: extract from <t>
              (isElem \ "t").headOption.map(_.text) match
                case Some(text) => Right(CellValue.Text(text))
                case None => Left(s"$cellType <is> missing <t> element and has no <r> runs")
          case _ => Left(s"$cellType <is> is not an Elem")

      case "s" =>
        // SST index - resolve using SharedStrings table
        (elem \ "v").headOption.map(_.text) match
          case Some(idxStr) =>
            idxStr.toIntOption match
              case Some(idx) =>
                sst.flatMap(_.apply(idx)) match
                  case Some(entry) => Right(sst.get.toCellValue(entry))
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
