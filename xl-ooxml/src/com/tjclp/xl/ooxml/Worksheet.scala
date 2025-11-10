package com.tjclp.xl.ooxml

import scala.xml.*
import XmlUtil.*
import com.tjclp.xl.{ARef, Cell, CellValue, Row as CoreRow, Column, Sheet}

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
        Seq(elem("is")(elem("t")(Text(text))))
      case CellValue.Text(text) => // SST index
        Seq(elem("v")(Text(text))) // text here would be the SST index as string
      case CellValue.Number(num) =>
        Seq(elem("v")(Text(num.toString)))
      case CellValue.Bool(b) =>
        Seq(elem("v")(Text(if b then "1" else "0")))
      case CellValue.Formula(expr) =>
        Seq(elem("f")(Text(expr))) // Simplified - full formula support later
      case CellValue.Error(err) =>
        import com.tjclp.xl.CellError.toExcel
        Seq(elem("v")(Text(err.toExcel)))
      case CellValue.DateTime(_) =>
        // DateTime is serialized as number with date format
        Seq(elem("v")(Text("0"))) // TODO: proper date serialization

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
 * Contains the actual cell data in <sheetData>.
 */
case class OoxmlWorksheet(
  rows: Seq[OoxmlRow]
) extends XmlWritable:

  def toXml: Elem =
    val sheetDataElem = elem("sheetData")(
      rows.sortBy(_.rowIndex).map(_.toXml)*
    )

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
      sheetDataElem
    )

object OoxmlWorksheet extends XmlReadable[OoxmlWorksheet]:
  /** Create minimal empty worksheet */
  def empty: OoxmlWorksheet = OoxmlWorksheet(Seq.empty)

  /** Create worksheet from domain Sheet (inline strings only for now) */
  def fromDomain(sheet: Sheet, styleIndexMap: Map[String, Int] = Map.empty): OoxmlWorksheet =
    // Group cells by row
    val cellsByRow = sheet.cells.values
      .groupBy(_.ref.row.index1) // 1-based row index
      .toSeq
      .sortBy(_._1)

    val rows = cellsByRow.map { case (rowIdx, cells) =>
      val ooxmlCells = cells.map { cell =>
        val styleIdx = cell.styleId
        OoxmlCell(cell.ref, cell.value, styleIdx)
      }.toSeq
      OoxmlRow(rowIdx, ooxmlCells)
    }

    OoxmlWorksheet(rows)

  def fromXml(elem: Elem): Either[String, OoxmlWorksheet] =
    for
      sheetDataElem <- getChild(elem, "sheetData")
      rowElems = getChildren(sheetDataElem, "row")
      rows <- parseRows(rowElems)
    yield OoxmlWorksheet(rows)

  private def parseRows(elems: Seq[Elem]): Either[String, Seq[OoxmlRow]] =
    val parsed = elems.map { e =>
      for
        rStr <- getAttr(e, "r")
        rowIdx <- rStr.toIntOption.toRight(s"Invalid row index: $rStr")
        cellElems = getChildren(e, "c")
        cells <- parseCells(cellElems)
      yield OoxmlRow(rowIdx, cells)
    }

    val errors = parsed.collect { case Left(err) => err }
    if errors.nonEmpty then Left(s"Row parse errors: ${errors.mkString(", ")}")
    else Right(parsed.collect { case Right(row) => row })

  private def parseCells(elems: Seq[Elem]): Either[String, Seq[OoxmlCell]] =
    val parsed = elems.map { e =>
      for
        refStr <- getAttr(e, "r")
        ref <- ARef.parse(refStr)
        cellType = getAttrOpt(e, "t").getOrElse("")
        styleIdx = getAttrOpt(e, "s").flatMap(_.toIntOption)
        value <- parseCellValue(e, cellType)
      yield OoxmlCell(ref, value, styleIdx, cellType)
    }

    val errors = parsed.collect { case Left(err) => err }
    if errors.nonEmpty then Left(s"Cell parse errors: ${errors.mkString(", ")}")
    else Right(parsed.collect { case Right(cell) => cell })

  private def parseCellValue(elem: Elem, cellType: String): Either[String, CellValue] =
    cellType match
      case "inlineStr" =>
        // <is><t>text</t></is>
        (elem \ "is" \ "t").headOption.map(_.text) match
          case Some(text) => Right(CellValue.Text(text))
          case None => Left("inlineStr cell missing <is><t>")

      case "s" =>
        // SST index - for now just store as text
        (elem \ "v").headOption.map(_.text) match
          case Some(idx) => Right(CellValue.Text(idx)) // TODO: resolve SST
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
            import com.tjclp.xl.CellError
            CellError.parse(errStr).map(CellValue.Error.apply)
          case None => Left("Error cell missing <v>")

      case other =>
        Left(s"Unsupported cell type: $other")
