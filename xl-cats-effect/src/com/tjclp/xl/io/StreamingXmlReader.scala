package com.tjclp.xl.io

import cats.effect.Sync
import cats.syntax.all.*
import fs2.{Stream, Pipe}
import fs2.data.xml
import fs2.data.xml.XmlEvent
import com.tjclp.xl.{CellValue, CellError}
import com.tjclp.xl.ooxml.SharedStrings

/**
 * True streaming XML reader using fs2-data-xml for constant-memory reads.
 *
 * Parses XML events incrementally without materializing the full document tree.
 */
object StreamingXmlReader:

  /**
   * Parse SharedStrings table from XML byte stream.
   *
   * Note: SST is materialized in memory (acceptable overhead, typically <10MB). This is necessary
   * for resolving string cell references during worksheet parsing.
   */
  def parseSharedStrings[F[_]: Sync](xmlBytes: Stream[F, Byte]): F[Option[SharedStrings]] =
    xmlBytes
      .through(fs2.text.utf8.decode)
      .through(xml.events[F, String]())
      .compile
      .fold(SSTBuilder.empty)((builder, event) => builder.add(event))
      .map { builder =>
        if builder.strings.isEmpty then None
        else
          // For streaming, we don't track totalCount during parsing
          // Use uniqueCount as a safe default (actual count may be higher)
          val uniqueCount = builder.strings.size
          Some(SharedStrings(builder.strings.toVector, builder.buildIndexMap, uniqueCount))
      }

  /**
   * Stream worksheet rows incrementally from XML byte stream.
   *
   * Parses XML events and emits RowData as each row is completed. Memory usage is O(1) regardless
   * of total row count.
   *
   * @param xmlBytes
   *   Worksheet XML as byte stream
   * @param sst
   *   Optional SharedStrings table for resolving string references
   * @return
   *   Stream of RowData (one per row)
   */
  def parseWorksheetStream[F[_]: Sync](
    xmlBytes: Stream[F, Byte],
    sst: Option[SharedStrings]
  ): Stream[F, RowData] =
    xmlBytes
      .through(fs2.text.utf8.decode)
      .through(xml.events[F, String]())
      .through(worksheetEventsToRows(sst))

  /**
   * Convert XML events to RowData using a state machine.
   *
   * Tracks current row index and accumulates cells until </row> event.
   */
  private def worksheetEventsToRows[F[_]](
    sst: Option[SharedStrings]
  ): Pipe[F, XmlEvent, RowData] =
    _.scan(RowBuilder.empty) { (builder, event) =>
      builder.process(event, sst)
    }
      .collect { case RowBuilder.Complete(rowData) =>
        rowData
      }

  // ----- Internal State Machines -----

  /** Builder for accumulating SharedStrings during parsing */
  private case class SSTBuilder(
    strings: Vector[String],
    currentText: Option[String],
    inT: Boolean
  ):
    def add(event: XmlEvent): SSTBuilder = event match
      case XmlEvent.StartTag(xml.QName(_, "t"), _, _) =>
        copy(inT = true, currentText = Some(""))

      case XmlEvent.XmlString(text, _) if inT =>
        copy(currentText = currentText.map(_ + text).orElse(Some(text)))

      case XmlEvent.EndTag(xml.QName(_, "t")) =>
        val newStrings = currentText.fold(strings)(s => strings :+ s)
        copy(strings = newStrings, currentText = None, inT = false)

      case _ => this

    def buildIndexMap: Map[String, Int] =
      strings.zipWithIndex.map { case (s, idx) => s -> idx }.toMap

  private object SSTBuilder:
    val empty: SSTBuilder = SSTBuilder(Vector.empty, None, false)

  /** Builder for accumulating row data during parsing */
  private sealed trait RowBuilder:
    def process(event: XmlEvent, sst: Option[SharedStrings]): RowBuilder

  private object RowBuilder:
    val empty: RowBuilder = Idle

    case object Idle extends RowBuilder:
      def process(event: XmlEvent, sst: Option[SharedStrings]): RowBuilder = event match
        case XmlEvent.StartTag(xml.QName(_, "row"), attrs, _) =>
          // Extract row index from r="1" attribute
          val rowIndex = attrs
            .find(_.name match { case xml.QName(_, "r") => true; case _ => false })
            .flatMap(attr => extractText(attr.value).flatMap(_.toIntOption))
            .getOrElse(1)
          InRow(rowIndex, Map.empty, None)
        case _ => this

    case class InRow(
      rowIndex: Int,
      cells: Map[Int, CellValue],
      currentCell: Option[CellBuilder]
    ) extends RowBuilder:
      def process(event: XmlEvent, sst: Option[SharedStrings]): RowBuilder = event match
        case XmlEvent.StartTag(xml.QName(_, "c"), attrs, _) =>
          // Start new cell
          val cellRef = attrs
            .find(_.name match { case xml.QName(_, "r") => true; case _ => false })
            .flatMap(attr => extractText(attr.value))
          val cellType = attrs
            .find(_.name match { case xml.QName(_, "t") => true; case _ => false })
            .flatMap(attr => extractText(attr.value))
          copy(currentCell = Some(CellBuilder(cellRef, cellType, None)))

        case XmlEvent.EndTag(xml.QName(_, "c")) =>
          // Complete cell and add to row
          val updatedCells = currentCell.flatMap(_.toCellValue(sst)) match
            case Some((colIdx, value)) => cells + (colIdx -> value)
            case None => cells
          copy(currentCell = None, cells = updatedCells)

        case XmlEvent.EndTag(xml.QName(_, "row")) =>
          // Complete row
          Complete(RowData(rowIndex, cells))

        case _ =>
          // Pass event to current cell builder
          currentCell match
            case Some(cell) => copy(currentCell = Some(cell.process(event)))
            case None => this

    case class Complete(rowData: RowData) extends RowBuilder:
      def process(event: XmlEvent, sst: Option[SharedStrings]): RowBuilder = event match
        case XmlEvent.StartTag(xml.QName(_, "row"), attrs, _) =>
          // Start next row
          val rowIndex = attrs
            .find(_.name match { case xml.QName(_, "r") => true; case _ => false })
            .flatMap(attr => extractText(attr.value).flatMap(_.toIntOption))
            .getOrElse(1)
          InRow(rowIndex, Map.empty, None)
        case _ => Idle

  /** Builder for accumulating cell data */
  private case class CellBuilder(
    ref: Option[String], // "A1", "B2", etc.
    cellType: Option[String], // "s", "n", "b", "str", "e", "inlineStr"
    value: Option[String] // Text content from <v> or <t>
  ):
    def process(event: XmlEvent): CellBuilder = event match
      case XmlEvent.StartTag(xml.QName(_, name), _, _) if name == "v" || name == "t" =>
        this

      case XmlEvent.XmlString(text, _) =>
        copy(value = Some(value.getOrElse("") + text))

      case _ => this

    def toCellValue(sst: Option[SharedStrings]): Option[(Int, CellValue)] =
      for
        cellRef <- ref
        colIdx <- parseCellColumn(cellRef)
        cellVal <- value.map(v => interpretCellValue(v, cellType, sst))
      yield (colIdx, cellVal)

    private def parseCellColumn(cellRef: String): Option[Int] =
      // Extract column from "A1" → 0, "B2" → 1, "AA1" → 26, etc.
      val col = cellRef.takeWhile(_.isLetter)
      if col.isEmpty then None
      else
        Some(
          col.foldLeft(0) { (acc, c) =>
            acc * 26 + (c.toUpper - 'A' + 1)
          } - 1
        )

    private def interpretCellValue(
      value: String,
      cellType: Option[String],
      sst: Option[SharedStrings]
    ): CellValue =
      cellType match
        case Some("s") =>
          // Shared string reference
          value.toIntOption
            .flatMap(idx => sst.flatMap(_.strings.lift(idx)))
            .map(CellValue.Text(_))
            .getOrElse(CellValue.Empty)

        case Some("inlineStr") =>
          CellValue.Text(value)

        case Some("n") =>
          // Number
          try CellValue.Number(BigDecimal(value))
          catch case _: NumberFormatException => CellValue.Empty

        case Some("b") =>
          // Boolean
          CellValue.Bool(value == "1" || value.equalsIgnoreCase("true"))

        case Some("e") =>
          // Error - map common Excel error strings
          val errorOpt = value match
            case "#DIV/0!" => Some(CellError.Div0)
            case "#N/A" => Some(CellError.NA)
            case "#NAME?" => Some(CellError.Name)
            case "#NULL!" => Some(CellError.Null)
            case "#NUM!" => Some(CellError.Num)
            case "#REF!" => Some(CellError.Ref)
            case "#VALUE!" => Some(CellError.Value)
            case _ => None
          errorOpt.map(CellValue.Error(_)).getOrElse(CellValue.Empty)

        case Some("str") =>
          // Formula result (string)
          CellValue.Text(value)

        case _ =>
          // Default: treat as number if possible, else text
          try CellValue.Number(BigDecimal(value))
          catch
            case _: NumberFormatException =>
              if value.nonEmpty then CellValue.Text(value) else CellValue.Empty

  // Helper: Extract text from List[XmlTexty]
  private def extractText(texties: List[xml.XmlEvent.XmlTexty]): Option[String] =
    if texties.isEmpty then None
    else Some(texties.collect { case xml.XmlEvent.XmlString(text, _) => text }.mkString)
