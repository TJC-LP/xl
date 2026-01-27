package com.tjclp.xl.io.streaming

import java.io.InputStream
import javax.xml.parsers.SAXParserFactory
import org.xml.sax.{Attributes, InputSource}
import org.xml.sax.helpers.DefaultHandler
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.ooxml.SharedStrings

/**
 * SAX-based single cell reader with early-abort optimization.
 *
 * Uses stackless exception pattern for O(position) time, O(1) memory extraction of a single cell.
 * Aborts parsing as soon as target cell is found or row index exceeds target.
 */
object SaxSingleCellReader:

  /**
   * Result of extracting a single cell from worksheet XML.
   *
   * @param value
   *   Cell value (Number, Text, Bool, Formula, etc.)
   * @param styleId
   *   Optional style index for resolving CellStyle
   * @param formulaText
   *   Formula expression if cell contains a formula
   */
  final case class CellResult(
    value: CellValue,
    styleId: Option[Int],
    formulaText: Option[String]
  )

  // Stackless exception for efficient early abort (no message, no stacktrace)
  @SuppressWarnings(Array("org.wartremover.warts.Null"))
  private final class CellFound(val result: CellResult)
      extends RuntimeException(null, null, false, false)

  // Stackless exception for when target cell's row has been passed
  @SuppressWarnings(Array("org.wartremover.warts.Null"))
  private final class CellNotFound extends RuntimeException(null, null, false, false)

  /**
   * Extract a single cell from worksheet XML using SAX parser with early-abort.
   *
   * Time complexity: O(position of cell in file) Memory: O(1) - only current cell buffered
   *
   * @param stream
   *   Worksheet XML input stream
   * @param targetRef
   *   Cell reference to find (e.g., A1, B5)
   * @param sst
   *   Optional SharedStrings table for resolving string references
   * @return
   *   Some(CellResult) if cell exists, None if cell is empty/missing
   */
  def extractCell(
    stream: InputStream,
    targetRef: ARef,
    sst: Option[SharedStrings]
  ): Option[CellResult] =
    try
      val factory = SAXParserFactory.newInstance()
      factory.setNamespaceAware(true)
      val parser = factory.newSAXParser()
      val handler = new SingleCellHandler(targetRef, sst)
      parser.parse(InputSource(stream), handler)
      // Reached end of document without finding cell
      None
    catch
      case found: CellFound => Some(found.result)
      case _: CellNotFound => None

  /**
   * SAX handler for extracting a single cell.
   *
   * Tracks row index from <row r="N"> attributes and aborts early when: - Target cell is found
   * (throws CellFound) - Row index exceeds target row (throws CellNotFound, cell doesn't exist)
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private class SingleCellHandler(
    targetRef: ARef,
    sst: Option[SharedStrings]
  ) extends DefaultHandler:
    // Target cell reference in A1 notation for matching
    private val targetRefA1 = targetRef.toA1
    private val targetRowIndex = targetRef.row.index1

    // Mutable state for current parsing context
    var currentRowIndex: Int = 0
    var inTargetCell = false
    var currentCellStyleId: Option[Int] = None

    // Cell content state
    var inValue = false
    var inFormula = false
    var inInlineStr = false
    var inTextElement = false
    var cachedValue: Option[String] = None
    var formulaText: Option[String] = None
    var currentCellType: Option[String] = None
    val valueText = new StringBuilder

    override def startElement(
      uri: String,
      localName: String,
      qName: String,
      attributes: Attributes
    ): Unit =
      localName match
        case "row" =>
          currentRowIndex = Option(attributes.getValue("r"))
            .flatMap(_.toIntOption)
            .getOrElse(currentRowIndex + 1)
          // Early abort if we've passed the target row
          if currentRowIndex > targetRowIndex then throw new CellNotFound

        case "c" =>
          val cellRef = Option(attributes.getValue("r"))
          if cellRef.contains(targetRefA1) then
            inTargetCell = true
            currentCellType = Option(attributes.getValue("t"))
            currentCellStyleId = Option(attributes.getValue("s")).flatMap(_.toIntOption)
            valueText.clear()
            cachedValue = None
            formulaText = None

        case "v" if inTargetCell =>
          inValue = true
          valueText.clear()

        case "f" if inTargetCell =>
          inFormula = true
          valueText.clear()

        case "is" if inTargetCell =>
          inInlineStr = true

        case "t" if inTargetCell && inInlineStr =>
          inTextElement = true
          valueText.clear()

        case _ => ()

    override def characters(ch: Array[Char], start: Int, length: Int): Unit =
      if inTargetCell && (inValue || inFormula || inTextElement) then
        valueText.appendAll(ch, start, length)

    override def endElement(uri: String, localName: String, qName: String): Unit =
      localName match
        case "v" if inValue && inTargetCell =>
          cachedValue = Some(valueText.toString)
          inValue = false
          valueText.clear()

        case "f" if inFormula && inTargetCell =>
          formulaText = Some(valueText.toString)
          inFormula = false
          valueText.clear()

        case "t" if inTextElement && inTargetCell =>
          inTextElement = false

        case "is" if inInlineStr && inTargetCell =>
          cachedValue = Some(valueText.toString)
          inInlineStr = false
          valueText.clear()

        case "c" if inTargetCell =>
          // Cell complete - construct result and abort
          val cellValue = (formulaText, cachedValue) match
            case (Some(formula), Some(cached)) =>
              val parsedCached = interpretCellValue(cached, currentCellType, sst)
              val cachedOpt =
                if parsedCached == CellValue.Empty then None else Some(parsedCached)
              CellValue.Formula(formula, cachedOpt)
            case (Some(formula), None) =>
              CellValue.Formula(formula, None)
            case (None, Some(value)) =>
              interpretCellValue(value, currentCellType, sst)
            case (None, None) =>
              CellValue.Empty

          val result = CellResult(
            value = cellValue,
            styleId = currentCellStyleId,
            formulaText = formulaText
          )
          throw new CellFound(result)

        case _ => ()

    private def interpretCellValue(
      value: String,
      cellType: Option[String],
      sst: Option[SharedStrings]
    ): CellValue =
      cellType match
        case Some("s") =>
          (for {
            sharedStrings <- sst
            idx <- value.toIntOption
            entry <- sharedStrings.apply(idx)
          } yield sharedStrings.toCellValue(entry))
            .getOrElse(CellValue.Empty)

        case Some("inlineStr") =>
          CellValue.Text(value)

        case Some("n") =>
          try CellValue.Number(BigDecimal(value))
          catch case _: NumberFormatException => CellValue.Empty

        case Some("b") =>
          CellValue.Bool(value == "1" || value.equalsIgnoreCase("true"))

        case Some("e") =>
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
          CellValue.Text(value)

        case _ =>
          try CellValue.Number(BigDecimal(value))
          catch
            case _: NumberFormatException =>
              if value.nonEmpty then CellValue.Text(value) else CellValue.Empty
