package com.tjclp.xl.io.streaming

import java.io.InputStream
import org.xml.sax.{Attributes, InputSource}
import org.xml.sax.helpers.DefaultHandler
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.ooxml.{SharedFormula, SharedStrings, XmlSecurity, XmlUtil}

/**
 * SAX-based single cell reader with early-abort optimization.
 *
 * Uses a stackless exception pattern for O(position) time and O(1) target state. Parsing normally
 * aborts as soon as the target is resolved; a shared dependent may scan ahead for its master.
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

  /** Target data retained only when its shared master has not appeared yet. */
  private final case class PendingSharedTarget(
    ref: ARef,
    sharedIndex: SharedFormula.Index,
    cachedValue: Option[String],
    cellType: Option[String],
    styleId: Option[Int]
  )

  /**
   * Extract a single cell from worksheet XML using SAX parser with early-abort.
   *
   * Time complexity: O(position of cell in file) for ordinary/master-first cells. A shared
   * dependent whose master physically follows it scans until that master. Only shared masters whose
   * declared ranges contain the target are retained, so memory does not grow with total sheet
   * groups.
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
      // GH-350: shared XXE hardening + benign-doctype strip, matching the in-memory parseSafe path
      val parser = XmlSecurity.secureSaxParserFactory().newSAXParser()
      val handler = new SingleCellHandler(targetRef, sst)
      parser.parse(InputSource(XmlSecurity.stripLeadingDoctypeStream(stream)), handler)
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
    var currentCellRef: Option[ARef] = None
    var currentCellStyleId: Option[Int] = None

    // Cell content state
    var inValue = false
    var inFormula = false
    var inInlineStr = false
    var inTextElement = false
    var cachedValue: Option[String] = None
    var formulaText: Option[String] = None
    var formulaType: Option[String] = None
    var sharedFormulaIndex: Option[SharedFormula.Index] = None
    var sharedFormulaRange: Option[CellRange] = None
    var currentCellType: Option[String] = None
    var pendingTarget: Option[PendingSharedTarget] = None
    val valueText = new StringBuilder
    val sharedFormulaMasters
      : scala.collection.mutable.Map[SharedFormula.Index, SharedFormula.Master] =
      scala.collection.mutable.Map.empty

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
          if currentRowIndex > targetRowIndex && pendingTarget.isEmpty then throw new CellNotFound

        case "c" =>
          val cellRefText = Option(attributes.getValue("r"))
          currentCellRef = cellRefText.flatMap(ARef.parse(_).toOption)
          inTargetCell = cellRefText.contains(targetRefA1)
          currentCellType = None
          currentCellStyleId = None
          cachedValue = None
          formulaText = None
          formulaType = None
          sharedFormulaIndex = None
          sharedFormulaRange = None
          if inTargetCell then
            currentCellType = Option(attributes.getValue("t"))
            currentCellStyleId = Option(attributes.getValue("s")).flatMap(_.toIntOption)
            valueText.clear()

        case "v" if inTargetCell =>
          inValue = true
          valueText.clear()

        case "f" =>
          inFormula = true
          formulaType = Option(attributes.getValue("t"))
          sharedFormulaIndex = Option(attributes.getValue("si")).flatMap(SharedFormula.parseIndex)
          sharedFormulaRange = Option(attributes.getValue("ref"))
            .flatMap(CellRange.parse(_).toOption)
          valueText.clear()

        case "is" if inTargetCell =>
          inInlineStr = true

        case "t" if inTargetCell && inInlineStr =>
          inTextElement = true
          valueText.clear()

        case _ => ()

    override def characters(ch: Array[Char], start: Int, length: Int): Unit =
      // Every formula is observed so a shared master before or after the target can be retained.
      if inFormula || (inTargetCell && (inValue || inTextElement)) then
        valueText.appendAll(ch, start, length)

    override def endElement(uri: String, localName: String, qName: String): Unit =
      localName match
        case "v" if inValue && inTargetCell =>
          cachedValue = Some(valueText.toString)
          inValue = false
          valueText.clear()

        case "f" if inFormula =>
          val trimmed = valueText.toString.trim
          if inTargetCell then formulaText = Option.when(trimmed.nonEmpty)(trimmed)

          if formulaType.contains("shared") && trimmed.nonEmpty then
            for
              index <- sharedFormulaIndex
              ref <- currentCellRef
              range <- sharedFormulaRange.filter(_.contains(ref))
            do
              val master = SharedFormula.Master(ref, trimmed, Some(range))
              // A dependent may physically precede its master. Resolve it only from a true
              // master whose declared range contains both its origin and the target.
              pendingTarget
                .filter(pending => pending.sharedIndex == index && range.contains(pending.ref))
                .foreach { pending =>
                  throw new CellFound(resultForPending(pending, master))
                }
              // Retain only target-relevant masters, keeping memory independent of unrelated
              // shared-formula groups elsewhere in the sheet.
              if range.contains(targetRef) && !sharedFormulaMasters.contains(index) then
                sharedFormulaMasters(index) = master

          inFormula = false
          valueText.clear()

        case "t" if inTextElement && inTargetCell =>
          inTextElement = false

        case "is" if inInlineStr && inTargetCell =>
          cachedValue = Some(valueText.toString)
          inInlineStr = false
          valueText.clear()

        case "c" if inTargetCell =>
          val isSharedMaster =
            formulaType.contains("shared") &&
              formulaText.nonEmpty &&
              sharedFormulaIndex.nonEmpty &&
              currentCellRef.exists(ref => sharedFormulaRange.exists(_.contains(ref)))
          val expandedFormula =
            if !formulaType.contains("shared") || isSharedMaster then formulaText
            else
              sharedFormulaIndex match
                // With no usable group index, retain a non-empty expression as a tolerant
                // isolated formula fallback; otherwise surface the malformed reference.
                case None => Some(formulaText.getOrElse("#REF!"))
                case Some(index) =>
                  for
                    master <- sharedFormulaMasters.get(index)
                    ref <- currentCellRef
                    if master.range.exists(_.contains(ref))
                  yield SharedFormula.translate(master, ref)

          if formulaType.contains("shared") && expandedFormula.isEmpty then
            val index = sharedFormulaIndex.getOrElse(
              throw new IllegalStateException(s"Shared formula at $targetRefA1 lost its si")
            )
            pendingTarget = Some(
              PendingSharedTarget(
                targetRef,
                index,
                cachedValue,
                currentCellType,
                currentCellStyleId
              )
            )
            resetCellState()
          else
            throw new CellFound(
              buildResult(expandedFormula, cachedValue, currentCellType, currentCellStyleId)
            )

        case "c" => resetCellState()

        case _ => ()

    override def endDocument(): Unit =
      pendingTarget.foreach { pending =>
        throw new CellFound(
          buildResult(Some("#REF!"), pending.cachedValue, pending.cellType, pending.styleId)
        )
      }

    private def resultForPending(
      pending: PendingSharedTarget,
      master: SharedFormula.Master
    ): CellResult =
      val formula = SharedFormula.translate(master, pending.ref)
      buildResult(Some(formula), pending.cachedValue, pending.cellType, pending.styleId)

    private def buildResult(
      expandedFormula: Option[String],
      cached: Option[String],
      cellType: Option[String],
      styleId: Option[Int]
    ): CellResult =
      val cellValue = (expandedFormula, cached) match
        case (Some(formula), Some(cachedText)) =>
          val parsedCached = interpretCellValue(cachedText, cellType, sst)
          val cachedOpt = Option.when(parsedCached != CellValue.Empty)(parsedCached)
          CellValue.Formula(formula, cachedOpt)
        case (Some(formula), None) => CellValue.Formula(formula, None)
        case (None, Some(value)) => interpretCellValue(value, cellType, sst)
        case (None, None) => CellValue.Empty

      CellResult(cellValue, styleId, expandedFormula)

    private def resetCellState(): Unit =
      currentCellRef = None
      inTargetCell = false
      currentCellStyleId = None
      currentCellType = None
      cachedValue = None
      formulaText = None
      formulaType = None
      sharedFormulaIndex = None
      sharedFormulaRange = None
      inValue = false
      inFormula = false
      inInlineStr = false
      inTextElement = false
      valueText.clear()

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
          CellValue.Text(XmlUtil.decodeXstring(value))

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
          CellValue.Text(XmlUtil.decodeXstring(value))

        case _ =>
          try CellValue.Number(BigDecimal(value))
          catch
            case _: NumberFormatException =>
              if value.nonEmpty then CellValue.Text(value) else CellValue.Empty
