package com.tjclp.xl.io

import cats.effect.Sync
import cats.syntax.all.*
import fs2.Stream
import java.io.InputStream
import org.xml.sax.{InputSource, Attributes}
import org.xml.sax.helpers.DefaultHandler
import scala.collection.mutable
import com.tjclp.xl.cells.{CellValue, CellError}
import com.tjclp.xl.ooxml.{SharedStrings, XmlSecurity, XmlUtil}
import java.util.concurrent.{ArrayBlockingQueue, BlockingQueue}
import java.util.concurrent.atomic.AtomicBoolean
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.ooxml.SharedFormula

/**
 * SAX-based streaming XML reader for maximum performance.
 *
 * Uses javax.xml.parsers.SAXParser (native Java parser) instead of fs2-data-xml for 3-4x speedup.
 * Emits rows in chunks (default 1024) to minimize queue synchronization overhead.
 */
object SaxStreamingReader:
  // Chunk size for batching rows - reduces queue operations from N to N/chunkSize
  private val chunkSize = 1024
  // Queue capacity in chunks (not rows)
  private val queueCapacity = 16

  private sealed trait ChunkEvent
  private object ChunkEvent:
    final case class Rows(rows: Vector[RowData]) extends ChunkEvent
    final case class Error(err: Throwable) extends ChunkEvent
    case object End extends ChunkEvent

  // Stackless exception for efficient early abort (no message, no stacktrace)
  @SuppressWarnings(Array("org.wartremover.warts.Null"))
  private final class AbortParsing extends RuntimeException(null, null, false, false)

  /**
   * Stream worksheet rows using SAX parser (3-4x faster than fs2-data-xml).
   *
   * Uses chunked batching to minimize queue synchronization overhead. Memory usage is O(chunkSize +
   * active shared-formula groups), rather than O(total rows or total formula groups).
   *
   * @param stream
   *   Worksheet XML input stream
   * @param sst
   *   Optional SharedStrings table for resolving string references
   * @return
   *   Stream of RowData
   */
  def parseWorksheetStream[F[_]: Sync](
    stream: InputStream,
    sst: Option[SharedStrings]
  ): Stream[F, RowData] =
    parseWorksheetStream(stream, sst, None, None)

  def parseWorksheetStream[F[_]: Sync](
    stream: InputStream,
    sst: Option[SharedStrings],
    rowBounds: Option[(Int, Int)],
    colBounds: Option[(Int, Int)]
  ): Stream[F, RowData] =
    Stream
      .bracket {
        Sync[F].delay {
          val queue: BlockingQueue[ChunkEvent] = new ArrayBlockingQueue(queueCapacity)
          val cancelled = new AtomicBoolean(false)
          val parserThread = new Thread(
            () => runParser(stream, sst, rowBounds, colBounds, queue, cancelled),
            "xl-sax-stream"
          )
          parserThread.setDaemon(true)
          parserThread.start()
          (queue, cancelled, parserThread)
        }
      } { case (_, cancelled, parserThread) =>
        Sync[F].delay {
          cancelled.set(true)
          parserThread.interrupt()
        }
      }
      .flatMap { case (queue, _, _) =>
        // Use interruptible instead of blocking for lower overhead
        Stream
          .unfoldEval(false) { done =>
            if done then Sync[F].pure(None)
            else
              Sync[F].interruptible(queue.take()).map {
                case ChunkEvent.Rows(rows) => Some((Stream.emits(rows), false))
                case ChunkEvent.End => Some((Stream.empty, true))
                case ChunkEvent.Error(err) => Some((Stream.raiseError[F](err), true))
              }
          }
          .flatten
      }

  private def runParser(
    stream: InputStream,
    sst: Option[SharedStrings],
    rowBounds: Option[(Int, Int)],
    colBounds: Option[(Int, Int)],
    queue: BlockingQueue[ChunkEvent],
    cancelled: AtomicBoolean
  ): Unit =
    val rowBuffer = mutable.ArrayBuffer[RowData]()

    def flushBuffer(): Unit =
      if rowBuffer.nonEmpty then
        if cancelled.get then throw new AbortParsing
        try queue.put(ChunkEvent.Rows(rowBuffer.toVector))
        catch
          case _: InterruptedException =>
            Thread.currentThread().interrupt()
            throw new AbortParsing
        rowBuffer.clear()

    def emitRow(row: RowData): Unit =
      if cancelled.get then throw new AbortParsing
      rowBuffer += row
      if rowBuffer.size >= chunkSize then flushBuffer()

    try
      // GH-350: shared XXE hardening + benign-doctype strip, matching the in-memory parseSafe path
      val parser = XmlSecurity.secureSaxParserFactory().newSAXParser()
      val handler = new WorksheetHandler(sst, rowBounds, colBounds, emitRow, cancelled)
      parser.parse(InputSource(XmlSecurity.stripLeadingDoctypeStream(stream)), handler)
      // Flush any remaining rows
      flushBuffer()
      queue.put(ChunkEvent.End)
    catch
      case _: AbortParsing =>
        if !cancelled.get then
          try
            flushBuffer()
            queue.put(ChunkEvent.End)
          catch case _: InterruptedException => Thread.currentThread().interrupt()
      case _: InterruptedException =>
        Thread.currentThread().interrupt()
        ()
      case t: Throwable =>
        try queue.put(ChunkEvent.Error(t))
        catch case _: InterruptedException => Thread.currentThread().interrupt()

  /**
   * SAX handler for parsing worksheet XML.
   *
   * Maintains mutable state machine for current row/cell being parsed. Emits completed RowData when
   * </row> is encountered.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private class WorksheetHandler(
    sst: Option[SharedStrings],
    rowBounds: Option[(Int, Int)],
    colBounds: Option[(Int, Int)],
    emitRow: RowData => Unit,
    cancelled: AtomicBoolean
  ) extends DefaultHandler:
    // Mutable state for current parsing context
    var currentRowIndex: Int = 0
    var currentRowCells: mutable.Map[Int, CellValue] = mutable.Map.empty
    var currentRowStyles: mutable.Map[Int, Int] = mutable.Map.empty
    var inRow = false
    var skipRow = false

    // Current cell state
    var currentCellRef: Option[String] = None
    var currentCellARef: Option[ARef] = None
    var currentCellType: Option[String] = None
    var currentCellColIdx: Option[Int] = None
    var currentCellStyleId: Option[Int] = None
    var skipCell = false
    var inValue = false
    var inFormula = false
    var inInlineStr = false
    var inPhoneticRun = false
    var inInlineRun = false
    var sawInlineRun = false
    var inTextElement = false
    var cachedValue: Option[String] = None
    var formulaText: Option[String] = None
    var formulaType: Option[String] = None
    var sharedFormulaIndex: Option[SharedFormula.Index] = None
    var sharedFormulaRange: Option[CellRange] = None
    var inlineValue: Option[String] = None
    val valueText = new StringBuilder
    // Decoded inline-string runs for the current cell (each run decoded at </t>, GH-305)
    val inlineRuns = new StringBuilder
    // Bare <t> text directly under <is> (CT_Rst is (t?, r*, rPh*, phoneticPr?)); the DOM
    // reader ignores it whenever <r> runs exist, so it is accumulated separately.
    val bareInlineText = new StringBuilder
    // GH-370: shared formula masters are tiny compared with row data and let a one-pass stream
    // expand every spec-normal (master-first) dependent. Masters are captured even when their
    // cells fall outside requested row/column bounds.
    val sharedFormulaMasters: mutable.Map[SharedFormula.Index, SharedFormula.Master] =
      mutable.Map.empty

    // Check cancelled less frequently - every 10k elements
    private var elementCount = 0
    private def maybeCheckCancelled(): Unit =
      elementCount += 1
      if (elementCount & 0x3fff) == 0 then // Check every 16384 elements
        if cancelled.get then throw new AbortParsing

    override def startElement(
      uri: String,
      localName: String,
      qName: String,
      attributes: Attributes
    ): Unit =
      maybeCheckCancelled()
      localName match
        case "row" =>
          currentRowIndex = Option(attributes.getValue("r"))
            .flatMap(_.toIntOption)
            .getOrElse(currentRowIndex + 1)
          skipRow = rowBounds match
            case Some((startRow, endRow)) if currentRowIndex < startRow => true
            case Some((_, endRow)) if currentRowIndex > endRow => throw new AbortParsing
            case _ => false
          currentRowCells = mutable.Map.empty
          currentRowStyles = mutable.Map.empty
          inRow = true

        case "c" =>
          currentCellRef = Option(attributes.getValue("r"))
          currentCellARef = currentCellRef.flatMap(ARef.parse(_).toOption)
          currentCellARef.foreach(evictExpiredSharedMasters)
          currentCellType = Option(attributes.getValue("t"))
          currentCellStyleId = Option(attributes.getValue("s")).flatMap(_.toIntOption)
          currentCellColIdx = currentCellRef.flatMap(parseCellColumn)
          skipCell = skipRow || colBounds.exists { case (startCol, endCol) =>
            currentCellColIdx.forall(colIdx => colIdx < startCol || colIdx > endCol)
          }
          valueText.clear()

        case "v" =>
          inValue = true
          valueText.clear()

        case "f" =>
          inFormula = true
          formulaType = Option(attributes.getValue("t"))
          sharedFormulaIndex = Option(attributes.getValue("si")).flatMap(SharedFormula.parseIndex)
          sharedFormulaRange = Option(attributes.getValue("ref"))
            .flatMap(CellRange.parse(_).toOption)
          valueText.clear()

        case "is" =>
          inInlineStr = true
          inlineRuns.clear()
          bareInlineText.clear()
          sawInlineRun = false

        case "r" if inInlineStr && !inPhoneticRun =>
          inInlineRun = true

        case "rPh" if inInlineStr =>
          // Phonetic (furigana) runs are presentation metadata, not cell text;
          // the DOM reader and the SAX SST reader both exclude them.
          inPhoneticRun = true

        case "t" if inInlineStr && !inPhoneticRun =>
          inTextElement = true
          valueText.clear()

        case _ => ()

    override def characters(ch: Array[Char], start: Int, length: Int): Unit =
      // Formula metadata must be read even for a skipped cell: a bounded stream may select a
      // dependent while its shared master lies just outside the requested row/column range.
      if inFormula || ((inValue || inTextElement) && !skipRow && !skipCell) then
        valueText.appendAll(ch, start, length)

    override def endElement(uri: String, localName: String, qName: String): Unit =
      localName match
        case "v" if inValue =>
          if !skipRow && !skipCell then cachedValue = Some(valueText.toString)
          inValue = false
          valueText.clear()

        case "f" if inFormula =>
          // GH-293: the in-memory reader trims <f> text and treats a whitespace-only formula as
          // absent. Keep that behavior while retaining empty shared dependents (GH-370).
          val trimmed = valueText.toString.trim
          formulaText = Option.when(trimmed.nonEmpty)(trimmed)
          if formulaType.contains("shared") && trimmed.nonEmpty then
            for
              index <- sharedFormulaIndex
              ref <- currentCellARef
              range <- sharedFormulaRange.filter(_.contains(ref))
              if !sharedFormulaMasters.contains(index) // malformed active duplicate: first wins
            do sharedFormulaMasters(index) = SharedFormula.Master(ref, trimmed, Some(range))
          inFormula = false
          valueText.clear()

        case "t" if inTextElement =>
          // GH-305: decode each run BEFORE concatenation, mirroring the DOM reader's
          // per-run decode - an _xHHHH_ escape is only honored within a single <t>,
          // never across run boundaries. Bare <t> text (outside any <r>) is stashed
          // separately: the DOM reader ignores it whenever <r> runs exist.
          val decoded = XmlUtil.decodeXstring(valueText.toString)
          if inInlineRun then
            sawInlineRun = true
            inlineRuns.append(decoded)
          else bareInlineText.append(decoded)
          inTextElement = false
          valueText.clear()

        case "r" if inInlineRun =>
          inInlineRun = false

        case "rPh" if inPhoneticRun =>
          inPhoneticRun = false

        case "is" if inInlineStr =>
          // GH-293: defer the commit to </c> so the s= style index is recorded there,
          // exactly like SST and number cells. Runs win over a coexisting bare <t>
          // (DOM reader's rElems.nonEmpty dispatch); the bare <t> contributes only
          // when the <is> contains no runs.
          if !skipRow && !skipCell then
            inlineValue = Some(
              if sawInlineRun then inlineRuns.toString else bareInlineText.toString
            )
          inInlineStr = false
          inlineRuns.clear()
          bareInlineText.clear()

        case "c" =>
          if !skipRow && !skipCell then
            for colIdx <- currentCellColIdx do
              // Mirror of the DOM reader's dispatch (WorksheetReader.parseCellValue):
              // a nonEmpty formula wins; otherwise <is> wins for inline-string cell
              // types; otherwise interpret <v>.
              val inlineText = inlineValue.filter(_ => isInlineStringType)
              val isSharedMaster =
                formulaType.contains("shared") &&
                  formulaText.nonEmpty &&
                  sharedFormulaIndex.nonEmpty &&
                  currentCellARef.exists(ref => sharedFormulaRange.exists(_.contains(ref)))
              val expandedFormula =
                if !formulaType.contains("shared") || isSharedMaster then formulaText
                else
                  sharedFormulaIndex match
                    // A malformed/missing si cannot identify a group. Preserve a non-empty
                    // expression when present, otherwise keep the cell visibly formula-shaped.
                    case None => Some(formulaText.getOrElse("#REF!"))
                    case Some(index) =>
                      sharedFormulaMasters.get(index) match
                        case Some(master) =>
                          Some(
                            currentCellARef
                              .filter(ref => master.range.exists(_.contains(ref)))
                              .map(ref => SharedFormula.translate(master, ref))
                              .getOrElse("#REF!")
                          )
                        case None =>
                          // A one-pass row stream cannot look ahead for a later master without
                          // buffering already-emitted rows. Fail explicitly instead of silently
                          // converting the dependent's cached result into a literal.
                          val address = currentCellRef.getOrElse("<unknown>")
                          throw new IllegalArgumentException(
                            s"Shared formula at $address with si=$index requires the master to appear earlier in the worksheet"
                          )

              val cellValue = (expandedFormula, inlineText, cachedValue) match
                case (Some(formula), _, Some(cached)) =>
                  val parsedCached = interpretCellValue(cached, currentCellType, sst)
                  val cachedOpt =
                    if parsedCached == CellValue.Empty then None else Some(parsedCached)
                  CellValue.Formula(formula, cachedOpt)
                case (Some(formula), _, None) =>
                  CellValue.Formula(formula, None)
                case (None, Some(text), _) =>
                  CellValue.Text(text)
                case (None, None, Some(value)) =>
                  interpretCellValue(value, currentCellType, sst)
                case (None, None, None) =>
                  CellValue.Empty
              if cellValue != CellValue.Empty then
                currentRowCells(colIdx) = cellValue
                currentCellStyleId.foreach(sid => currentRowStyles(colIdx) = sid)

          currentCellRef = None
          currentCellARef = None
          currentCellType = None
          currentCellColIdx = None
          currentCellStyleId = None
          skipCell = false
          cachedValue = None
          formulaText = None
          formulaType = None
          sharedFormulaIndex = None
          sharedFormulaRange = None
          inlineValue = None
          inValue = false
          inFormula = false
          inInlineStr = false
          inPhoneticRun = false
          inInlineRun = false
          sawInlineRun = false
          inTextElement = false
          valueText.clear()
          inlineRuns.clear()
          bareInlineText.clear()

        case "row" if inRow =>
          if !skipRow then
            emitRow(RowData(currentRowIndex, currentRowCells.toMap, currentRowStyles.toMap))
          inRow = false
          skipRow = false

        case _ => ()

    /**
     * Masters with a valid `ref` range are live only through that range's row-major end. Eviction
     * keeps memory proportional to overlapping shared groups, not total sheet groups. A dependent
     * encountered after its master has expired is indistinguishable from a not-yet-seen/orphan
     * group and fails visibly instead of silently becoming a cached constant.
     */
    private def evictExpiredSharedMasters(current: ARef): Unit =
      sharedFormulaMasters.filterInPlace { case (_, master) =>
        master.range.exists { range =>
          val end = range.end
          end.row.index0 > current.row.index0 ||
          (end.row.index0 == current.row.index0 && end.col.index0 >= current.col.index0)
        }
      }

    /** Cell types whose value may come from an <is> inline string (DOM reader contract). */
    private def isInlineStringType: Boolean =
      currentCellType.exists(t => t == "inlineStr" || t == "str")

    private def parseCellColumn(cellRef: String): Option[Int] =
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
