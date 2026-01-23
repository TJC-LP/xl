package com.tjclp.xl.io

import cats.effect.Sync
import cats.syntax.all.*
import fs2.Stream
import java.io.InputStream
import javax.xml.parsers.SAXParserFactory
import org.xml.sax.{InputSource, Attributes}
import org.xml.sax.helpers.DefaultHandler
import scala.collection.mutable
import com.tjclp.xl.cells.{CellValue, CellError}
import com.tjclp.xl.ooxml.SharedStrings
import java.util.concurrent.{ArrayBlockingQueue, BlockingQueue}
import java.util.concurrent.atomic.AtomicBoolean

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
   * Uses chunked batching to minimize queue synchronization overhead. Memory usage is O(chunkSize)
   * rather than O(total rows).
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
      val factory = SAXParserFactory.newInstance()
      factory.setNamespaceAware(true)
      val parser = factory.newSAXParser()
      val handler = new WorksheetHandler(sst, rowBounds, colBounds, emitRow, cancelled)
      parser.parse(InputSource(stream), handler)
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
    var inRow = false
    var skipRow = false

    // Current cell state
    var currentCellRef: Option[String] = None
    var currentCellType: Option[String] = None
    var currentCellColIdx: Option[Int] = None
    var skipCell = false
    var inValue = false
    var inFormula = false
    var inInlineStr = false
    var inTextElement = false
    var cachedValue: Option[String] = None
    var formulaText: Option[String] = None
    val valueText = new StringBuilder

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
          inRow = true

        case "c" =>
          currentCellRef = Option(attributes.getValue("r"))
          currentCellType = Option(attributes.getValue("t"))
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
          valueText.clear()

        case "is" =>
          inInlineStr = true

        case "t" if inInlineStr =>
          inTextElement = true
          valueText.clear()

        case _ => ()

    override def characters(ch: Array[Char], start: Int, length: Int): Unit =
      if (inValue || inFormula || inTextElement) && !skipRow && !skipCell then
        valueText.appendAll(ch, start, length)

    override def endElement(uri: String, localName: String, qName: String): Unit =
      localName match
        case "v" if inValue =>
          if !skipRow && !skipCell then cachedValue = Some(valueText.toString)
          inValue = false
          valueText.clear()

        case "f" if inFormula =>
          if !skipRow && !skipCell then formulaText = Some(valueText.toString)
          inFormula = false
          valueText.clear()

        case "t" if inTextElement =>
          inTextElement = false

        case "is" if inInlineStr =>
          if !skipRow && !skipCell then
            val text = valueText.toString
            for colIdx <- currentCellColIdx do currentRowCells(colIdx) = CellValue.Text(text)
          inInlineStr = false
          valueText.clear()

        case "c" =>
          if !skipRow && !skipCell then
            for colIdx <- currentCellColIdx do
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
              if cellValue != CellValue.Empty then currentRowCells(colIdx) = cellValue

          currentCellRef = None
          currentCellType = None
          currentCellColIdx = None
          skipCell = false
          cachedValue = None
          formulaText = None
          inValue = false
          inFormula = false
          inInlineStr = false
          inTextElement = false
          valueText.clear()

        case "row" if inRow =>
          if !skipRow then emitRow(RowData(currentRowIndex, currentRowCells.toMap))
          inRow = false
          skipRow = false

        case _ => ()

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
