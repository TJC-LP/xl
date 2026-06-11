package com.tjclp.xl.io

import fs2.Stream
import fs2.data.xml.*
import fs2.data.xml.XmlEvent.*
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.ooxml.{FormulaInjectionPolicy, SSTEntry, SharedStrings, XmlUtil}
import com.tjclp.xl.ooxml.style.{OoxmlStyles, StyleIndex}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.border.Border
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.styles.units.StyleId

/**
 * Tracks worksheet bounds during streaming. O(1) memory (4 integers).
 *
 * Thread-safety: NOT thread-safe. Use from a single fiber/thread only. The fs2 Stream processing
 * model ensures single-threaded access when used via evalTap.
 *
 * @example
 *   {{{
 *   val bounds = new BoundsAccumulator()
 *   rows.evalTap(row => Sync[F].delay(bounds.update(row)))
 *   // After stream completes:
 *   bounds.dimension // Option[CellRange] with tracked extent
 *   }}}
 */
// Mutable accumulator for O(1) memory bounds tracking during streaming
@SuppressWarnings(Array("org.wartremover.warts.Var"))
final class BoundsAccumulator:
  private var minRow: Int = Int.MaxValue
  private var maxRow: Int = Int.MinValue
  private var minCol: Int = Int.MaxValue
  private var maxCol: Int = Int.MinValue
  private var hasData: Boolean = false

  /**
   * Update bounds with row data.
   *
   * @param row
   *   Row data with 1-based rowIndex and 0-based column keys in cells map
   */
  // IterableOps: .min/.max safe because row.cells.nonEmpty checked first
  @SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
  def update(row: RowData): Unit =
    if row.cells.nonEmpty then
      hasData = true
      minRow = minRow min row.rowIndex
      maxRow = maxRow max row.rowIndex
      val cols = row.cells.keys
      minCol = minCol min cols.min
      maxCol = maxCol max cols.max

  /**
   * Get tracked dimension as CellRange.
   *
   * @return
   *   Some(CellRange) if any data was written, None if stream was empty
   */
  def dimension: Option[CellRange] =
    if hasData then
      // rowIndex is 1-based, so convert to 0-based for ARef.from0
      Some(
        CellRange(
          ARef.from0(minCol, minRow - 1),
          ARef.from0(maxCol, maxRow - 1)
        )
      )
    else None

/**
 * Two-pass shared-strings accumulator for streaming writes (GH-223).
 *
 * Pass 1 (while rows stream): every plain-text cell calls [[indexFor]], which assigns SST indices
 * in FIRST-OCCURRENCE order — so the index a cell emits is final the moment it is assigned, and the
 * worksheet body never needs a second pass. Pass 2 (after the body): [[toSharedStrings]] yields the
 * table for xl/sharedStrings.xml.
 *
 * Memory: O(distinct strings) — the accepted envelope per docs/design/smart-streaming.md. The
 * reference count feeds the SST `count` attribute (counts references, not unique entries).
 *
 * Thread-safety: NOT thread-safe; single-fiber use only (same contract as [[BoundsAccumulator]]).
 */
@SuppressWarnings(Array("org.wartremover.warts.Var"))
final class SstAccumulator:
  private val indexByString = scala.collection.mutable.LinkedHashMap.empty[String, Int]
  private var references: Int = 0

  /** SST index for `s`, assigning the next first-occurrence index when unseen. */
  def indexFor(s: String): Int =
    references += 1
    indexByString.getOrElseUpdate(s, indexByString.size)

  def uniqueCount: Int = indexByString.size

  def referenceCount: Int = references

  /** Build the table for xl/sharedStrings.xml: entries in first-occurrence order. */
  def toSharedStrings: SharedStrings =
    val entries = indexByString.keysIterator.map(s => Left(s): SSTEntry).toVector
    SharedStrings(entries, indexByString.toMap, references)

/**
 * True streaming XML writer using fs2-data-xml for constant-memory writes.
 *
 * Emits XML events incrementally as rows arrive, never materializing the full document tree in
 * memory.
 *
 * The two-phase write approach enables dimension element support:
 *   - Phase 1: Stream rows to temp file, track bounds with BoundsAccumulator
 *   - Phase 2: Assemble final ZIP with dimension element (from tracked bounds)
 *
 * This maintains O(1) memory while producing files with instant metadata queries.
 *
 * Shared strings (GH-223): pass an [[SstAccumulator]] to the body/row/cell helpers to emit plain
 * text as `t="s"` SST references (first-occurrence indices); without one, text stays
 * `t="inlineStr"`. RichText always stays inline (mixed dialects are valid OOXML).
 */
object StreamingXmlWriter:

  // Helper to create attributes
  private def attr(name: String, value: String): Attr =
    Attr(QName(name), List(XmlString(value, false)))

  // OOXML namespaces
  private val nsSpreadsheetML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  private val nsRelationships =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  /** Convert column index (0-based) to Excel letter (A, B, AA, etc.) */
  def columnToLetter(col: Int): String =
    def loop(n: Int, acc: String): String =
      if n < 0 then acc
      else loop((n / 26) - 1, s"${((n % 26) + 'A').toChar}$acc")
    loop(col, "")

  /**
   * Generate XML events for a single cell
   *
   * Converts CellValue to fs2-data-xml events for true streaming output.
   *
   * REQUIRES:
   *   - colIndex in range 0..16383 (Excel column limit)
   *   - rowIndex in range 1..1048576 (Excel row limit, 1-based)
   *   - value is valid CellValue
   * ENSURES:
   *   - Returns empty list for CellValue.Empty (no cell emitted)
   *   - Returns complete <c> element events for non-empty cells
   *   - Text cells: adds xml:space="preserve" via QName(Some("xml"), "space") when needed
   *   - Whitespace preserved byte-for-byte for text with leading/trailing/double spaces
   *   - All events are well-formed and balance StartTag/EndTag
   *   - Formula injection escaping applied when policy is Escape
   *   - Style attribute s="N" emitted when styleId is provided
   * DETERMINISTIC: Yes (pure transformation of inputs) ERROR CASES: None (total function over valid
   * CellValue)
   *
   * @param colIndex
   *   Column index (0-based)
   * @param rowIndex
   *   Row index (1-based for Excel compatibility)
   * @param value
   *   Cell value to serialize
   * @param injectionPolicy
   *   Formula injection escaping policy (default: None)
   * @param styleId
   *   Optional style index for s="N" attribute (from StyleIndex)
   * @param sst
   *   Optional SST accumulator (GH-223): when present, plain text is emitted as a t="s" reference
   *   whose index is assigned on first occurrence; when absent, text stays inline
   * @return
   *   List of XML events representing the cell
   */
  def cellToEvents(
    colIndex: Int,
    rowIndex: Int,
    value: CellValue,
    injectionPolicy: FormulaInjectionPolicy = FormulaInjectionPolicy.None,
    styleId: Option[Int] = None,
    sst: Option[SstAccumulator] = None
  ): List[XmlEvent] =
    val ref = s"${columnToLetter(colIndex)}$rowIndex"

    val (cellType, valueEvents) = value match
      case CellValue.Empty =>
        ("", Nil)

      case CellValue.Text(s) =>
        // Apply formula injection escaping if policy requires it (BEFORE SST accumulation, so the
        // table stores the escaped text), then _xHHHH_ escapes for the inline dialect (GH-288)
        val injectionEscaped = injectionPolicy match
          case FormulaInjectionPolicy.Escape => CellValue.escape(s)
          case FormulaInjectionPolicy.None => s
        sst match
          case Some(acc) =>
            // <c r="A1" t="s"><v>idx</v></c> — SST reference (GH-223). The accumulator stores the
            // raw string; SharedStrings serialization applies _xHHHH_/xml:space at emission.
            val idx = acc.indexFor(injectionEscaped)
            (
              "s",
              List(
                XmlEvent.StartTag(QName("v"), Nil, false),
                XmlEvent.XmlString(idx.toString, false),
                XmlEvent.EndTag(QName("v"))
              )
            )
          case None =>
            // <c r="A1" t="inlineStr"><is><t>text</t></is></c>
            val escaped = XmlUtil.escapeXstring(injectionEscaped)
            // Add xml:space="preserve" for text with leading/trailing/multiple spaces
            // Note: we check the escaped text because whitespace in the original (e.g., "  =formula")
            // must still be preserved after escaping adds a leading quote (e.g., "'  =formula")
            val needsPreserve = XmlUtil.needsXmlSpacePreserve(escaped)
            val tAttrs =
              if needsPreserve then
                List(Attr(QName(Some("xml"), "space"), List(XmlString("preserve", false))))
              else Nil

            (
              "inlineStr",
              List(
                XmlEvent.StartTag(QName("is"), Nil, false),
                XmlEvent.StartTag(QName("t"), tAttrs, false),
                XmlEvent.XmlString(escaped, false),
                XmlEvent.EndTag(QName("t")),
                XmlEvent.EndTag(QName("is"))
              )
            )

      case CellValue.Number(n) =>
        // <c r="A1" t="n"><v>42</v></c>
        (
          "n",
          List(
            XmlEvent.StartTag(QName("v"), Nil, false),
            XmlEvent.XmlString(n.toString, false),
            XmlEvent.EndTag(QName("v"))
          )
        )

      case CellValue.Bool(b) =>
        // <c r="A1" t="b"><v>1</v></c>
        (
          "b",
          List(
            XmlEvent.StartTag(QName("v"), Nil, false),
            XmlEvent.XmlString(if b then "1" else "0", false),
            XmlEvent.EndTag(QName("v"))
          )
        )

      case CellValue.Formula(expr, cachedValue) =>
        // <c r="A1"><f>SUM(A1:A10)</f><v>100</v></c>
        val formulaEvents = List(
          XmlEvent.StartTag(QName("f"), Nil, false),
          XmlEvent.XmlString(expr, false),
          XmlEvent.EndTag(QName("f"))
        )
        val cachedEvents = cachedValue.toList.flatMap {
          case CellValue.Number(num) =>
            List(
              XmlEvent.StartTag(QName("v"), Nil, false),
              XmlEvent.XmlString(num.toString, false),
              XmlEvent.EndTag(QName("v"))
            )
          case CellValue.Text(s) =>
            List(
              XmlEvent.StartTag(QName("v"), Nil, false),
              XmlEvent.XmlString(XmlUtil.escapeXstring(s), false),
              XmlEvent.EndTag(QName("v"))
            )
          case CellValue.Bool(b) =>
            List(
              XmlEvent.StartTag(QName("v"), Nil, false),
              XmlEvent.XmlString(if b then "1" else "0", false),
              XmlEvent.EndTag(QName("v"))
            )
          case CellValue.Error(err) =>
            import com.tjclp.xl.cells.CellError.toExcel
            List(
              XmlEvent.StartTag(QName("v"), Nil, false),
              XmlEvent.XmlString(err.toExcel, false),
              XmlEvent.EndTag(QName("v"))
            )
          case _ => Nil // Empty, RichText, Formula, DateTime - don't write
        }
        // Cell type based on cached value
        val cellType = cachedValue match
          case Some(CellValue.Number(_)) => "n"
          case Some(CellValue.Bool(_)) => "b"
          case Some(CellValue.Error(_)) => "e"
          case _ => ""
        (cellType, formulaEvents ++ cachedEvents)

      case CellValue.Error(err) =>
        import com.tjclp.xl.cells.CellError.toExcel
        // <c r="A1" t="e"><v>#DIV/0!</v></c>
        (
          "e",
          List(
            XmlEvent.StartTag(QName("v"), Nil, false),
            XmlEvent.XmlString(err.toExcel, false),
            XmlEvent.EndTag(QName("v"))
          )
        )

      case CellValue.DateTime(dt) =>
        // Convert to Excel serial number
        val serial = CellValue.dateTimeToExcelSerial(dt)
        (
          "n",
          List(
            XmlEvent.StartTag(QName("v"), Nil, false),
            XmlEvent.XmlString(serial.toString, false),
            XmlEvent.EndTag(QName("v"))
          )
        )

      case CellValue.RichText(richText) =>
        // <c r="A1" t="inlineStr"><is><r><rPr>...</rPr><t>text</t></r>...</is></c>
        val isEvents = richText.runs.flatMap { run =>
          val rStart = XmlEvent.StartTag(QName("r"), Nil, false)

          // Optional <rPr> with font properties.
          // NOTE: fs2-data-xml renders StartTag(isEmpty = true) as a self-closing element by
          // pairing it with its MATCHING EndTag event — an isEmpty StartTag without an EndTag
          // desynchronizes the renderer and swallows the next close tag (</rPr> went missing,
          // producing malformed worksheets for any styled rich-text run).
          val rPrEvents = run.font.toList.flatMap { f =>
            val propsBuilder = List.newBuilder[XmlEvent]

            def selfClosing(label: String, attrs: List[Attr] = Nil): Unit =
              propsBuilder += XmlEvent.StartTag(QName(label), attrs, true)
              propsBuilder += XmlEvent.EndTag(QName(label))

            // Font style properties
            if f.bold then selfClosing("b")
            if f.italic then selfClosing("i")
            if f.underline then selfClosing("u")

            // Color
            f.color.foreach { c =>
              selfClosing(
                "color",
                List(Attr(QName("rgb"), List(XmlString(c.toHex.drop(1), false))))
              )
            }

            // Size and name
            selfClosing("sz", List(Attr(QName("val"), List(XmlString(f.sizePt.toString, false)))))
            selfClosing("name", List(Attr(QName("val"), List(XmlString(f.name, false)))))

            XmlEvent.StartTag(QName("rPr"), Nil, false) :: propsBuilder.result() ::: List(
              XmlEvent.EndTag(QName("rPr"))
            )
          }

          // Text element with optional xml:space="preserve" (_xHHHH_-escaped per GH-288)
          val runText = XmlUtil.escapeXstring(run.text)
          val tAttrs =
            if XmlUtil.needsXmlSpacePreserve(runText) then
              List(Attr(QName(Some("xml"), "space"), List(XmlString("preserve", false))))
            else Nil

          val textEvents = List(
            XmlEvent.StartTag(QName("t"), tAttrs, false),
            XmlEvent.XmlString(runText, false),
            XmlEvent.EndTag(QName("t"))
          )

          rStart :: (rPrEvents ::: textEvents ::: List(XmlEvent.EndTag(QName("r"))))
        }.toList

        (
          "inlineStr",
          XmlEvent.StartTag(QName("is"), Nil, false) :: isEvents ::: List(
            XmlEvent.EndTag(QName("is"))
          )
        )

    // Build cell element with optional style attribute
    val baseAttrs =
      if cellType.nonEmpty then List(attr("r", ref), attr("t", cellType))
      else List(attr("r", ref))
    // Add s="N" attribute if styleId provided (must be > 0 to override default)
    val attrs = styleId.filter(_ > 0) match
      case Some(sid) => baseAttrs :+ attr("s", sid.toString)
      case None => baseAttrs

    if valueEvents.isEmpty && value == CellValue.Empty then Nil // Don't emit empty cells
    else
      XmlEvent.StartTag(QName("c"), attrs, valueEvents.isEmpty) ::
        valueEvents :::
        (if valueEvents.nonEmpty then List(XmlEvent.EndTag(QName("c"))) else Nil)

  /** Generate XML events for a row */
  def rowToEvents(
    rowData: RowData,
    injectionPolicy: FormulaInjectionPolicy = FormulaInjectionPolicy.None,
    sst: Option[SstAccumulator] = None
  ): List[XmlEvent] =
    val rowAttrs = List(attr("r", rowData.rowIndex.toString))

    // Sort cells by column for deterministic output
    val cellEvents = rowData.cells.toList
      .sortBy(_._1)
      .flatMap { case (colIdx, value) =>
        cellToEvents(colIdx, rowData.rowIndex, value, injectionPolicy, None, sst)
      }

    if cellEvents.isEmpty then Nil // Skip empty rows
    else
      XmlEvent.StartTag(QName("row"), rowAttrs, false) ::
        cellEvents :::
        XmlEvent.EndTag(QName("row")) :: Nil

  /** Generate XML events for a styled row (with s="N" attributes) */
  def styledRowToEvents(
    rowData: StyledRowData,
    injectionPolicy: FormulaInjectionPolicy = FormulaInjectionPolicy.None,
    sst: Option[SstAccumulator] = None
  ): List[XmlEvent] =
    val rowAttrs = List(attr("r", rowData.rowIndex.toString))

    // Sort cells by column for deterministic output
    val cellEvents = rowData.cells.toList
      .sortBy(_._1)
      .flatMap { case (colIdx, value) =>
        val styleId = rowData.cellStyles.get(colIdx)
        cellToEvents(colIdx, rowData.rowIndex, value, injectionPolicy, styleId, sst)
      }

    if cellEvents.isEmpty then Nil // Skip empty rows
    else
      XmlEvent.StartTag(QName("row"), rowAttrs, false) ::
        cellEvents :::
        XmlEvent.EndTag(QName("row")) :: Nil

  /**
   * Generate worksheet header events with optional dimension element.
   *
   * The dimension element must appear before sheetData in OOXML. When provided, it enables instant
   * metadata queries without streaming scan.
   *
   * @param dimension
   *   Optional cell range for dimension element. None omits the element.
   * @return
   *   List of XML events for worksheet opening and sheetData start
   */
  def worksheetHeader(dimension: Option[CellRange]): List[XmlEvent] =
    val dimEvents = dimension.toList.flatMap { range =>
      List(
        XmlEvent.StartTag(
          QName("dimension"),
          List(attr("ref", range.toA1)),
          true // self-closing: <dimension ref="A1:Z100"/>
        ),
        // fs2-data-xml pairs an isEmpty StartTag with its matching EndTag to render the
        // self-closing form; an unpaired isEmpty StartTag desynchronizes the renderer
        XmlEvent.EndTag(QName("dimension"))
      )
    }

    List(
      XmlEvent.StartTag(
        QName("worksheet"),
        List(
          attr("xmlns", nsSpreadsheetML),
          Attr(QName(Some("xmlns"), "r"), List(XmlString(nsRelationships, false)))
        ),
        false
      )
    ) ++ dimEvents ++ List(
      XmlEvent.StartTag(QName("sheetData"), Nil, false)
    )

  /**
   * Generate row events for worksheet body.
   *
   * Emits events incrementally as rows arrive, never materializing the full dataset.
   *
   * @param rows
   *   Stream of row data
   * @param injectionPolicy
   *   Formula injection escaping policy
   * @param sst
   *   Optional SST accumulator (GH-223) for t="s" string references
   * @return
   *   Stream of XML events for all rows
   */
  def worksheetBody[F[_]](
    rows: Stream[F, RowData],
    injectionPolicy: FormulaInjectionPolicy = FormulaInjectionPolicy.None,
    sst: Option[SstAccumulator] = None
  ): Stream[F, XmlEvent] =
    rows.flatMap(row => Stream.emits(rowToEvents(row, injectionPolicy, sst)))

  /**
   * Generate row events for styled worksheet body.
   *
   * Emits events incrementally as rows arrive, including s="N" style attributes, for row-streaming
   * paths that provide explicit style ids.
   *
   * @param rows
   *   Stream of styled row data
   * @param injectionPolicy
   *   Formula injection escaping policy
   * @param sst
   *   Optional SST accumulator (GH-223) for t="s" string references
   * @return
   *   Stream of XML events for all rows with styles
   */
  def worksheetBodyStyled[F[_]](
    rows: Stream[F, StyledRowData],
    injectionPolicy: FormulaInjectionPolicy = FormulaInjectionPolicy.None,
    sst: Option[SstAccumulator] = None
  ): Stream[F, XmlEvent] =
    rows.flatMap(row => Stream.emits(styledRowToEvents(row, injectionPolicy, sst)))

  /**
   * Generate worksheet footer events.
   *
   * @return
   *   List of XML events for closing sheetData and worksheet
   */
  def worksheetFooter: List[XmlEvent] =
    List(
      XmlEvent.EndTag(QName("sheetData")),
      XmlEvent.EndTag(QName("worksheet"))
    )

  /**
   * Stream worksheet XML events with header/footer scaffolding.
   *
   * This is the original single-pass API that omits dimension. Use worksheetHeader/Body/Footer for
   * two-phase writes with dimension support.
   *
   * @param rows
   *   Stream of row data
   * @param injectionPolicy
   *   Formula injection escaping policy (default: None)
   */
  def worksheetEvents[F[_]](
    rows: Stream[F, RowData],
    injectionPolicy: FormulaInjectionPolicy = FormulaInjectionPolicy.None
  ): Stream[F, XmlEvent] =
    Stream.emits(worksheetHeader(None)) ++
      worksheetBody(rows, injectionPolicy) ++
      Stream.emits(worksheetFooter)

  /**
   * Build the styles.xml registry for a styled streaming write (GH-223).
   *
   * `styles(i)` is the CellStyle a `StyledRowData.cellStyles` value `i` refers to. Returns the
   * serializable [[OoxmlStyles]] (cellXf 0 is always CellStyle.default; duplicates collapse by
   * canonicalKey) plus the caller-index → emitted-cellXf-index remap to apply to each row's
   * `cellStyles` before emission. Deterministic: emitted order is default + first occurrence.
   */
  def buildStyleTable(styles: Vector[CellStyle]): (OoxmlStyles, Map[Int, Int]) =
    import scala.collection.mutable
    val emitted =
      mutable.LinkedHashMap[String, (Int, CellStyle)](
        CellStyle.default.canonicalKey -> (0, CellStyle.default)
      )
    val remap = styles.zipWithIndex.map { case (style, callerIdx) =>
      val (idx, _) = emitted.getOrElseUpdate(style.canonicalKey, (emitted.size, style))
      callerIdx -> idx
    }.toMap
    val unified = emitted.valuesIterator.map(_._2).toVector

    // Component dedup in first-occurrence order (mirrors StyleIndex.fromWorkbookWithoutSource)
    val fonts = mutable.LinkedHashSet.empty[Font]
    val fills = mutable.LinkedHashSet.empty[Fill]
    val borders = mutable.LinkedHashSet.empty[Border]
    val customCodes = mutable.LinkedHashSet.empty[String]
    unified.foreach { style =>
      fonts += style.font
      fills += style.fill
      borders += style.border
      style.numFmt match
        case NumFmt.Custom(code) => customCodes += code
        case _ => ()
    }
    val customNumFmts = customCodes.toVector.zipWithIndex.map { case (code, idx) =>
      (164 + idx, NumFmt.Custom(code): NumFmt)
    }

    val index = StyleIndex(
      fonts = fonts.toVector,
      fills = fills.toVector,
      borders = borders.toVector,
      numFmts = customNumFmts,
      cellStyles = unified,
      styleToIndex = emitted.view.map { case (key, (idx, _)) => key -> StyleId(idx) }.toMap
    )
    (OoxmlStyles(index), remap)
