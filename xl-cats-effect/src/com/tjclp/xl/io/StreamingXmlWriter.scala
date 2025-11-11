package com.tjclp.xl.io

import fs2.Stream
import fs2.data.xml.*
import fs2.data.xml.XmlEvent.*
import com.tjclp.xl.CellValue

/**
 * True streaming XML writer using fs2-data-xml for constant-memory writes.
 *
 * Emits XML events incrementally as rows arrive, never materializing the full document tree in
 * memory.
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

  /** Generate XML events for a single cell */
  def cellToEvents(colIndex: Int, rowIndex: Int, value: CellValue): List[XmlEvent] =
    val ref = s"${columnToLetter(colIndex)}$rowIndex"

    val (cellType, valueEvents) = value match
      case CellValue.Empty =>
        ("", Nil)

      case CellValue.Text(s) =>
        // <c r="A1" t="inlineStr"><is><t>text</t></is></c>
        // Add xml:space="preserve" for text with leading/trailing/multiple spaces
        val needsPreserve = s.startsWith(" ") || s.endsWith(" ") || s.contains("  ")
        val tAttrs =
          if needsPreserve then
            List(Attr(QName(Some("xml"), "space"), List(XmlString("preserve", false))))
          else Nil

        (
          "inlineStr",
          List(
            XmlEvent.StartTag(QName("is"), Nil, false),
            XmlEvent.StartTag(QName("t"), tAttrs, false),
            XmlEvent.XmlString(s, false),
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

      case CellValue.Formula(expr) =>
        // <c r="A1"><f>SUM(A1:A10)</f></c>
        (
          "",
          List(
            XmlEvent.StartTag(QName("f"), Nil, false),
            XmlEvent.XmlString(expr, false),
            XmlEvent.EndTag(QName("f"))
          )
        )

      case CellValue.Error(err) =>
        import com.tjclp.xl.CellError.toExcel
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

          // Optional <rPr> with font properties
          val rPrEvents = run.font.toList.flatMap { f =>
            val propsBuilder = List.newBuilder[XmlEvent]

            // Font style properties
            if f.bold then propsBuilder += XmlEvent.StartTag(QName("b"), Nil, true)
            if f.italic then propsBuilder += XmlEvent.StartTag(QName("i"), Nil, true)
            if f.underline then propsBuilder += XmlEvent.StartTag(QName("u"), Nil, true)

            // Color
            f.color.foreach { c =>
              propsBuilder += XmlEvent.StartTag(
                QName("color"),
                List(Attr(QName("rgb"), List(XmlString(c.toHex.drop(1), false)))),
                true
              )
            }

            // Size and name
            propsBuilder += XmlEvent.StartTag(
              QName("sz"),
              List(Attr(QName("val"), List(XmlString(f.sizePt.toString, false)))),
              true
            )
            propsBuilder += XmlEvent.StartTag(
              QName("name"),
              List(Attr(QName("val"), List(XmlString(f.name, false)))),
              true
            )

            XmlEvent.StartTag(QName("rPr"), Nil, false) :: propsBuilder.result() ::: List(
              XmlEvent.EndTag(QName("rPr"))
            )
          }

          // Text element with optional xml:space="preserve"
          val tAttrs =
            if run.text.startsWith(" ") || run.text.endsWith(" ") then
              List(Attr(QName(Some("xml"), "space"), List(XmlString("preserve", false))))
            else Nil

          val textEvents = List(
            XmlEvent.StartTag(QName("t"), tAttrs, false),
            XmlEvent.XmlString(run.text, false),
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

    // Build cell element
    val attrs =
      if cellType.nonEmpty then List(attr("r", ref), attr("t", cellType))
      else List(attr("r", ref))

    if valueEvents.isEmpty && value == CellValue.Empty then Nil // Don't emit empty cells
    else
      XmlEvent.StartTag(QName("c"), attrs, valueEvents.isEmpty) ::
        valueEvents :::
        (if valueEvents.nonEmpty then List(XmlEvent.EndTag(QName("c"))) else Nil)

  /** Generate XML events for a row */
  def rowToEvents(rowData: RowData): List[XmlEvent] =
    val rowAttrs = List(attr("r", rowData.rowIndex.toString))

    // Sort cells by column for deterministic output
    val cellEvents = rowData.cells.toList
      .sortBy(_._1)
      .flatMap { case (colIdx, value) =>
        cellToEvents(colIdx, rowData.rowIndex, value)
      }

    if cellEvents.isEmpty then Nil // Skip empty rows
    else
      XmlEvent.StartTag(QName("row"), rowAttrs, false) ::
        cellEvents :::
        XmlEvent.EndTag(QName("row")) :: Nil

  /** Stream worksheet XML events with header/footer scaffolding */
  def worksheetEvents[F[_]](rows: Stream[F, RowData]): Stream[F, XmlEvent] =
    val header = List(
      XmlEvent.StartTag(
        QName("worksheet"),
        List(
          attr("xmlns", nsSpreadsheetML),
          Attr(QName(Some("xmlns"), "r"), List(XmlString(nsRelationships, false)))
        ),
        false
      ),
      XmlEvent.StartTag(QName("sheetData"), Nil, false)
    )

    val footer = List(
      XmlEvent.EndTag(QName("sheetData")),
      XmlEvent.EndTag(QName("worksheet"))
    )

    Stream.emits(header) ++
      rows.flatMap(row => Stream.emits(rowToEvents(row))) ++
      Stream.emits(footer)
