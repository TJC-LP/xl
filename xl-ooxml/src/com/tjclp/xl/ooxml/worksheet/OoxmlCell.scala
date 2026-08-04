package com.tjclp.xl.ooxml.worksheet

import scala.xml.*

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.ooxml.SaxSupport.*
import com.tjclp.xl.ooxml.XmlUtil.{elem, elemOrdered, needsXmlSpacePreserve}
import com.tjclp.xl.ooxml.{FormulaKindCodec, SaxWriter, XmlSecurity, XmlUtil}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.font.Underline

/** Cell data for worksheet - maps domain Cell to XML representation */
case class OoxmlCell(
  ref: ARef,
  value: CellValue,
  styleIndex: Option[Int] = None,
  cellType: String = "inlineStr" // "s" for SST, "inlineStr" for inline, "n" for number, etc.
):
  def toA1: String = ref.toA1

  def writeSax(writer: SaxWriter): Unit =
    writer.startElement("c")

    val attrs = Seq.newBuilder[(String, String)]
    attrs += ("r" -> toA1)
    styleIndex.foreach(s => attrs += ("s" -> s.toString))
    if cellType.nonEmpty then attrs += ("t" -> cellType)

    SaxWriter.withAttributes(writer, attrs.result()*) {
      value match
        case CellValue.Empty => ()

        case CellValue.Text(text) if cellType == "inlineStr" =>
          val escaped = XmlUtil.escapeXstring(text)
          writer.startElement("is")
          writer.startElement("t")
          if needsXmlSpacePreserve(escaped) then writer.writeAttribute("xml:space", "preserve")
          writer.writeCharacters(escaped)
          writer.endElement() // t
          writer.endElement() // is

        case CellValue.Text(text) =>
          writer.startElement("v")
          writer.writeCharacters(text)
          writer.endElement() // v

        case CellValue.RichText(richText) =>
          writeRichTextSax(writer, richText)

        case CellValue.Number(num) =>
          writer.startElement("v")
          writer.writeCharacters(XmlUtil.plainNumber(num))
          writer.endElement() // v

        case CellValue.Bool(b) =>
          writer.startElement("v")
          writer.writeCharacters(if b then "1" else "0")
          writer.endElement() // v

        case CellValue.Formula(expr, cachedValue, kind) =>
          // GH-430: record attrs (t/ref/dt2D/dtr/r1/r2/...) render via the shared codec; a
          // dataTable record carries no formula text (self-closing <f/>).
          kind match
            case _: FormulaKind.DataTable =>
              writer.emptyElement("f", FormulaKindCodec.toAttrs(kind))
            case _ =>
              writer.startElement("f")
              FormulaKindCodec.toAttrs(kind).foreach { case (name, v) =>
                writer.writeAttribute(name, v)
              }
              // GH-456: <f> carries the expression, never the display form's leading '='
              writer.writeCharacters(expr.stripPrefix("="))
              writer.endElement() // f
          // Write cached value if present
          cachedValue.foreach {
            case CellValue.Number(num) =>
              writer.startElement("v")
              writer.writeCharacters(XmlUtil.plainNumber(num))
              writer.endElement()
            case CellValue.Text(s) =>
              writer.startElement("v")
              writer.writeCharacters(XmlUtil.escapeXstring(s))
              writer.endElement()
            case CellValue.Bool(b) =>
              writer.startElement("v")
              writer.writeCharacters(if b then "1" else "0")
              writer.endElement()
            case CellValue.Error(err) =>
              import com.tjclp.xl.cells.CellError.toExcel
              writer.startElement("v")
              writer.writeCharacters(err.toExcel)
              writer.endElement()
            case CellValue.DateTime(dt) =>
              // GH-378: cached DateTime serializes as the Excel serial (t="n"), like Excel itself
              writer.startElement("v")
              writer.writeCharacters(XmlUtil.plainNumber(CellValue.dateTimeToExcelSerial(dt)))
              writer.endElement()
            case _ => () // Empty, RichText, Formula - don't write
          }

        case CellValue.Error(err) =>
          import com.tjclp.xl.cells.CellError.toExcel
          writer.startElement("v")
          writer.writeCharacters(err.toExcel)
          writer.endElement() // v

        case CellValue.DateTime(dt) =>
          val serial = CellValue.dateTimeToExcelSerial(dt)
          writer.startElement("v")
          writer.writeCharacters(XmlUtil.plainNumber(serial))
          writer.endElement() // v
    }

    writer.endElement() // c

  private def writeRichTextSax(writer: SaxWriter, richText: com.tjclp.xl.richtext.RichText): Unit =
    writer.startElement("is")

    richText.runs.foreach { run =>
      writer.startElement("r")

      // Write rPr either from preserved raw XML or constructed from Font
      val preservedRpr = run.rawRPrXml.flatMap { xmlString =>
        XmlSecurity
          .parseSafe(xmlString, "worksheet richtext rPr")
          .toOption
          .map(XmlUtil.stripNamespaces)
      }

      preservedRpr match
        case Some(elem) =>
          writer.writeElem(elem)
        case None =>
          run.font.foreach(writeFontRPrSax(writer, _))

      val runText = XmlUtil.escapeXstring(run.text)
      writer.startElement("t")
      if needsXmlSpacePreserve(runText) then writer.writeAttribute("xml:space", "preserve")
      writer.writeCharacters(runText)
      writer.endElement() // t

      writer.endElement() // r
    }

    writer.endElement() // is

  private def writeFontRPrSax(writer: SaxWriter, font: com.tjclp.xl.styles.Font): Unit =
    writer.startElement("rPr")
    if font.bold then
      writer.startElement("b")
      writer.endElement()
    if font.italic then
      writer.startElement("i")
      writer.endElement()
    if font.underline != Underline.None then
      writer.startElement("u")
      if font.underline != Underline.Single then
        writer.writeAttribute("val", Underline.token(font.underline))
      writer.endElement()

    font.color.foreach {
      case Color.Rgb(argb) =>
        writer.startElement("color")
        writer.writeAttribute("rgb", f"$argb%08X")
        writer.endElement()
      case Color.Theme(slot, tint) =>
        writer.startElement("color")
        writer.writeAttribute("theme", slot.ordinal.toString)
        writer.writeAttribute("tint", tint.toString)
        writer.endElement()
    }

    writer.startElement("sz")
    writer.writeAttribute("val", font.sizePt.toString)
    writer.endElement()

    // CT_RPrElt spells the font element <rFont>, not the CT_Font <name> (GH-383)
    writer.startElement("rFont")
    writer.writeAttribute("val", font.name)
    writer.endElement()

    writer.endElement() // rPr

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
        val escaped = XmlUtil.sanitizeXmlText(XmlUtil.escapeXstring(text))
        val needsPreserve = needsXmlSpacePreserve(escaped)
        val tElem =
          if needsPreserve then
            Elem(
              null,
              "t",
              PrefixedAttribute("xml", "space", "preserve", Null),
              TopScope,
              true,
              Text(escaped)
            )
          else elem("t")(Text(escaped))
        Seq(elem("is")(tElem))
      case CellValue.Text(text) => // SST index
        Seq(
          elem("v")(Text(XmlUtil.sanitizeXmlText(text)))
        ) // text here would be the SST index as string
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
                f.underline match
                  case Underline.None => ()
                  case Underline.Single => fontProps += elem("u")()
                  case other => fontProps += elem("u", "val" -> Underline.token(other))()

                // Font color
                f.color.foreach {
                  case Color.Rgb(argb) =>
                    fontProps += elem("color", "rgb" -> f"$argb%08X")()
                  case Color.Theme(slot, tint) =>
                    fontProps += elem(
                      "color",
                      "theme" -> slot.ordinal.toString,
                      "tint" -> tint.toString
                    )()
                }

                // Font size and name
                fontProps += elem("sz", "val" -> f.sizePt.toString)()
                // CT_RPrElt spells the font element <rFont>, not the CT_Font <name> (GH-383)
                fontProps += elem("rFont", "val" -> f.name)()

                elem("rPr")(fontProps.result()*)
              }.toList

          // Text run: <r> with optional <rPr> and <t>
          // Add xml:space="preserve" to preserve leading/trailing/multiple spaces
          val runText = XmlUtil.sanitizeXmlText(XmlUtil.escapeXstring(run.text))
          val textElem =
            if needsXmlSpacePreserve(runText) then
              Elem(
                null,
                "t",
                PrefixedAttribute("xml", "space", "preserve", Null),
                TopScope,
                true,
                Text(runText)
              )
            else elem("t")(Text(runText))

          elem("r")(
            rPrElems ++ Seq(textElem)*
          )
        }

        Seq(elem("is")(runElems*))
      case CellValue.Number(num) =>
        Seq(elem("v")(Text(XmlUtil.plainNumber(num))))
      case CellValue.Bool(b) =>
        Seq(elem("v")(Text(if b then "1" else "0")))
      case CellValue.Formula(expr, cachedValue, kind) =>
        // Write formula element. GH-430: record attrs via the shared codec in schema order
        // (elemOrdered keeps the given order); dataTable records carry no text.
        val recordAttrs = FormulaKindCodec.toAttrs(kind)
        // GH-456: <f> carries the expression, never the display form's leading '='
        val formulaElem = kind match
          case _: FormulaKind.DataTable => elemOrdered("f", recordAttrs*)()
          case _ => elemOrdered("f", recordAttrs*)(Text(expr.stripPrefix("=")))
        // Write cached value if present
        val cachedElem = cachedValue.flatMap {
          case CellValue.Number(num) => Some(elem("v")(Text(XmlUtil.plainNumber(num))))
          case CellValue.Text(s) => Some(elem("v")(Text(XmlUtil.escapeXstring(s))))
          case CellValue.Bool(b) => Some(elem("v")(Text(if b then "1" else "0")))
          case CellValue.Error(err) =>
            import com.tjclp.xl.cells.CellError.toExcel
            Some(elem("v")(Text(err.toExcel)))
          case CellValue.DateTime(dt) =>
            // GH-378: cached DateTime serializes as the Excel serial (t="n"), like Excel itself
            Some(elem("v")(Text(XmlUtil.plainNumber(CellValue.dateTimeToExcelSerial(dt)))))
          case _ => None // Empty, RichText, Formula - don't write
        }
        Seq(formulaElem) ++ cachedElem.toList
      case CellValue.Error(err) =>
        import com.tjclp.xl.cells.CellError.toExcel
        Seq(elem("v")(Text(err.toExcel)))
      case CellValue.DateTime(dt) =>
        // DateTime is serialized as number with Excel serial format
        val serial = CellValue.dateTimeToExcelSerial(dt)
        Seq(elem("v")(Text(XmlUtil.plainNumber(serial))))

    elemOrdered("c", finalAttrs*)(valueElem*)
