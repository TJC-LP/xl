package com.tjclp.xl.ooxml.worksheet

import scala.xml.*

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.ooxml.SaxSupport.*
import com.tjclp.xl.ooxml.XmlUtil.{elem, elemOrdered, needsXmlSpacePreserve}
import com.tjclp.xl.ooxml.{SaxWriter, XmlSecurity, XmlUtil}
import com.tjclp.xl.styles.color.Color

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
          writer.startElement("is")
          writer.startElement("t")
          if needsXmlSpacePreserve(text) then writer.writeAttribute("xml:space", "preserve")
          writer.writeCharacters(text)
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
          writer.writeCharacters(num.toString)
          writer.endElement() // v

        case CellValue.Bool(b) =>
          writer.startElement("v")
          writer.writeCharacters(if b then "1" else "0")
          writer.endElement() // v

        case CellValue.Formula(expr, cachedValue) =>
          writer.startElement("f")
          writer.writeCharacters(expr)
          writer.endElement() // f
          // Write cached value if present
          cachedValue.foreach {
            case CellValue.Number(num) =>
              writer.startElement("v")
              writer.writeCharacters(num.toString)
              writer.endElement()
            case CellValue.Text(s) =>
              writer.startElement("v")
              writer.writeCharacters(s)
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
            case _ => () // Empty, RichText, Formula, DateTime - don't write
          }

        case CellValue.Error(err) =>
          import com.tjclp.xl.cells.CellError.toExcel
          writer.startElement("v")
          writer.writeCharacters(err.toExcel)
          writer.endElement() // v

        case CellValue.DateTime(dt) =>
          val serial = CellValue.dateTimeToExcelSerial(dt)
          writer.startElement("v")
          writer.writeCharacters(serial.toString)
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

      writer.startElement("t")
      if needsXmlSpacePreserve(run.text) then writer.writeAttribute("xml:space", "preserve")
      writer.writeCharacters(run.text)
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
    if font.underline then
      writer.startElement("u")
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

    writer.startElement("name")
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
                if f.underline then fontProps += elem("u")()

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
      case CellValue.Formula(expr, cachedValue) =>
        // Write formula element
        val formulaElem = elem("f")(Text(expr))
        // Write cached value if present
        val cachedElem = cachedValue.flatMap {
          case CellValue.Number(num) => Some(elem("v")(Text(num.toString)))
          case CellValue.Text(s) => Some(elem("v")(Text(s)))
          case CellValue.Bool(b) => Some(elem("v")(Text(if b then "1" else "0")))
          case CellValue.Error(err) =>
            import com.tjclp.xl.cells.CellError.toExcel
            Some(elem("v")(Text(err.toExcel)))
          case _ => None // Empty, RichText, Formula, DateTime - don't write
        }
        Seq(formulaElem) ++ cachedElem.toList
      case CellValue.Error(err) =>
        import com.tjclp.xl.cells.CellError.toExcel
        Seq(elem("v")(Text(err.toExcel)))
      case CellValue.DateTime(dt) =>
        // DateTime is serialized as number with Excel serial format
        val serial = CellValue.dateTimeToExcelSerial(dt)
        Seq(elem("v")(Text(serial.toString)))

    elemOrdered("c", finalAttrs*)(valueElem*)
