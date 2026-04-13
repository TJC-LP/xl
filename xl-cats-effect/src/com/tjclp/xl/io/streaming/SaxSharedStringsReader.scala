package com.tjclp.xl.io.streaming

import java.io.InputStream
import javax.xml.parsers.SAXParserFactory
import org.xml.sax.{Attributes, InputSource}
import org.xml.sax.helpers.DefaultHandler
import scala.collection.mutable
import scala.xml.Utility
import com.tjclp.xl.ooxml.{SSTEntry, SharedStrings, XmlSecurity, XmlUtil}
import com.tjclp.xl.richtext.{RichText, TextRun}
import com.tjclp.xl.styles.font.Font

/**
 * SAX-based shared-strings reader.
 *
 * Parses xl/sharedStrings.xml incrementally so streaming reads only materialize the final SST
 * entries, not the full XML payload and a DOM tree.
 */
object SaxSharedStringsReader:
  @SuppressWarnings(Array("org.wartremover.warts.Null"))
  private final class ParseFailure(message: String)
      extends RuntimeException(message, null, false, false)

  def parse(stream: InputStream): Either[String, SharedStrings] =
    try
      val factory = SAXParserFactory.newInstance()
      factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true)
      factory.setFeature("http://xml.org/sax/features/external-general-entities", false)
      factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false)
      factory.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false)
      factory.setXIncludeAware(false)
      factory.setNamespaceAware(true)

      val parser = factory.newSAXParser()
      val handler = new SharedStringsHandler()
      parser.parse(InputSource(stream), handler)
      Right(handler.result())
    catch
      case err: ParseFailure => Left(err.getMessage)
      case err: Exception =>
        Left(Option(err.getMessage).getOrElse(err.getClass.getSimpleName))

  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private class SharedStringsHandler extends DefaultHandler:
    private val entries = mutable.ArrayBuffer.empty[SSTEntry]
    private var totalCount: Option[Int] = None

    private var inSi = false
    private var inRun = false
    private var inPhoneticRun = false
    private var inPlainText = false
    private var inRunText = false
    private var runPropsDepth = 0

    private val textBuffer = new StringBuilder
    private val plainTextBuffer = new StringBuilder
    private var plainTextSeen = false

    private val currentRuns = mutable.ArrayBuffer.empty[TextRun]
    private val currentRunText = new StringBuilder
    private var currentRunTextSeen = false
    private val currentRunRPrXml = new StringBuilder
    private var currentRunRawRPr: Option[String] = None

    def result(): SharedStrings =
      if inSi then fail("Unexpected end of xl/sharedStrings.xml while parsing <si>")

      val entryVec = entries.toVector
      val indexMap = entryVec.iterator.zipWithIndex.map { case (entry, idx) =>
        val key = entry match
          case Left(text) => text
          case Right(richText) => richText.toPlainText
        SharedStrings.normalize(key) -> idx
      }.toMap

      SharedStrings(entryVec, indexMap, totalCount.getOrElse(entryVec.size))

    override def startElement(
      uri: String,
      localName: String,
      qName: String,
      attributes: Attributes
    ): Unit =
      val name = elementName(localName, qName)

      if runPropsDepth > 0 then
        runPropsDepth += 1
        appendStartTag(currentRunRPrXml, name, attributes)
      else
        name match
          case "sst" =>
            totalCount = Option(attributes.getValue("count")).flatMap(_.toIntOption)

          case "si" =>
            resetEntry()
            inSi = true

          case "rPh" if inSi =>
            inPhoneticRun = true

          case "r" if inSi && !inPhoneticRun =>
            inRun = true
            currentRunText.clear()
            currentRunTextSeen = false
            currentRunRawRPr = None
            currentRunRPrXml.clear()

          case "rPr" if inRun =>
            runPropsDepth = 1
            currentRunRPrXml.clear()
            appendStartTag(currentRunRPrXml, name, attributes)

          case "t" if inRun =>
            inRunText = true
            textBuffer.clear()

          case "t" if inSi && !inPhoneticRun =>
            inPlainText = true
            textBuffer.clear()

          case _ => ()

    override def characters(ch: Array[Char], start: Int, length: Int): Unit =
      if runPropsDepth > 0 then
        currentRunRPrXml.append(Utility.escape(String.valueOf(ch, start, length)))

      if inPlainText || inRunText then textBuffer.appendAll(ch, start, length)

    override def endElement(uri: String, localName: String, qName: String): Unit =
      val name = elementName(localName, qName)

      if runPropsDepth > 0 then
        appendEndTag(currentRunRPrXml, name)
        runPropsDepth -= 1
        if runPropsDepth == 0 then currentRunRawRPr = Some(currentRunRPrXml.toString)
      else
        name match
          case "t" if inRunText =>
            currentRunText.append(textBuffer.toString)
            currentRunTextSeen = true
            inRunText = false
            textBuffer.clear()

          case "t" if inPlainText =>
            plainTextBuffer.append(textBuffer.toString)
            plainTextSeen = true
            inPlainText = false
            textBuffer.clear()

          case "r" if inRun =>
            if !currentRunTextSeen then fail("Text run <r> missing <t> element")

            val (font, rawRPrXml) = finalizeRunProperties(currentRunRawRPr)
            currentRuns += TextRun(currentRunText.toString, font, rawRPrXml)
            inRun = false

          case "rPh" if inPhoneticRun =>
            inPhoneticRun = false

          case "si" if inSi =>
            entries += buildEntry()
            inSi = false

          case _ => ()

    private def buildEntry(): SSTEntry =
      if currentRuns.nonEmpty && plainTextSeen then
        fail("SharedString <si> cannot mix plain <t> content with rich <r> runs")
      else if currentRuns.nonEmpty then Right(RichText(currentRuns.toVector))
      else if plainTextSeen then Left(plainTextBuffer.toString)
      else fail("SharedString <si> missing <t> element and has no <r> runs")

    private def resetEntry(): Unit =
      inRun = false
      inPhoneticRun = false
      inPlainText = false
      inRunText = false
      runPropsDepth = 0

      textBuffer.clear()
      plainTextBuffer.clear()
      plainTextSeen = false

      currentRuns.clear()
      currentRunText.clear()
      currentRunTextSeen = false
      currentRunRPrXml.clear()
      currentRunRawRPr = None

    private def finalizeRunProperties(rawXml: Option[String]): (Option[Font], Option[String]) =
      rawXml match
        case None => (None, None)
        case Some(xml) =>
          XmlSecurity.parseSafe(xml, "xl/sharedStrings.xml <rPr>").toOption match
            case Some(elem) =>
              (Some(XmlUtil.parseRunProperties(elem)), Some(XmlUtil.compact(elem)))
            case None =>
              (None, Some(xml))

    private def appendStartTag(builder: StringBuilder, name: String, attributes: Attributes): Unit =
      builder.append('<').append(name)
      (0 until attributes.getLength).foreach { idx =>
        builder
          .append(' ')
          .append(attributeName(attributes, idx))
          .append("=\"")
          .append(Utility.escape(attributes.getValue(idx)))
          .append('"')
      }
      builder.append('>')

    private def appendEndTag(builder: StringBuilder, name: String): Unit =
      builder.append("</").append(name).append('>')

    private def attributeName(attributes: Attributes, idx: Int): String =
      val qName = attributes.getQName(idx)
      if qName != null && qName.nonEmpty then qName
      else attributes.getLocalName(idx)

    private def elementName(localName: String, qName: String): String =
      if localName != null && localName.nonEmpty then localName else qName

    private def fail(message: String): Nothing =
      throw new ParseFailure(message)
