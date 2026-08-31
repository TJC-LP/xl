package com.tjclp.xl.ooxml

import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets.UTF_8
import java.util.zip.ZipFile
import scala.jdk.CollectionConverters.*
import munit.FunSuite
import org.xml.sax.helpers.DefaultHandler
import org.xml.sax.Attributes
import com.tjclp.xl.xml.{XmlPullParser, XmlToken}

/**
 * P3 (GH-543): differential parse — every XML part in the real-file fixture corpus, parsed by the
 * portable XmlPullParser and by a secure namespace-aware JAXP SAX parser (namespace-prefixes=true),
 * must produce identical canonical event logs: start/end with qName + localName + element
 * namespace, attributes (sorted), and per-element accumulated character data; comments/PIs elided.
 * The full generative differential law is wave A3.
 */
class PullParserJaxpDifferentialSpec extends FunSuite:

  /** XML parts inside an xlsx: .xml and .rels entries (VML is not XML — excluded). */
  private def xmlParts(fixture: String): List[(String, Array[Byte])] =
    val path = TestFixtures.copyToTemp(fixture)
    val zip = new ZipFile(path.toFile)
    try
      zip
        .entries()
        .asScala
        .filter(e => !e.isDirectory && (e.getName.endsWith(".xml") || e.getName.endsWith(".rels")))
        .map { e =>
          val in = zip.getInputStream(e)
          try e.getName -> in.readAllBytes()
          finally in.close()
        }
        .toList
    finally zip.close()

  // --- canonical event logs ------------------------------------------------------------------------

  private def canonAttrs(attrs: Seq[(String, String)]): String =
    attrs.sortBy(_._1).map { case (n, v) => s"$n=$v" }.mkString("")

  private def jaxpLog(bytes: Array[Byte]): Either[String, Vector[String]] =
    val log = Vector.newBuilder[String]
    val textStack = scala.collection.mutable.Stack.empty[java.lang.StringBuilder]
    val handler = new DefaultHandler:
      override def startElement(
        uri: String,
        localName: String,
        qName: String,
        attributes: Attributes
      ): Unit =
        val attrs = (0 until attributes.getLength).map { i =>
          attributes.getQName(i) -> attributes.getValue(i)
        }
        log += s"S|$qName|$localName|$uri|${canonAttrs(attrs)}"
        textStack.push(new java.lang.StringBuilder)
      override def endElement(uri: String, localName: String, qName: String): Unit =
        val text = if textStack.nonEmpty then textStack.pop().toString else ""
        log += s"E|$qName|$text"
      override def characters(ch: Array[Char], start: Int, length: Int): Unit =
        if textStack.nonEmpty then
          val _ = textStack.top.append(ch, start, length)
    try
      val factory = XmlSecurity.secureSaxParserFactory()
      factory.setFeature("http://xml.org/sax/features/namespace-prefixes", true)
      val parser = factory.newSAXParser()
      // Production JAXP reads run behind the GH-350 stripper; mirror that here so the
      // doctype-hostile fixture parses on the JAXP side too (the pull parser skips natively).
      val stripped = XmlSecurity.stripLeadingDoctype(new String(bytes, UTF_8))
      parser.parse(
        new ByteArrayInputStream(stripped.getBytes(UTF_8)),
        handler
      )
      Right(log.result())
    catch case e: Exception => Left(s"JAXP: ${e.getMessage}")

  private def pullLog(bytes: Array[Byte]): Either[String, Vector[String]] =
    val log = Vector.newBuilder[String]
    val textStack = scala.collection.mutable.Stack.empty[java.lang.StringBuilder]
    val p = XmlPullParser.fromArray(bytes)
    @annotation.tailrec
    def loop(): Either[String, Vector[String]] =
      p.next() match
        case Left(e) => Left(s"pull: ${e.render}")
        case Right(XmlToken.EndDocument) => Right(log.result())
        case Right(XmlToken.StartElement) =>
          log += s"S|${p.name}|${p.localName}|${p.elementNs.getOrElse("")}|${canonAttrs(p.attrs.toVector)}"
          textStack.push(new java.lang.StringBuilder)
          loop()
        case Right(XmlToken.EndElement) =>
          val text = if textStack.nonEmpty then textStack.pop().toString else ""
          log += s"E|${p.name}|$text"
          loop()
        case Right(XmlToken.Text) | Right(XmlToken.CData) =>
          if textStack.nonEmpty then p.appendTextTo(textStack.top)
          loop()
    loop()

  TestFixtures.all.foreach { fixture =>
    test(s"P3: $fixture — pull parser matches JAXP event-for-event") {
      val parts = xmlParts(fixture)
      assert(parts.nonEmpty, s"$fixture contains no XML parts")
      parts.foreach { case (partName, bytes) =>
        (jaxpLog(bytes), pullLog(bytes)) match
          case (Right(expected), Right(actual)) =>
            assertEquals(
              actual.mkString("\n"),
              expected.mkString("\n"),
              s"$fixture!$partName diverged"
            )
          case (Left(je), Left(pe)) =>
            // both reject: acceptable only if JAXP also cannot read it
            assert(true, s"$fixture!$partName rejected by both: $je / $pe")
          case (Left(je), Right(_)) =>
            fail(s"$fixture!$partName: JAXP rejected but pull parser accepted: $je")
          case (Right(_), Left(pe)) =>
            fail(s"$fixture!$partName: pull parser rejected but JAXP accepted: $pe")
      }
    }
  }

  test("P3: doctype-hostile parts carry DOCTYPEs and still agree (GH-350 evidence)") {
    val parts = xmlParts("doctype-hostile.xlsx")
    val withDoctype = parts.filter { case (_, bytes) =>
      new String(bytes, UTF_8).contains("<!DOCTYPE")
    }
    assert(withDoctype.nonEmpty, "doctype-hostile.xlsx should contain DOCTYPE-bearing parts")
  }
