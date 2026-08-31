package com.tjclp.xl.ooxml

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets.UTF_8
import org.scalacheck.Gen
import org.scalacheck.Prop.forAll

/**
 * W1 (GH-543): adapter parity — PortableSaxWriter must be byte-identical to StaxSaxWriter for every
 * SaxWriter protocol xl-ooxml's emitters drive: xmlns through writeAttribute,
 * SaxWriter.withAttributes ordering, nested redeclarations, knownNamespaces auto-declaration, fresh
 * namespaces, xml:*, emptyElement minimization, the GH-237 sanitize path, and watermark-idempotent
 * flushing.
 */
class PortableSaxWriterParitySpec extends munit.ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(200)

  private def both(body: SaxWriter => Unit): (Array[Byte], Array[Byte]) =
    val staxOut = new ByteArrayOutputStream()
    val stax = StaxSaxWriter.create(staxOut)
    body(stax)
    stax.flush()
    val portableOut = new ByteArrayOutputStream()
    val portable = PortableSaxWriter.create(portableOut)
    body(portable)
    portable.flush()
    (staxOut.toByteArray, portableOut.toByteArray)

  private def assertParity(label: String)(body: SaxWriter => Unit): Unit =
    val (stax, portable) = both(body)
    assertEquals(new String(portable, UTF_8), new String(stax, UTF_8), label)
    assert(java.util.Arrays.equals(portable, stax), s"$label: byte-level divergence")

  // --- hand cases (ported from StaxSaxWriterNamespaceSpec + the probed quirks) --------------------

  test("W1: prefixed attribute under an explicit xmlns declaration") {
    assertParity("x14ac decl + prefixed attr") { w =>
      w.startDocument()
      w.startElement("root")
      w.writeAttribute("xmlns:x14ac", XmlUtil.nsX14ac)
      w.startElement("child")
      w.writeAttribute("x14ac:dyDescent", "0.25")
      w.endElement()
      w.endElement()
      w.endDocument()
    }
  }

  test("W1: knownNamespaces prefix auto-declares once (r:id)") {
    assertParity("r:id auto declaration") { w =>
      w.startDocument()
      w.startElement("worksheet", XmlUtil.nsSpreadsheetML)
      w.startElement("sheet")
      w.writeAttribute("r:id", "rId1")
      w.endElement()
      w.startElement("sheet")
      w.writeAttribute("r:id", "rId2") // second use: already bound, no duplicate decl
      w.endElement()
      w.endElement()
      w.endDocument()
    }
  }

  test("W1: namespaced root via startElement(name, namespace)") {
    assertParity("default-ns root") { w =>
      w.startDocument()
      w.startElement("styleSheet", XmlUtil.nsSpreadsheetML)
      w.startElement("fonts")
      w.writeAttribute("count", "1")
      w.endElement()
      w.endElement()
      w.endDocument()
    }
  }

  test("W1: prefixed root via startElement(name, namespace)") {
    assertParity("prefixed root") { w =>
      w.startDocument()
      w.startElement("c:chartSpace", XmlUtil.nsChart)
      w.startElement("c:chart")
      w.endElement()
      w.endElement()
      w.endDocument()
    }
  }

  test("W1: nested redeclaration of the same prefix restores on pop") {
    assertParity("nested redeclaration") { w =>
      w.startDocument()
      w.startElement("a")
      w.writeAttribute("xmlns:p", "urn:one")
      w.writeAttribute("p:x", "1")
      w.startElement("b")
      w.writeAttribute("xmlns:p", "urn:two")
      w.writeAttribute("p:x", "2")
      w.endElement()
      w.startElement("c")
      w.writeAttribute("p:x", "3") // p restored to urn:one — no new declaration
      w.endElement()
      w.endElement()
      w.endDocument()
    }
  }

  test("W1: same-URI redeclaration is skipped (alreadyBound)") {
    assertParity("alreadyBound skip") { w =>
      w.startDocument()
      w.startElement("a")
      w.writeAttribute("xmlns:p", "urn:same")
      w.startElement("b")
      w.writeAttribute("xmlns:p", "urn:same")
      w.writeAttribute("p:x", "1")
      w.endElement()
      w.endElement()
      w.endDocument()
    }
  }

  test("W1: xml:space needs no declaration") {
    assertParity("xml:space") { w =>
      w.startDocument()
      w.startElement("si")
      w.startElement("t")
      w.writeAttribute("xml:space", "preserve")
      w.writeCharacters("  padded  ")
      w.endElement()
      w.endElement()
      w.endDocument()
    }
  }

  test("W1: default namespace via writeAttribute(xmlns)") {
    assertParity("default ns attr") { w =>
      w.startDocument()
      w.startElement("Types")
      SaxWriter.withAttributes(w, "xmlns" -> XmlUtil.nsContentTypes) {
        w.startElement("Default")
        SaxWriter.withAttributes(w, "Extension" -> "xml", "ContentType" -> "application/xml")(())
        w.endElement()
      }
      w.endElement()
      w.endDocument()
    }
  }

  test("W1: withAttributes puts xmlns first then sorts (protocol used by every part)") {
    assertParity("withAttributes ordering") { w =>
      w.startDocument()
      w.startElement("worksheet")
      SaxWriter.withAttributes(
        w,
        "zeta" -> "1",
        "xmlns:r" -> XmlUtil.nsRelationships,
        "alpha" -> "2",
        "xmlns" -> XmlUtil.nsSpreadsheetML
      ) {
        w.startElement("d")
        w.writeAttribute("r:id", "rId9")
        w.endElement()
      }
      w.endElement()
      w.endDocument()
    }
  }

  test("W1: emptyElement minimizes with attributes in order (GH-430)") {
    assertParity("emptyElement") { w =>
      w.startDocument()
      w.startElement("row")
      w.emptyElement("f", Seq("t" -> "dataTable", "ref" -> "B2:C3", "dt2D" -> "1"))
      w.emptyElement("v", Seq.empty)
      w.endElement()
      w.endDocument()
    }
  }

  test("W1: emptyElement with xmlns attribute") {
    assertParity("emptyElement xmlns") { w =>
      w.startDocument()
      w.startElement("root")
      w.emptyElement(
        "ext",
        Seq(
          "xmlns:x14" -> "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
          "uri" -> "{X}"
        )
      )
      w.endElement()
      w.endDocument()
    }
  }

  test("W1: GH-237 sanitize strips XML-illegal control characters identically") {
    assertParity("sanitize") { w =>
      w.startDocument()
      w.startElement("t")
      w.writeCharacters("ok \u0000bad\u0008worse\u001fdone\ttab\nnl kept")
      w.endElement()
      w.endDocument()
    }
  }

  test("W1: escape matrix in text and attributes") {
    assertParity("escapes") { w =>
      w.startDocument()
      w.startElement("a")
      w.writeAttribute("v", """q"uo & <lt> 'ap'""")
      w.writeCharacters("""1<2 & "x" 'y' 3>2 ]]>""")
      w.endElement()
      w.endDocument()
    }
  }

  test("W1: unicode and astral content") {
    assertParity("unicode") { w =>
      w.startDocument()
      w.startElement("t")
      w.writeAttribute("n", "naïve 日本語")
      w.writeCharacters("héllo wörld 🚀 日本語テキスト €")
      w.endElement()
      w.endDocument()
    }
  }

  test("W1: double flush does not duplicate bytes (Deflated worksheet path flushes twice)") {
    val staxOut = new ByteArrayOutputStream()
    val stax = StaxSaxWriter.create(staxOut)
    val portableOut = new ByteArrayOutputStream()
    val portable = PortableSaxWriter.create(portableOut)
    def phase1(w: SaxWriter): Unit =
      w.startDocument()
      w.startElement("worksheet")
      w.startElement("sheetData")
      w.writeCharacters("first")
      w.flush()
    def phase2(w: SaxWriter): Unit =
      w.writeCharacters("second")
      w.endElement()
      w.endElement()
      w.endDocument()
      w.flush()
      w.flush() // third flush with nothing new: must add zero bytes
    phase1(stax); phase2(stax)
    phase1(portable); phase2(portable)
    assertEquals(new String(portableOut.toByteArray, UTF_8), new String(staxOut.toByteArray, UTF_8))
    val s = new String(portableOut.toByteArray, UTF_8)
    assertEquals(s.sliding("first".length).count(_ == "first"), 1, s)
    assertEquals(s.sliding("second".length).count(_ == "second"), 1, s)
  }

  test("W1: writeElem generic walker (Table/Comments/Chart/Drawing protocol shape)") {
    import com.tjclp.xl.ooxml.SaxSupport.*
    val elem = scala.xml.XML.loadString(
      """<table xmlns="urn:t" id="1" displayName="T1"><autoFilter ref="A1:B2"/>""" +
        """<tableColumns count="2"><tableColumn id="1" name="A &amp; B"/>""" +
        """<tableColumn id="2" name="süß"/></tableColumns></table>"""
    )
    assertParity("writeElem") { w =>
      w.startDocument()
      w.writeElem(elem)
      w.endDocument()
      w.flush()
    }
  }

  // --- generative protocol parity -------------------------------------------------------------------

  private enum POp derives CanEqual:
    case Start(name: String)
    case StartNs(name: String, ns: String)
    case Attr(name: String, value: String)
    case Chars(text: String)
    case Empty(name: String, attrs: List[(String, String)])
    case End

  private val genPlainName: Gen[String] =
    for
      h <- Gen.alphaChar
      t <- Gen.listOfN(2, Gen.alphaNumChar)
    yield (h :: t).mkString

  private val genPrefix: Gen[String] =
    Gen.oneOf("r", "mc", "x14ac", "xm", "p", "q", "xml")

  private val genName: Gen[String] =
    Gen.frequency(
      4 -> genPlainName,
      2 -> genPrefix.flatMap(p => genPlainName.map(n => s"$p:$n"))
    )

  private val genUri: Gen[String] =
    Gen.oneOf("urn:one", "urn:two", XmlUtil.nsSpreadsheetML, XmlUtil.nsRelationships)

  private val genValue: Gen[String] =
    Gen
      .listOf(
        Gen.frequency(
          8 -> Gen.alphaNumChar.map(_.toString),
          3 -> Gen.oneOf("&", "<", ">", "\"", "'", " ", "\t", "\n", "\r"),
          1 -> Gen.oneOf("é", "日", "🚀")
        )
      )
      .map(_.mkString)

  private val genAttr: Gen[POp] =
    Gen.frequency(
      5 -> genName.flatMap(n => genValue.map(v => POp.Attr(n, v))),
      1 -> genUri.map(u => POp.Attr("xmlns", u)),
      2 -> genPrefix
        .suchThat(_ != "xml")
        .flatMap(p => genUri.map(u => POp.Attr(s"xmlns:$p", u)))
    )

  /**
   * Production emitters route element attributes through SaxWriter.withAttributes, which writes
   * xmlns declarations FIRST (sorted), then the rest (sorted) — a prefixed attribute may otherwise
   * auto-declare a known prefix that a later explicit xmlns on the SAME element would rebind, which
   * StAX rejects. Mirror that protocol invariant; `exclude` drops declarations for a prefix the
   * startElement(name, ns) form already force-declared on this element.
   */
  private def protocolOrder(attrs: List[POp], exclude: Option[String]): List[POp] =
    val unique = attrs
      .collect { case a: POp.Attr => a }
      .distinctBy(_.name)
      .filterNot { a =>
        exclude.contains(a.name) // "xmlns" or "xmlns:p" force-declared by startElement
      }
    val (xmlns, other) = unique.partition(a => a.name == "xmlns" || a.name.startsWith("xmlns:"))
    xmlns.sortBy(_.name) ++ other.sortBy(_.name)

  private def genBody(depth: Int): Gen[List[POp]] =
    if depth <= 0 then Gen.const(Nil)
    else
      Gen.choose(0, 3).flatMap { n =>
        Gen
          .listOfN(
            n,
            Gen.frequency(
              3 -> Gen.lzy(genTree(depth - 1)),
              2 -> genValue.map(v => List(POp.Chars(v))),
              1 -> genName.flatMap(nm =>
                Gen
                  .listOf(genAttr)
                  .map(as =>
                    List(
                      POp.Empty(
                        nm,
                        protocolOrder(as, None).collect { case POp.Attr(an, av) => (an, av) }
                      )
                    )
                  )
              )
            )
          )
          .map(_.flatten)
      }

  private def genTree(depth: Int): Gen[List[POp]] =
    for
      name <- genName
      withNs <- Gen.frequency(3 -> Gen.const(None), 1 -> genUri.map(Some(_)))
      rawAttrs <- Gen.listOf(genAttr)
      body <- genBody(depth)
    yield
      val start = withNs match
        case Some(ns) => POp.StartNs(name, ns)
        case None => POp.Start(name)
      // startElement(name, ns) force-declares the element's own prefix (or the default ns):
      // exclude explicit re-declarations of that prefix on the same element (out of protocol).
      val forced = withNs.map { _ =>
        name.split(":", 2) match
          case Array(p, _) => s"xmlns:$p"
          case _ => "xmlns"
      }
      val attrs = protocolOrder(rawAttrs, forced)
      (start :: attrs) ++ body ++ List(POp.End)

  private def drive(w: SaxWriter, ops: List[POp]): Unit =
    w.startDocument()
    ops.foreach {
      case POp.Start(n) => w.startElement(n)
      case POp.StartNs(n, ns) => w.startElement(n, ns)
      case POp.Attr(n, v) => w.writeAttribute(n, v)
      case POp.Chars(t) => w.writeCharacters(t)
      case POp.Empty(n, as) => w.emptyElement(n, as)
      case POp.End => w.endElement()
    }
    w.endDocument()
    w.flush()

  property("W1: generative SaxWriter protocol sequences are byte-identical") {
    forAll(genTree(3)) { ops =>
      assertParity(ops.toString)(w => drive(w, ops))
      true
    }
  }
