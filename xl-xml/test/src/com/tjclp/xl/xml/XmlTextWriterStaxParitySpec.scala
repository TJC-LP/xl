package com.tjclp.xl.xml

import java.io.ByteArrayOutputStream
import javax.xml.stream.{XMLOutputFactory, XMLStreamWriter}
import org.scalacheck.Gen
import org.scalacheck.Prop.forAll

/**
 * W0: raw writer parity — XmlTextWriter must be byte-identical to the JDK StAX writer built exactly
 * as StaxSaxWriter.create builds it (IS_REPAIRING_NAMESPACES=false, "UTF-8"). JVM-only oracle test
 * (moves to the test-jvm leg in the B-track cross split).
 */
class XmlTextWriterStaxParitySpec extends munit.ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(200)

  // --- op-sequence model ------------------------------------------------------------------------

  private enum Op derives CanEqual:
    case Start(name: String)
    case Empty(name: String)
    case Attr(name: String, value: String)
    case Chars(text: String)
    case End

  private def driveOurs(ops: List[Op]): Array[Byte] =
    val w = new XmlTextWriter
    w.writeStartDocument()
    ops.foreach {
      case Op.Start(n) => w.writeStartElement(n)
      case Op.Empty(n) => w.writeEmptyElement(n)
      case Op.Attr(n, v) => w.writeAttribute(n, v)
      case Op.Chars(t) => w.writeCharacters(t)
      case Op.End => w.writeEndElement()
    }
    w.writeEndDocument()
    w.toBytes

  private def driveStax(ops: List[Op]): Array[Byte] =
    val baos = new ByteArrayOutputStream()
    val factory = XMLOutputFactory.newInstance()
    factory.setProperty(XMLOutputFactory.IS_REPAIRING_NAMESPACES, false)
    val w: XMLStreamWriter = factory.createXMLStreamWriter(baos, "UTF-8")
    w.writeStartDocument("UTF-8", "1.0")
    ops.foreach {
      case Op.Start(n) => w.writeStartElement(n)
      case Op.Empty(n) => w.writeEmptyElement(n)
      case Op.Attr(n, v) => w.writeAttribute(n, v)
      case Op.Chars(t) => w.writeCharacters(t)
      case Op.End => w.writeEndElement()
    }
    w.writeEndDocument()
    w.flush()
    baos.toByteArray

  private def assertParity(ops: List[Op]): Unit =
    val ours = driveOurs(ops)
    val stax = driveStax(ops)
    assertEquals(
      new String(ours, "UTF-8"),
      new String(stax, "UTF-8"),
      s"ops: $ops"
    )
    assert(java.util.Arrays.equals(ours, stax), s"byte-level divergence on ops: $ops")

  // --- generators -------------------------------------------------------------------------------------

  private val genName: Gen[String] =
    for
      h <- Gen.alphaChar
      t <- Gen.listOfN(2, Gen.alphaNumChar)
      prefixed <- Gen.oneOf(true, false, false)
      p <- Gen.alphaChar
    yield if prefixed then s"$p${t.mkString}:$h${t.mkString}" else (h :: t).mkString

  private val genValueChar: Gen[String] =
    Gen.frequency(
      8 -> Gen.alphaNumChar.map(_.toString),
      4 -> Gen.oneOf("&", "<", ">", "\"", "'", "\t", "\n", "\r", " "),
      2 -> Gen.oneOf("é", "日", "€", "😀")
    )

  private val genValue: Gen[String] = Gen.listOf(genValueChar).map(_.mkString)

  private def genOps(depth: Int): Gen[List[Op]] =
    if depth <= 0 then Gen.const(Nil)
    else
      Gen.choose(0, 3).flatMap { n =>
        Gen
          .listOfN(
            n,
            Gen.frequency(
              3 -> genLeafOrTree(depth),
              2 -> genValue.map(v => List(Op.Chars(v)))
            )
          )
          .map(_.flatten)
      }

  private def genLeafOrTree(depth: Int): Gen[List[Op]] =
    for
      name <- genName
      nAttrs <- Gen.choose(0, 3)
      attrNames <- Gen.listOfN(nAttrs, genName).map(_.distinct)
      attrs <- Gen.sequence[List[Op], Op](
        attrNames.map(n => genValue.map(v => Op.Attr(n, v): Op))
      )
      empty <- Gen.oneOf(true, false)
      body <- if empty then Gen.const(Nil) else Gen.lzy(genOps(depth - 1))
    yield
      if empty then Op.Empty(name) :: attrs
      else (Op.Start(name) :: attrs) ++ body ++ List(Op.End)

  private val genDoc: Gen[List[Op]] =
    for
      root <- genName
      nAttrs <- Gen.choose(0, 2)
      attrNames <- Gen.listOfN(nAttrs, genName).map(_.distinct)
      attrs <- Gen.sequence[List[Op], Op](
        attrNames.map(n => genValue.map(v => Op.Attr(n, v): Op))
      )
      body <- genOps(3)
    yield (Op.Start(root) :: attrs) ++ body ++ List(Op.End)

  // --- generative parity --------------------------------------------------------------------------------

  property("W0: generative op sequences produce identical bytes") {
    forAll(genDoc) { ops =>
      assertParity(ops)
      true
    }
  }

  // --- fixed goldens for the probed quirks --------------------------------------------------------------

  test("declaration bytes: no standalone, no newline") {
    val w = new XmlTextWriter
    w.writeStartDocument()
    assertEquals(new String(w.toBytes, "UTF-8"), "<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
  }

  test("start/end with no children is NEVER minimized: <a></a>") {
    val ops = List(Op.Start("a"), Op.End)
    assertEquals(new String(driveOurs(ops), "UTF-8").drop(38), "<a></a>")
    assertParity(ops)
  }

  test("writeEmptyElement minimizes: <a b=\"c\"/> closed lazily by the next event (GH-430)") {
    val ops = List(Op.Start("r"), Op.Empty("f"), Op.Attr("t", "dataTable"), Op.Empty("g"), Op.End)
    assertEquals(
      new String(driveOurs(ops), "UTF-8").drop(38),
      "<r><f t=\"dataTable\"/><g/></r>"
    )
    assertParity(ops)
  }

  test("text escapes ONLY & < >; quotes pass raw") {
    val ops = List(Op.Start("a"), Op.Chars("""1<2 & "x" 'y' 3>2"""), Op.End)
    assertEquals(
      new String(driveOurs(ops), "UTF-8").drop(38),
      "<a>1&lt;2 &amp; \"x\" 'y' 3&gt;2</a>"
    )
    assertParity(ops)
  }

  test("attr escapes & < > \"; apostrophe passes raw") {
    val ops = List(Op.Start("a"), Op.Attr("x", """1<2 & "q" 'y' 3>2"""), Op.End)
    assertEquals(
      new String(driveOurs(ops), "UTF-8").drop(38),
      "<a x=\"1&lt;2 &amp; &quot;q&quot; 'y' 3&gt;2\"></a>"
    )
    assertParity(ops)
  }

  test("hex assertion: CR stays byte 0x0D in text and attribute values") {
    val ops = List(Op.Start("a"), Op.Attr("x", "p\rq"), Op.Chars("u\rv"), Op.End)
    val ours = driveOurs(ops)
    val crCount = ours.count(_ == 0x0d.toByte)
    assertEquals(crCount, 2, ours.map(b => f"$b%02x").mkString(" "))
    assertParity(ops)
  }

  test("tab and LF pass through raw in text and attributes") {
    val ops = List(Op.Start("a"), Op.Attr("x", "p\tq\nr"), Op.Chars("s\tt\nu"), Op.End)
    val s = new String(driveOurs(ops), "UTF-8").drop(38)
    assertEquals(s, "<a x=\"p\tq\nr\">s\tt\nu</a>")
    assertParity(ops)
  }

  test("']]>' is inert in text because '>' always escapes") {
    val ops = List(Op.Start("a"), Op.Chars("x]]>y"), Op.End)
    assertEquals(new String(driveOurs(ops), "UTF-8").drop(38), "<a>x]]&gt;y</a>")
    assertParity(ops)
  }

  test("non-ASCII is raw UTF-8 including astral planes") {
    val ops = List(Op.Start("a"), Op.Attr("x", "naïve 日本語"), Op.Chars("😀 €"), Op.End)
    assertParity(ops)
    val bytes = driveOurs(ops)
    // U+1F600 as UTF-8: F0 9F 98 80
    val s = bytes.map(b => f"${b & 0xff}%02x").mkString(" ")
    assert(s.contains("f0 9f 98 80"), s)
  }

  test("writeEndDocument auto-closes all open elements") {
    val w = new XmlTextWriter
    w.writeStartDocument()
    w.writeStartElement("a")
    w.writeStartElement("b")
    w.writeCharacters("x")
    w.writeEndDocument()
    assertEquals(new String(w.toBytes, "UTF-8").drop(38), "<a><b>x</b></a>")

    val baos = new ByteArrayOutputStream()
    val factory = XMLOutputFactory.newInstance()
    factory.setProperty(XMLOutputFactory.IS_REPAIRING_NAMESPACES, false)
    val sw = factory.createXMLStreamWriter(baos, "UTF-8")
    sw.writeStartDocument("UTF-8", "1.0")
    sw.writeStartElement("a")
    sw.writeStartElement("b")
    sw.writeCharacters("x")
    sw.writeEndDocument()
    sw.flush()
    assertEquals(new String(w.toBytes, "UTF-8"), new String(baos.toByteArray, "UTF-8"))
  }

  test("attribute and xmlns output order is call order") {
    val ops = List(
      Op.Start("a"),
      Op.Attr("xmlns", "urn:d"),
      Op.Attr("xmlns:p", "urn:p"),
      Op.Attr("z", "1"),
      Op.Attr("a", "2"),
      Op.End
    )
    assertEquals(
      new String(driveOurs(ops), "UTF-8").drop(38),
      "<a xmlns=\"urn:d\" xmlns:p=\"urn:p\" z=\"1\" a=\"2\"></a>"
    )
    assertParity(ops)
  }

  test("toBytes does not reset: successive snapshots are prefixes") {
    val w = new XmlTextWriter
    w.writeStartDocument()
    w.writeStartElement("a")
    val snap1 = w.toBytes
    w.writeCharacters("x")
    w.writeEndDocument()
    val snap2 = w.toBytes
    assert(snap2.length > snap1.length)
    assert(java.util.Arrays.equals(snap2.take(snap1.length), snap1))
  }

  test("out-of-protocol calls are no-ops, never throw") {
    val w = new XmlTextWriter
    w.writeAttribute("x", "1") // no open tag
    w.writeEndElement() // empty stack
    assertEquals(w.toBytes.length, 0)
    w.writeStartDocument()
    w.writeStartElement("a")
    w.writeCharacters("t")
    w.writeAttribute("x", "1") // start tag already closed by characters — no-op
    w.writeEndElement()
    assertEquals(new String(w.toBytes, "UTF-8").drop(38), "<a>t</a>")
  }
