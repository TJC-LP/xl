package com.tjclp.xl.xml

import org.scalacheck.Gen
import org.scalacheck.Prop.forAll

/**
 * R1: parse(XmlTextWriter render of a generated tree) reproduces the tree's canonical trace. P2:
 * progress/termination — total next() calls are linearly bounded by input size.
 */
class RoundTripPropertySpec extends munit.ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(200)

  // --- generated XML trees ------------------------------------------------------------------------

  private final case class GenElem(
    name: String,
    attrs: Vector[(String, String)],
    children: Vector[Either[String, GenElem]] // Left = text run
  )

  private val genNcName: Gen[String] =
    for
      h <- Gen.alphaChar
      t <- Gen.listOfN(3, Gen.oneOf(Gen.alphaNumChar, Gen.const('-'), Gen.const('.')))
    yield (h :: t).mkString

  // Text alphabet exercises every escape and raw pass-through EXCEPT raw CR (parse-time line-end
  // normalization folds it to LF by design — the ST_Xstring premise) and, for attributes, raw
  // tab/LF (attribute-value normalization folds them to spaces by design).
  private val genTextChar: Gen[Char] =
    Gen.frequency(
      8 -> Gen.alphaNumChar,
      3 -> Gen.oneOf('&', '<', '>', '"', '\'', ' ', '\t', '\n'),
      2 -> Gen.oneOf('é', '日', '€'),
      1 -> Gen.const('\ud83d') // paired below
    )

  private val genText: Gen[String] =
    Gen.nonEmptyListOf(genTextChar).map { cs =>
      // pair any lone high surrogate into a valid emoji
      cs.map(c => if c == '\ud83d' then "😀" else c.toString).mkString
    }

  private val genAttrValue: Gen[String] =
    Gen
      .listOf(
        Gen.frequency(
          8 -> Gen.alphaNumChar.map(_.toString),
          3 -> Gen.oneOf("&", "<", ">", "\"", "'", " "),
          1 -> Gen.const("é")
        )
      )
      .map(_.mkString)

  private def genElem(depth: Int): Gen[GenElem] =
    for
      name <- genNcName
      nAttrs <- Gen.choose(0, 3)
      attrNames <- Gen.listOfN(nAttrs, genNcName).map(_.distinct)
      attrs <- Gen.sequence[Vector[(String, String)], (String, String)](
        attrNames.toVector.map(n => genAttrValue.map(v => n -> v))
      )
      nKids <- if depth <= 0 then Gen.const(0) else Gen.choose(0, 3)
      kids <- Gen.listOfN(
        nKids,
        Gen.frequency(
          3 -> genText.map(Left(_)),
          2 -> Gen.lzy(genElem(depth - 1)).map(Right(_))
        )
      )
    yield GenElem(name, attrs, mergeAdjacentText(kids.toVector))

  private def mergeAdjacentText(
    kids: Vector[Either[String, GenElem]]
  ): Vector[Either[String, GenElem]] =
    kids.foldLeft(Vector.empty[Either[String, GenElem]]) {
      case (acc :+ Left(a), Left(b)) => acc :+ Left(a + b)
      case (acc, k) => acc :+ k
    }

  // --- render and canonical traces -----------------------------------------------------------------

  private def render(root: GenElem): Array[Byte] =
    val w = new XmlTextWriter
    w.writeStartDocument()
    def emit(e: GenElem): Unit =
      if e.children.isEmpty && e.attrs.isEmpty then
        w.writeStartElement(e.name)
        w.writeEndElement()
      else
        w.writeStartElement(e.name)
        e.attrs.foreach { case (n, v) => w.writeAttribute(n, v) }
        e.children.foreach {
          case Left(text) => w.writeCharacters(text)
          case Right(kid) => emit(kid)
        }
        w.writeEndElement()
    emit(root)
    w.writeEndDocument()
    w.toBytes

  /** What the parser must see: attr values normalized per §3.3.3 (writer emits raw tab/LF). */
  private def normAttr(v: String): String = v.map(c => if c == '\t' || c == '\n' then ' ' else c)

  private def expectedTrace(root: GenElem): Vector[OwnedToken] =
    val b = Vector.newBuilder[OwnedToken]
    def walk(e: GenElem, depth: Int): Unit =
      b += OwnedToken.Start(
        e.name,
        "",
        e.name,
        e.attrs.map { case (n, v) => n -> normAttr(v) },
        depth
      )
      e.children.foreach {
        case Left(text) => b += OwnedToken.Text(text)
        case Right(kid) => walk(kid, depth + 1)
      }
      b += OwnedToken.End(e.name)
    walk(root, 1)
    b.result()

  // --- laws ------------------------------------------------------------------------------------------

  property("R1: writer -> parser round-trip preserves structure") {
    forAll(genElem(3)) { root =>
      val bytes = render(root)
      TestTokens.parseBytes(bytes) match
        case Right(ts) =>
          assertEquals(ts, expectedTrace(root), new String(bytes, "UTF-8"))
          true
        case Left(e) => fail(s"round-trip parse failed: ${e.render}\n${new String(bytes, "UTF-8")}")
    }
  }

  property("P2: progress — total next() calls bounded by 2 * chars + O(1)") {
    forAll(genElem(3)) { root =>
      val bytes = render(root)
      val p = XmlPullParser.fromArray(bytes)
      var calls = 0
      var r = p.next()
      calls += 1
      while r.isRight && r != Right(XmlToken.EndDocument) && calls <= 2 * bytes.length + 16 do
        r = p.next()
        calls += 1
      calls <= 2 * bytes.length + 16
    }
  }

  test("R1 fixed case: every construct in one document") {
    val xml = new XmlTextWriter
    xml.writeStartDocument()
    xml.writeStartElement("root")
    xml.writeAttribute("a", "1 & 2 <ok> \"q\"")
    xml.writeCharacters("head & tail")
    xml.writeEmptyElement("leaf")
    xml.writeAttribute("x", "y")
    xml.writeStartElement("mid")
    xml.writeCharacters("日本語 😀")
    xml.writeEndElement()
    xml.writeEndDocument()
    val ts = TestTokens.parseBytes(xml.toBytes) match
      case Right(v) => v
      case Left(e) => fail(e.render)
    assertEquals(
      ts,
      Vector(
        OwnedToken.Start("root", "", "root", Vector("a" -> "1 & 2 <ok> \"q\""), 1),
        OwnedToken.Text("head & tail"),
        OwnedToken.Start("leaf", "", "leaf", Vector("x" -> "y"), 2),
        OwnedToken.End("leaf"),
        OwnedToken.Start("mid", "", "mid", Vector.empty, 2),
        OwnedToken.Text("日本語 😀"),
        OwnedToken.End("mid"),
        OwnedToken.End("root")
      )
    )
  }
