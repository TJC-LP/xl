package com.tjclp.xl.xml

/**
 * Owned (non-aliased) token snapshot for law-based tests — TEST-ONLY surface; the shipped API stays
 * the zero-allocation cursor.
 */
enum OwnedToken derives CanEqual:
  case Start(
    qName: String,
    prefix: String,
    local: String,
    attrs: Vector[(String, String)],
    depth: Int
  )
  case End(qName: String)
  case Text(value: String)
  case CData(value: String)

object TestTokens:
  /** Materialize the complete token trace, or the first error. Never throws. */
  def tokens(p: XmlPullParser): Either[XmlParseError, Vector[OwnedToken]] =
    val b = Vector.newBuilder[OwnedToken]
    @annotation.tailrec
    def loop(): Either[XmlParseError, Vector[OwnedToken]] =
      p.next() match
        case Left(e) => Left(e)
        case Right(XmlToken.EndDocument) => Right(b.result())
        case Right(XmlToken.StartElement) =>
          b += OwnedToken.Start(p.name, p.prefix, p.localName, p.attrs.toVector, p.depth)
          loop()
        case Right(XmlToken.EndElement) =>
          b += OwnedToken.End(p.name)
          loop()
        case Right(XmlToken.Text) =>
          b += OwnedToken.Text(p.text)
          loop()
        case Right(XmlToken.CData) =>
          b += OwnedToken.CData(p.text)
          loop()
    loop()

  def parse(
    xml: String,
    limits: XmlLimits = XmlLimits.default
  ): Either[XmlParseError, Vector[OwnedToken]] =
    tokens(XmlPullParser.fromString(xml, limits))

  def parseBytes(
    bytes: Array[Byte],
    limits: XmlLimits = XmlLimits.default
  ): Either[XmlParseError, Vector[OwnedToken]] =
    tokens(XmlPullParser.fromArray(bytes, limits))
