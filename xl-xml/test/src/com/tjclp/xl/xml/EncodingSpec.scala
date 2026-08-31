package com.tjclp.xl.xml

import munit.FunSuite
import OwnedToken.*

class EncodingSpec extends FunSuite:

  private def utf16(s: String, littleEndian: Boolean, bom: Boolean): Array[Byte] =
    val b = Array.newBuilder[Byte]
    if bom then
      if littleEndian then
        b += 0xff.toByte
        b += 0xfe.toByte
      else
        b += 0xfe.toByte
        b += 0xff.toByte
    s.foreach { c =>
      val hi = ((c >> 8) & 0xff).toByte
      val lo = (c & 0xff).toByte
      if littleEndian then
        b += lo
        b += hi
      else
        b += hi
        b += lo
    }
    b.result()

  private def okBytes(bytes: Array[Byte]): Vector[OwnedToken] =
    TestTokens.parseBytes(bytes) match
      case Right(ts) => ts
      case Left(e) => fail(s"expected success, got: ${e.render}")

  private def badBytes(bytes: Array[Byte]): XmlParseError =
    TestTokens.parseBytes(bytes) match
      case Left(e) => e
      case Right(ts) => fail(s"expected failure, got tokens: $ts")

  private val expected =
    Vector(Start("a", "", "a", Vector.empty, 1), Text("héllo"), End("a"))

  test("UTF-8 without BOM") {
    assertEquals(okBytes(internal.Utf8.encode("<a>héllo</a>")), expected)
  }

  test("UTF-8 with BOM") {
    val bom = Array(0xef.toByte, 0xbb.toByte, 0xbf.toByte)
    assertEquals(okBytes(bom ++ internal.Utf8.encode("<a>héllo</a>")), expected)
  }

  test("UTF-16LE with BOM") {
    assertEquals(okBytes(utf16("<a>héllo</a>", littleEndian = true, bom = true)), expected)
  }

  test("UTF-16BE with BOM") {
    assertEquals(okBytes(utf16("<a>héllo</a>", littleEndian = false, bom = true)), expected)
  }

  test("BOM-less UTF-16LE is sniffed from '<' 0x00") {
    assertEquals(okBytes(utf16("<a>héllo</a>", littleEndian = true, bom = false)), expected)
  }

  test("BOM-less UTF-16BE is sniffed from 0x00 '<'") {
    assertEquals(okBytes(utf16("<a>héllo</a>", littleEndian = false, bom = false)), expected)
  }

  test("astral characters decode in UTF-8 and UTF-16") {
    val doc = "<a>😀</a>" // U+1F600
    val exp = Vector(Start("a", "", "a", Vector.empty, 1), Text("😀"), End("a"))
    assertEquals(okBytes(internal.Utf8.encode(doc)), exp)
    assertEquals(okBytes(utf16(doc, littleEndian = true, bom = true)), exp)
    assertEquals(okBytes(utf16(doc, littleEndian = false, bom = true)), exp)
  }

  // --- encoding declaration consistency ------------------------------------------------------------

  test("encoding declaration utf-8 spellings accepted for UTF-8 input") {
    for enc <- List("UTF-8", "utf-8", "Utf-8", "US-ASCII", "us-ascii") do
      val doc = s"""<?xml version="1.0" encoding="$enc"?><a/>"""
      assert(TestTokens.parse(doc).isRight, s"encoding $enc should parse")
  }

  test("encoding declaration utf-16 spellings accepted for UTF-16 input") {
    for
      (le, spellings) <- List(
        true -> List("UTF-16", "utf-16le"),
        false -> List("UTF-16", "UTF-16BE")
      )
      enc <- spellings
    do
      val doc = s"""<?xml version="1.0" encoding="$enc"?><a/>"""
      val r = TestTokens.parseBytes(utf16(doc, littleEndian = le, bom = true))
      assert(r.isRight, s"encoding $enc (le=$le) should parse: $r")
  }

  test("BOM wins: UTF-8 input with encoding='UTF-16' is rejected as a conflict") {
    val e = badBytes(internal.Utf8.encode("""<?xml version="1.0" encoding="UTF-16"?><a/>"""))
    assert(e.message.contains("conflicts"), e.message)
  }

  test("UTF-16 input with encoding='UTF-8' is rejected as a conflict") {
    val e = badBytes(utf16("""<?xml version="1.0" encoding="UTF-8"?><a/>""", true, bom = true))
    assert(e.message.contains("conflicts"), e.message)
  }

  test("wrong UTF-16 endianness spelling is rejected") {
    val e = badBytes(utf16("""<?xml version="1.0" encoding="UTF-16BE"?><a/>""", true, bom = true))
    assert(e.message.contains("conflicts"), e.message)
  }

  test("unsupported encodings fail closed with a named error") {
    for enc <- List("ISO-8859-1", "latin1", "Shift_JIS", "EBCDIC-US") do
      val e = badBytes(internal.Utf8.encode(s"""<?xml version="1.0" encoding="$enc"?><a/>"""))
      assert(e.message.contains("unsupported encoding"), s"$enc: ${e.message}")
      assert(e.message.contains(enc), s"error should name the encoding: ${e.message}")
  }

  // --- malformed byte sequences ----------------------------------------------------------------------

  test("truncated multibyte UTF-8 at EOF is a positioned error") {
    val bytes = internal.Utf8.encode("<a>x") ++ Array(0xc3.toByte)
    val e = badBytes(bytes)
    assert(e.message.contains("truncated UTF-8"), e.message)
    assert(e.line >= 1 && e.col >= 1)
  }

  test("truncated multibyte UTF-8 mid-document is rejected") {
    val bytes = internal.Utf8.encode("<a>") ++ Array(0xe2.toByte, 0x82.toByte) ++
      internal.Utf8.encode("</a>")
    val e = badBytes(bytes)
    assert(e.message.contains("UTF-8"), e.message)
  }

  test("overlong UTF-8 encoding is rejected, never replaced") {
    // 0xC0 0xAF is an overlong encoding of '/'
    val bytes = internal.Utf8.encode("<a>") ++ Array(0xc0.toByte, 0xaf.toByte) ++
      internal.Utf8.encode("</a>")
    val e = badBytes(bytes)
    assert(e.message.contains("overlong"), e.message)
  }

  test("UTF-8-encoded surrogate is rejected") {
    // 0xED 0xA0 0x80 encodes U+D800
    val bytes = internal.Utf8.encode("<a>") ++ Array(0xed.toByte, 0xa0.toByte, 0x80.toByte) ++
      internal.Utf8.encode("</a>")
    val e = badBytes(bytes)
    assert(e.message.contains("surrogate"), e.message)
  }

  test("invalid UTF-8 lead byte is rejected") {
    val bytes = internal.Utf8.encode("<a>") ++ Array(0xff.toByte) ++ internal.Utf8.encode("</a>")
    val e = badBytes(bytes)
    assert(e.message.contains("invalid UTF-8"), e.message)
  }

  test("bad UTF-8 continuation byte is rejected") {
    val bytes = internal.Utf8.encode("<a>") ++ Array(0xc3.toByte, 0x28.toByte) ++
      internal.Utf8.encode("</a>")
    val e = badBytes(bytes)
    assert(e.message.contains("continuation"), e.message)
  }

  test("UTF-8 sequence beyond U+10FFFF is rejected") {
    val bytes = internal.Utf8.encode("<a>") ++
      Array(0xf4.toByte, 0x90.toByte, 0x80.toByte, 0x80.toByte) ++ internal.Utf8.encode("</a>")
    val e = badBytes(bytes)
    assert(e.message.contains("U+10FFFF"), e.message)
  }

  test("odd trailing byte in UTF-16 is rejected") {
    val bytes = utf16("<a></a>", littleEndian = true, bom = true) ++ Array(0x20.toByte)
    // trailing garbage after root: either odd-byte error or content-after-root — must be Left
    assert(TestTokens.parseBytes(bytes).isLeft)
  }

  test("unpaired UTF-16 high surrogate is rejected") {
    val bytes = utf16("<a>", littleEndian = true, bom = true) ++
      Array(0x00.toByte, 0xd8.toByte) ++ utf16("</a>", littleEndian = true, bom = false).drop(0)
    val e = badBytes(bytes)
    assert(e.message.contains("surrogate"), e.message)
  }

  test("unpaired UTF-16 low surrogate is rejected") {
    val bytes = utf16("<a>", littleEndian = false, bom = true) ++
      Array(0xdc.toByte, 0x00.toByte) ++ utf16("</a>", littleEndian = false, bom = false)
    val e = badBytes(bytes)
    assert(e.message.contains("surrogate"), e.message)
  }

  // --- illegal XML characters -------------------------------------------------------------------------

  test("NUL byte is an illegal XML character") {
    val bytes = internal.Utf8.encode("<a>") ++ Array(0x00.toByte) ++ internal.Utf8.encode("</a>")
    val e = badBytes(bytes)
    assert(e.message.contains("illegal XML character"), e.message)
  }

  test("0x0B vertical tab is an illegal XML character") {
    val bytes = internal.Utf8.encode("<a>") ++ Array(0x0b.toByte) ++ internal.Utf8.encode("</a>")
    val e = badBytes(bytes)
    assert(e.message.contains("illegal XML character"), e.message)
  }

  test("U+FFFE is an illegal XML character") {
    val bytes = internal.Utf8.encode("<a>") ++ internal.Utf8.encode("￾") ++
      internal.Utf8.encode("</a>")
    val e = badBytes(bytes)
    assert(e.message.contains("illegal XML character"), e.message)
  }

  test("empty byte input is a clean error") {
    val e = badBytes(Array.empty[Byte])
    assert(e.message.contains("no root element"), e.message)
  }

  test("input of only a BOM is a clean error") {
    val e = badBytes(Array(0xef.toByte, 0xbb.toByte, 0xbf.toByte))
    assert(e.message.contains("no root element"), e.message)
  }

  test("ByteInput source failure surfaces as a positioned parse error") {
    val failing = new ByteInput:
      def read(dst: Array[Byte], off: Int, len: Int): Either[String, Int] = Left("disk on fire")
    TestTokens.tokens(XmlPullParser(failing)) match
      case Left(e) => assert(e.message.contains("disk on fire"), e.message)
      case Right(ts) => fail(s"expected failure, got $ts")
  }
