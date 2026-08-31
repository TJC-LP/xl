package com.tjclp.xl.xml.internal

import com.tjclp.xl.xml.ByteInput

private[xml] object CharDecoder:
  /** Sentinel: end of input. */
  val Eof: Int = -1

  /** Sentinel: decode failure — [[CharDecoder.errorMsg]] carries the reason. Sticky. */
  val Err: Int = -2

/**
 * Validating streaming decoder: BOM sniff, UTF-8 / UTF-16(LE|BE), XML 1.0 Char-production
 * validation, line-end normalization (CRLF | lone CR -> LF, XML 1.0 §2.11), 1-based line/col
 * tracking and the maxTotalChars counter. Hand-rolled — no java.nio.charset — so decoding is
 * deterministic across platforms; malformed sequences are positioned errors, never replacement
 * characters.
 *
 * Protocol: after construction, [[cp]] is the current code point (or `Eof`/`Err`); [[advance]]
 * moves to the next. `Eof` and `Err` are sticky. [[line]]/[[col]] are the position of [[cp]].
 */
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[xml] final class CharDecoder(in: ByteInput, maxTotalChars: Long):
  import CharDecoder.{Eof, Err}

  // --- byte buffer -----------------------------------------------------------------------------
  private val buf = new Array[Byte](8192)
  private var bufLen = 0
  private var bufPos = 0
  private var srcEof = false
  private var failMsg = ""

  // encoding: 0 = UTF-8, 1 = UTF-16LE, 2 = UTF-16BE
  private var enc = 0
  private var bom = false

  // current code point and its 1-based position
  private var curCp = 0
  private var curLine = 1
  private var curCol = 1
  private var nextLine = 1
  private var nextCol = 1
  private var skipNextLf = false
  private var produced = 0L

  sniff()
  advance()

  def cp: Int = curCp
  def line: Int = curLine
  def col: Int = curCol
  def errorMsg: String = failMsg

  /** "UTF-8", "UTF-16LE" or "UTF-16BE" as detected from BOM / first bytes. */
  def encodingName: String = enc match
    case 1 => "UTF-16LE"
    case 2 => "UTF-16BE"
    case _ => "UTF-8"

  def hasBom: Boolean = bom

  /** Move to the next code point; sticky on `Eof` and `Err`. */
  def advance(): Unit =
    if curCp != Err then
      curLine = nextLine
      curCol = nextCol
      var c = decodeChecked()
      if skipNextLf then
        skipNextLf = false
        if c == '\n'.toInt then c = decodeChecked()
      if c == '\r'.toInt then
        skipNextLf = true
        c = '\n'.toInt
      curCp = c
      if c >= 0 then
        if c == '\n'.toInt then
          nextLine += 1
          nextCol = 1
        else nextCol += 1

  // --- decoding --------------------------------------------------------------------------------

  private def fail(msg: String): Int =
    if failMsg.isEmpty then failMsg = msg
    Err

  /** Decode one raw code point, validate the XML Char production, count against the limit. */
  private def decodeChecked(): Int =
    val c = decodeRaw()
    if c < 0 then c
    else
      produced += 1
      if produced > maxTotalChars then
        fail(s"maxTotalChars limit exceeded: input longer than $maxTotalChars characters")
      else if isXmlChar(c) then c
      else fail(f"illegal XML character 0x$c%X")

  /** XML 1.0 Char production (pre-normalization: CR is legal). */
  private def isXmlChar(c: Int): Boolean =
    c == 0x9 || c == 0xa || c == 0xd || (c >= 0x20 && c <= 0xd7ff) ||
      (c >= 0xe000 && c <= 0xfffd) || (c >= 0x10000 && c <= 0x10ffff)

  private def decodeRaw(): Int = enc match
    case 0 => decodeUtf8()
    case _ => decodeUtf16()

  private def decodeUtf8(): Int =
    val b0 = nextByte()
    if b0 < 0 then b0
    else if b0 < 0x80 then b0
    else if (b0 & 0xe0) == 0xc0 then
      val b1 = contByte()
      if b1 < 0 then b1
      else
        val v = ((b0 & 0x1f) << 6) | b1
        if v < 0x80 then fail("overlong UTF-8 sequence") else v
    else if (b0 & 0xf0) == 0xe0 then
      val b1 = contByte()
      if b1 < 0 then b1
      else
        val b2 = contByte()
        if b2 < 0 then b2
        else
          val v = ((b0 & 0x0f) << 12) | (b1 << 6) | b2
          if v < 0x800 then fail("overlong UTF-8 sequence")
          else if v >= 0xd800 && v <= 0xdfff then fail("UTF-8 sequence encodes a surrogate")
          else v
    else if (b0 & 0xf8) == 0xf0 then
      val b1 = contByte()
      if b1 < 0 then b1
      else
        val b2 = contByte()
        if b2 < 0 then b2
        else
          val b3 = contByte()
          if b3 < 0 then b3
          else
            val v = ((b0 & 0x07) << 18) | (b1 << 12) | (b2 << 6) | b3
            if v < 0x10000 then fail("overlong UTF-8 sequence")
            else if v > 0x10ffff then fail("UTF-8 sequence beyond U+10FFFF")
            else v
    else fail(f"invalid UTF-8 byte 0x$b0%02X")

  /** A continuation byte's 6 payload bits, or a failure/EOF sentinel. */
  private def contByte(): Int =
    val b = nextByte()
    if b == Eof then fail("truncated UTF-8 sequence")
    else if b == Err then b
    else if (b & 0xc0) != 0x80 then fail("malformed UTF-8 continuation byte")
    else b & 0x3f

  private def decodeUtf16(): Int =
    val u0 = nextUnit16()
    if u0 < 0 then u0
    else if u0 >= 0xd800 && u0 < 0xdc00 then
      val u1 = nextUnit16()
      if u1 == Eof then fail("truncated UTF-16 surrogate pair")
      else if u1 == Err then u1
      else if u1 >= 0xdc00 && u1 < 0xe000 then 0x10000 + ((u0 - 0xd800) << 10) + (u1 - 0xdc00)
      else fail("unpaired UTF-16 high surrogate")
    else if u0 >= 0xdc00 && u0 < 0xe000 then fail("unpaired UTF-16 low surrogate")
    else u0

  private def nextUnit16(): Int =
    val b0 = nextByte()
    if b0 < 0 then b0
    else
      val b1 = nextByte()
      if b1 == Eof then fail("truncated UTF-16 code unit (odd byte count)")
      else if b1 == Err then b1
      else if enc == 1 then b0 | (b1 << 8)
      else (b0 << 8) | b1

  // --- byte supply -----------------------------------------------------------------------------

  /** Next byte as 0..255, or `Eof`/`Err`. */
  private def nextByte(): Int =
    if bufPos < bufLen then
      val b = buf(bufPos) & 0xff
      bufPos += 1
      b
    else if srcEof then Eof
    else
      refill()
      if failMsg.nonEmpty then Err
      else if bufPos < bufLen then
        val b = buf(bufPos) & 0xff
        bufPos += 1
        b
      else Eof

  private def refill(): Unit =
    bufPos = 0
    bufLen = 0
    var zeroReads = 0
    var done = false
    while !done do
      in.read(buf, bufLen, buf.length - bufLen) match
        case Left(msg) =>
          val _ = fail(s"input source failure: $msg")
          done = true
        case Right(-1) =>
          srcEof = true
          done = true
        case Right(0) =>
          // Tolerate a few zero-length reads, but a source that never progresses is a failure —
          // termination is a law (P2), so this cannot spin forever.
          zeroReads += 1
          if zeroReads > 64 then
            val _ = fail("input source failure: read returned 0 bytes repeatedly")
            done = true
        case Right(n) =>
          bufLen += n
          done = true

  // --- BOM / encoding sniff --------------------------------------------------------------------

  private def sniff(): Unit =
    // Pull the first bytes without consuming past the BOM.
    refill()
    var zeroReads = 0
    while !srcEof && bufLen < 3 && failMsg.isEmpty && zeroReads <= 64 do
      in.read(buf, bufLen, buf.length - bufLen) match
        case Left(msg) => val _ = fail(s"input source failure: $msg")
        case Right(-1) => srcEof = true
        case Right(0) => zeroReads += 1
        case Right(n) => bufLen += n
    val b0 = if bufLen > 0 then buf(0) & 0xff else -1
    val b1 = if bufLen > 1 then buf(1) & 0xff else -1
    val b2 = if bufLen > 2 then buf(2) & 0xff else -1
    if b0 == 0xef && b1 == 0xbb && b2 == 0xbf then
      enc = 0
      bom = true
      bufPos = 3
    else if b0 == 0xff && b1 == 0xfe then
      enc = 1
      bom = true
      bufPos = 2
    else if b0 == 0xfe && b1 == 0xff then
      enc = 2
      bom = true
      bufPos = 2
    else if b0 == 0x3c && b1 == 0x00 then enc = 1
    else if b0 == 0x00 && b1 == 0x3c then enc = 2
    else enc = 0
