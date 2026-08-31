package com.tjclp.xl.xml

import com.tjclp.xl.xml.internal.Utf8

/**
 * Low-level streaming XML byte emitter, byte-identical to the JDK StAX writer
 * (`XMLStreamWriterImpl` built with `IS_REPAIRING_NAMESPACES=false`, "UTF-8") for the event
 * protocols xl-ooxml's emitters drive. StAX-shaped method names so the PortableSaxWriter adapter
 * diff is line-comparable to StaxSaxWriter's `underlying.*` calls.
 *
 * Probe-verified StAX byte behaviors replicated exactly:
 *   - Declaration: `<?xml version="1.0" encoding="UTF-8"?>` — no standalone, no trailing newline.
 *   - Text escapes ONLY `& < >`; `"` `'` and raw tab/LF/CR pass through (CR stays byte 0x0D).
 *   - Attribute values escape `& < > "`; `'` and tab/LF/CR pass through raw.
 *   - A start/end pair with no children emits `<a></a>`, NEVER minimized; only
 *     [[writeEmptyElement]] produces `/>` (closed lazily by the next event — GH-430).
 *   - [[writeEndDocument]] auto-closes all open elements. Attribute/xmlns output order = call
 *     order. Non-ASCII is raw UTF-8, including astral planes.
 *
 * This writer is deliberately dumb: no namespace-scope tracking (that lives in the adapter), no
 * OOXML sanitize policy (GH-237 stays in the adapter). Out-of-protocol calls (writeAttribute with
 * no open tag, writeEndElement on an empty stack) are documented no-ops — total, never throws. Only
 * our own emitters drive it. Lone surrogates in input (out of contract) encode as U+FFFD.
 *
 * Not thread-safe. Create one instance per document.
 */
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
final class XmlTextWriter:
  private var buf = new Array[Byte](1024)
  private var len = 0

  // pending open tag: 0 = none, 1 = start tag (close with '>'), 2 = empty element (close with "/>")
  private var pending = 0

  // open-element qName stack
  private var stack = new Array[String](16)
  private var stackSize = 0

  /** Bytes emitted so far (a copy; does not reset the writer). */
  def toBytes: Array[Byte] =
    val out = new Array[Byte](len)
    System.arraycopy(buf, 0, out, 0, len)
    out

  def writeStartDocument(): Unit =
    writeAscii("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")

  def writeStartElement(qName: String): Unit =
    closePending()
    writeByte('<')
    writeRaw(qName)
    push(qName)
    pending = 1

  def writeEmptyElement(qName: String): Unit =
    closePending()
    writeByte('<')
    writeRaw(qName)
    pending = 2

  /** No-op when no start/empty tag is open (out of protocol). */
  def writeAttribute(qName: String, value: String): Unit =
    if pending != 0 then
      writeByte(' ')
      writeRaw(qName)
      writeByte('=')
      writeByte('"')
      writeEscaped(value, attr = true)
      writeByte('"')

  def writeCharacters(text: String): Unit =
    closePending()
    writeEscaped(text, attr = false)

  /** Closes the most recently opened (non-empty) element; no-op on an empty stack. */
  def writeEndElement(): Unit =
    closePending()
    if stackSize > 0 then
      stackSize -= 1
      writeByte('<')
      writeByte('/')
      writeRaw(stack(stackSize))
      writeByte('>')

  /** Auto-closes every open element (StAX behavior). */
  def writeEndDocument(): Unit =
    closePending()
    while stackSize > 0 do writeEndElement()

  // --- internals ---------------------------------------------------------------------------------

  private def closePending(): Unit =
    if pending == 1 then
      writeByte('>')
      pending = 0
    else if pending == 2 then
      writeByte('/')
      writeByte('>')
      pending = 0

  private def push(qName: String): Unit =
    if stackSize == stack.length then
      val grown = new Array[String](stack.length * 2)
      System.arraycopy(stack, 0, grown, 0, stackSize)
      stack = grown
    stack(stackSize) = qName
    stackSize += 1

  private def ensure(n: Int): Unit =
    if len + n > buf.length then
      val grown = new Array[Byte](math.max(buf.length * 2, len + n))
      System.arraycopy(buf, 0, grown, 0, len)
      buf = grown

  private def writeByte(c: Char): Unit =
    ensure(1)
    buf(len) = c.toByte
    len += 1

  private def writeAscii(s: String): Unit =
    ensure(s.length)
    var i = 0
    while i < s.length do
      buf(len) = s.charAt(i).toByte
      len += 1
      i += 1

  /** Raw UTF-8, no escaping (names). */
  private def writeRaw(s: String): Unit =
    var i = 0
    while i < s.length do
      val c = s.charAt(i)
      if c < 0x80 then
        ensure(1)
        buf(len) = c.toByte
        len += 1
      else i = writeNonAscii(s, i)
      i += 1

  /** Text/attribute content with the StAX escape set. */
  private def writeEscaped(s: String, attr: Boolean): Unit =
    var i = 0
    while i < s.length do
      val c = s.charAt(i)
      if c == '&' then writeAscii("&amp;")
      else if c == '<' then writeAscii("&lt;")
      else if c == '>' then writeAscii("&gt;")
      else if attr && c == '"' then writeAscii("&quot;")
      else if c < 0x80 then
        ensure(1)
        buf(len) = c.toByte
        len += 1
      else i = writeNonAscii(s, i)
      i += 1

  /** Encode the (possibly supplementary) code point at `i`; returns the index of its last char. */
  private def writeNonAscii(s: String, i: Int): Int =
    val c = s.charAt(i)
    ensure(4)
    if Character.isHighSurrogate(c) && i + 1 < s.length && Character.isLowSurrogate(s.charAt(i + 1))
    then
      len = Utf8.appendCp(buf, len, Character.toCodePoint(c, s.charAt(i + 1)))
      i + 1
    else if Character.isSurrogate(c) then
      len = Utf8.appendCp(buf, len, 0xfffd)
      i
    else
      len = Utf8.appendCp(buf, len, c.toInt)
      i
