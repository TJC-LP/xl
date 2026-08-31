package com.tjclp.xl.xml

/**
 * Portable byte source abstraction — no java.io in the core API so the module cross-builds to
 * Native/JS unchanged (ADR-016). An InputStream adapter lives outside xl-xml, next to its JVM
 * consumers (wave A3).
 */
trait ByteInput:
  /**
   * Fill `dst[off, off + len)`. `Right(n > 0)` = bytes read, `Right(-1)` = EOF. `Left` = source
   * failure, surfaced by the parser as an [[XmlParseError]] at the current parse position.
   * Implementations MUST NOT throw.
   */
  def read(dst: Array[Byte], off: Int, len: Int): Either[String, Int]

object ByteInput:
  /** View over `bytes` (not copied — callers must not mutate the array while parsing). */
  def fromArray(bytes: Array[Byte]): ByteInput = new ByteInput:
    @SuppressWarnings(Array("org.wartremover.warts.Var"))
    private var pos = 0
    def read(dst: Array[Byte], off: Int, len: Int): Either[String, Int] =
      if pos >= bytes.length then Right(-1)
      else
        val n = math.min(len, bytes.length - pos)
        System.arraycopy(bytes, pos, dst, off, n)
        pos += n
        Right(n)

  /**
   * UTF-8-encodes `xml` once (test/DSL convenience). Lone surrogates — impossible in well-formed
   * Scala string literals but representable — encode as U+FFFD.
   */
  def fromString(xml: String): ByteInput = fromArray(internal.Utf8.encode(xml))
