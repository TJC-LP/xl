package com.tjclp.xl.xml.internal

/**
 * Hand-rolled UTF-8 encoder — java.nio.charset is deliberately avoided everywhere in xl-xml so
 * behavior is bit-identical across JVM/Native/JS (ADR-016).
 */
private[xml] object Utf8:

  /** Encode `s` as UTF-8. Lone surrogates encode as U+FFFD. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def encode(s: String): Array[Byte] =
    var out = new Array[Byte](math.max(16, s.length + (s.length >> 1)))
    var len = 0
    def ensure(n: Int): Unit =
      if len + n > out.length then
        val grown = new Array[Byte](math.max(out.length * 2, len + n))
        System.arraycopy(out, 0, grown, 0, len)
        out = grown
    var i = 0
    while i < s.length do
      val c = s.charAt(i)
      val cp =
        if Character.isHighSurrogate(c) && i + 1 < s.length &&
          Character.isLowSurrogate(s.charAt(i + 1))
        then
          i += 1
          Character.toCodePoint(c, s.charAt(i))
        else if Character.isSurrogate(c) then 0xfffd
        else c.toInt
      ensure(4)
      len = appendCp(out, len, cp)
      i += 1
    val result = new Array[Byte](len)
    System.arraycopy(out, 0, result, 0, len)
    result

  /**
   * Append code point `cp` (a valid scalar value) to `buf` at `pos`; returns the new position. The
   * caller must have reserved at least 4 bytes.
   */
  def appendCp(buf: Array[Byte], pos: Int, cp: Int): Int =
    if cp < 0x80 then
      buf(pos) = cp.toByte
      pos + 1
    else if cp < 0x800 then
      buf(pos) = (0xc0 | (cp >> 6)).toByte
      buf(pos + 1) = (0x80 | (cp & 0x3f)).toByte
      pos + 2
    else if cp < 0x10000 then
      buf(pos) = (0xe0 | (cp >> 12)).toByte
      buf(pos + 1) = (0x80 | ((cp >> 6) & 0x3f)).toByte
      buf(pos + 2) = (0x80 | (cp & 0x3f)).toByte
      pos + 3
    else
      buf(pos) = (0xf0 | (cp >> 18)).toByte
      buf(pos + 1) = (0x80 | ((cp >> 12) & 0x3f)).toByte
      buf(pos + 2) = (0x80 | ((cp >> 6) & 0x3f)).toByte
      buf(pos + 3) = (0x80 | (cp & 0x3f)).toByte
      pos + 4
