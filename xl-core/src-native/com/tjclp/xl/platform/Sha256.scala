package com.tjclp.xl.platform

/**
 * Pure-Scala SHA-256 (FIPS 180-4) for platforms without java.security.MessageDigest (Scala Native's
 * javalib has none). Verified against NIST known-answer vectors.
 */
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
private[xl] object Sha256:

  private val K: Array[Int] = Array(0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b,
    0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74,
    0x80deb1fe, 0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f,
    0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3,
    0xd5a79147, 0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354,
    0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819,
    0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3,
    0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa,
    0xa4506ceb, 0xbef9a3f7, 0xc67178f2)

  def digest(bytes: Array[Byte]): Array[Byte] =
    val h = hasher()
    h.update(bytes, 0, bytes.length)
    h.digest()

  /** Constant-time digest comparison. */
  def isEqual(a: Array[Byte], b: Array[Byte]): Boolean =
    if a.length != b.length then false
    else
      var acc = 0
      var i = 0
      while i < a.length do
        acc |= (a(i) ^ b(i))
        i += 1
      acc == 0

  final class Hasher private[Sha256] ():
    private val h: Array[Int] = Array(0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f,
      0x9b05688c, 0x1f83d9ab, 0x5be0cd19)
    private val block = new Array[Byte](64)
    private val w = new Array[Int](64)
    private var blockLen = 0
    private var totalBytes = 0L

    def update(buf: Array[Byte], off: Int, len: Int): Unit =
      var i = 0
      while i < len do
        block(blockLen) = buf(off + i)
        blockLen += 1
        if blockLen == 64 then
          compress()
          blockLen = 0
        i += 1
      totalBytes += len

    def digest(): Array[Byte] =
      // Padding: 0x80, zeros, 64-bit big-endian bit length
      val bitLen = totalBytes * 8
      block(blockLen) = 0x80.toByte
      blockLen += 1
      if blockLen > 56 then
        while blockLen < 64 do { block(blockLen) = 0; blockLen += 1 }
        compress()
        blockLen = 0
      while blockLen < 56 do { block(blockLen) = 0; blockLen += 1 }
      var i = 0
      while i < 8 do
        block(56 + i) = (bitLen >>> (56 - 8 * i)).toByte
        i += 1
      compress()
      val out = new Array[Byte](32)
      i = 0
      while i < 8 do
        out(4 * i) = (h(i) >>> 24).toByte
        out(4 * i + 1) = (h(i) >>> 16).toByte
        out(4 * i + 2) = (h(i) >>> 8).toByte
        out(4 * i + 3) = h(i).toByte
        i += 1
      out

    private def compress(): Unit =
      var t = 0
      while t < 16 do
        w(t) = ((block(4 * t) & 0xff) << 24) | ((block(4 * t + 1) & 0xff) << 16) |
          ((block(4 * t + 2) & 0xff) << 8) | (block(4 * t + 3) & 0xff)
        t += 1
      while t < 64 do
        val s0 = Integer.rotateRight(w(t - 15), 7) ^ Integer.rotateRight(w(t - 15), 18) ^
          (w(t - 15) >>> 3)
        val s1 = Integer.rotateRight(w(t - 2), 17) ^ Integer.rotateRight(w(t - 2), 19) ^
          (w(t - 2) >>> 10)
        w(t) = w(t - 16) + s0 + w(t - 7) + s1
        t += 1
      var a = h(0); var b = h(1); var c = h(2); var d = h(3)
      var e = h(4); var f = h(5); var g = h(6); var hh = h(7)
      t = 0
      while t < 64 do
        val S1 = Integer.rotateRight(e, 6) ^ Integer.rotateRight(e, 11) ^ Integer.rotateRight(e, 25)
        val ch = (e & f) ^ (~e & g)
        val temp1 = hh + S1 + ch + K(t) + w(t)
        val S0 = Integer.rotateRight(a, 2) ^ Integer.rotateRight(a, 13) ^ Integer.rotateRight(a, 22)
        val maj = (a & b) ^ (a & c) ^ (b & c)
        val temp2 = S0 + maj
        hh = g; g = f; f = e; e = d + temp1
        d = c; c = b; b = a; a = temp1 + temp2
        t += 1
      h(0) += a; h(1) += b; h(2) += c; h(3) += d
      h(4) += e; h(5) += f; h(6) += g; h(7) += hh

  def hasher(): Hasher = new Hasher
