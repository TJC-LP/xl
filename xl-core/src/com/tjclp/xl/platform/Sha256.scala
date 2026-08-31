package com.tjclp.xl.platform

import java.security.MessageDigest

/**
 * SHA-256 digest, backed by java.security.MessageDigest on the JVM.
 *
 * This is the ONLY place `java.security` may be imported in xl-core/xl-ooxml (ADR-016): all digest
 * call sites go through this shim so a pure-Scala implementation can replace it on platforms
 * without the JDK security provider (Scala Native / Scala.js, wave B1).
 */
private[xl] object Sha256:

  def digest(bytes: Array[Byte]): Array[Byte] =
    MessageDigest.getInstance("SHA-256").digest(bytes)

  /** Constant-time digest comparison. */
  def isEqual(a: Array[Byte], b: Array[Byte]): Boolean =
    MessageDigest.isEqual(a, b)

  final class Hasher private[Sha256] ():
    private val md = MessageDigest.getInstance("SHA-256")
    def update(buf: Array[Byte], off: Int, len: Int): Unit = md.update(buf, off, len)
    def digest(): Array[Byte] = md.digest()

  def hasher(): Hasher = new Hasher
