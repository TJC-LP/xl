package com.tjclp.xl.context

import java.nio.file.{Files, Path}
import java.security.MessageDigest
import scala.collection.immutable.ArraySeq

object SourceFingerprint:
  /** Compute fingerprint for a file path by streaming through its bytes. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def fromPath(path: Path): SourceFingerprint =
    val digest = MessageDigest.getInstance("SHA-256")
    val buffer = new Array[Byte](8192)
    val in = Files.newInputStream(path)
    try
      var bytesRead = 0L
      var read = in.read(buffer)
      while read != -1 do
        digest.update(buffer, 0, read)
        bytesRead += read
        read = in.read(buffer)
      SourceFingerprint(bytesRead, ArraySeq.unsafeWrapArray(digest.digest()))
    finally in.close()

final case class SourceFingerprint(size: Long, sha256: ArraySeq[Byte]) derives CanEqual:

  /** Verify that the provided digest and size matches the recorded fingerprint. */
  def matches(bytesRead: Long, digest: Array[Byte]): Boolean =
    bytesRead == size && MessageDigest.isEqual(sha256.toArray, digest)
