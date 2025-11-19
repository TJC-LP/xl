package com.tjclp.xl

import java.nio.file.{Files, Path}
import java.security.MessageDigest

import scala.collection.immutable.ArraySeq

import com.tjclp.xl.ooxml.PartManifest

final case class SourceFingerprint(size: Long, sha256: ArraySeq[Byte]) derives CanEqual:

  /** Verify that the provided digest and size matches the recorded fingerprint. */
  def matches(bytesRead: Long, digest: Array[Byte]): Boolean =
    bytesRead == size && MessageDigest.isEqual(sha256.toArray, digest)

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

/**
 * Captures metadata about the physical XLSX that produced a [[Workbook]]. The context enables
 * surgical write operations by preserving the manifest of ZIP entries. The PreservedPartStore can
 * be reconstructed from the sourcePath when needed for IO operations.
 */
final case class SourceContext(
  sourcePath: Path,
  partManifest: PartManifest,
  modificationTracker: ModificationTracker,
  fingerprint: SourceFingerprint
) derives CanEqual:

  /** True when no workbooks modifications have been recorded. */
  def isClean: Boolean = modificationTracker.isClean

  /** Mark a sheets as modified. */
  def markSheetModified(sheetIndex: Int): SourceContext =
    copy(modificationTracker = modificationTracker.markSheet(sheetIndex))

  /** Mark a sheets as deleted. */
  def markSheetDeleted(sheetIndex: Int): SourceContext =
    copy(modificationTracker = modificationTracker.delete(sheetIndex))

  /** Mark sheets order as changed. */
  def markReordered: SourceContext =
    copy(modificationTracker = modificationTracker.markReordered)

  /** Mark workbooks-level metadata as changed. */
  def markMetadataModified: SourceContext =
    copy(modificationTracker = modificationTracker.markMetadata)

object SourceContext:
  /** Construct a context for a workbooks that originated from a file. */
  def fromFile(path: Path, manifest: PartManifest, fingerprint: SourceFingerprint): SourceContext =
    SourceContext(path, manifest, ModificationTracker.clean, fingerprint)
