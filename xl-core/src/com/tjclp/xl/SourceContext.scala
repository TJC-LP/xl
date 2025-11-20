package com.tjclp.xl

import java.nio.file.{Files, Path}
import java.security.MessageDigest

import scala.collection.immutable.ArraySeq

import com.tjclp.xl.ooxml.PartManifest

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

  /** True when no workbook modifications have been recorded. */
  def isClean: Boolean = modificationTracker.isClean

  /** Mark a sheet as modified. */
  def markSheetModified(sheetIndex: Int): SourceContext =
    copy(modificationTracker = modificationTracker.markSheet(sheetIndex))

  /** Mark a sheet as deleted. */
  def markSheetDeleted(sheetIndex: Int): SourceContext =
    copy(modificationTracker = modificationTracker.delete(sheetIndex))

  /** Mark sheet order as changed. */
  def markReordered: SourceContext =
    copy(modificationTracker = modificationTracker.markReordered)

  /** Mark workbook-level metadata as changed. */
  def markMetadataModified: SourceContext =
    copy(modificationTracker = modificationTracker.markMetadata)

object SourceContext:
  /** Construct a context for a workbook that originated from a file. */
  def fromFile(path: Path, manifest: PartManifest, fingerprint: SourceFingerprint): SourceContext =
    SourceContext(path, manifest, ModificationTracker.clean, fingerprint)
