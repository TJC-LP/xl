package com.tjclp.xl.context

import com.tjclp.xl.ooxml.PartManifest
import com.tjclp.xl.workbooks.Workbook

import java.nio.file.{Files, Path}
import java.security.MessageDigest
import scala.collection.immutable.ArraySeq

/**
 * Captures metadata about the physical XLSX that produced a [[Workbook]]. The context enables
 * surgical write operations by preserving the manifest of ZIP entries. The PreservedPartStore can
 * be reconstructed from the sourcePath when needed for IO operations.
 *
 * @param commentPathMapping
 *   Mapping from 0-based sheet index to comment file path (e.g., "xl/comments1.xml"). Excel numbers
 *   comment files sequentially across only sheets that have comments, NOT by sheet index. This
 *   mapping preserves the original paths to enable correct surgical writes.
 */
final case class SourceContext(
  sourcePath: Path,
  partManifest: PartManifest,
  modificationTracker: ModificationTracker,
  fingerprint: SourceFingerprint,
  commentPathMapping: Map[Int, String] = Map.empty
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
  /**
   * Construct a context for a workbook that originated from a file.
   *
   * @param commentPathMapping
   *   Mapping from 0-based sheet index to comment file path. Excel numbers comment files
   *   sequentially (comments1.xml, comments2.xml...) across only sheets that have comments.
   */
  def fromFile(
    path: Path,
    manifest: PartManifest,
    fingerprint: SourceFingerprint,
    commentPathMapping: Map[Int, String] = Map.empty
  ): SourceContext =
    SourceContext(path, manifest, ModificationTracker.clean, fingerprint, commentPathMapping)
