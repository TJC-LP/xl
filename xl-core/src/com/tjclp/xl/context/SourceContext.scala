package com.tjclp.xl.context

import com.tjclp.xl.charts.Chart
import com.tjclp.xl.drawings.Drawing
import com.tjclp.xl.ooxml.PartManifest
import com.tjclp.xl.workbooks.Workbook

import java.nio.file.{Files, Path}
import java.security.MessageDigest
import scala.collection.immutable.ArraySeq

/**
 * Provenance of one typed chart parsed at read time (GH-222): which anchor hosted it, the `r:id`
 * that referenced it from the drawing part, the chart part's zip path, and the as-parsed typed
 * value. The writer's chart planner equality-matches edited drawings against these snapshots to
 * reuse parts/rel-ids instead of churning fresh ones.
 */
final case class ChartSnapshot(
  anchorIdx: Int,
  relId: String,
  partPath: String,
  chart: Chart
) derives CanEqual

/**
 * Captures metadata about the physical XLSX that produced a [[Workbook]]. The context enables
 * surgical write operations by preserving the manifest of ZIP entries. The PreservedPartStore can
 * be reconstructed from the sourcePath when needed for IO operations.
 *
 * @param commentPathMapping
 *   Mapping from 0-based sheet index to comment file path (e.g., "xl/comments1.xml"). Excel numbers
 *   comment files sequentially across only sheets that have comments, NOT by sheet index. This
 *   mapping preserves the original paths to enable correct surgical writes.
 * @param drawingPathMapping
 *   Mapping from 0-based sheet index to drawing part path (e.g., "xl/drawings/drawing1.xml") for
 *   sheets whose drawing part was parsed at read time (GH-221).
 * @param drawingSnapshots
 *   As-parsed `Sheet.drawings` vectors by 0-based sheet index — the SAME references stored on the
 *   sheets, so the writer's snapshot-equality dirty test hits the reference-equality fast path for
 *   untouched sheets (GH-221).
 * @param chartSnapshots
 *   As-parsed typed-chart provenance by 0-based sheet index (GH-222) — feeds the writer's
 *   equality-match part/rel-id reuse for dirty drawing parts.
 */
final case class SourceContext(
  sourcePath: Path,
  partManifest: PartManifest,
  modificationTracker: ModificationTracker,
  fingerprint: SourceFingerprint,
  commentPathMapping: Map[Int, String] = Map.empty,
  drawingPathMapping: Map[Int, String] = Map.empty,
  drawingSnapshots: Map[Int, Vector[Drawing]] = Map.empty,
  chartSnapshots: Map[Int, Vector[ChartSnapshot]] = Map.empty
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
   * @param drawingPathMapping
   *   Mapping from 0-based sheet index to drawing part path (GH-221).
   * @param drawingSnapshots
   *   As-parsed drawings vectors by 0-based sheet index (GH-221).
   * @param chartSnapshots
   *   As-parsed typed-chart provenance by 0-based sheet index (GH-222).
   */
  def fromFile(
    path: Path,
    manifest: PartManifest,
    fingerprint: SourceFingerprint,
    commentPathMapping: Map[Int, String] = Map.empty,
    drawingPathMapping: Map[Int, String] = Map.empty,
    drawingSnapshots: Map[Int, Vector[Drawing]] = Map.empty,
    chartSnapshots: Map[Int, Vector[ChartSnapshot]] = Map.empty
  ): SourceContext =
    SourceContext(
      path,
      manifest,
      ModificationTracker.clean,
      fingerprint,
      commentPathMapping,
      drawingPathMapping,
      drawingSnapshots,
      chartSnapshots
    )
