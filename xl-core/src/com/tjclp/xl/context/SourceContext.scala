package com.tjclp.xl.context

import com.tjclp.xl.addressing.SheetName
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
 * Per-sheet source mappings are keyed by STABLE SHEET IDENTITY — the sheet NAME as read (unique per
 * workbook), kept current through tracked edits (GH-315):
 *   - [[markSheetRenamed]] re-keys the mappings, so identity follows a `Workbook.rename`
 *   - [[markSheetDeleted]] drops the deleted sheet's entries, so a later sheet with the same name
 *     is a NEW sheet (no source identity)
 *   - reorders need nothing: names do not move
 *
 * Sheets surgically replaced without the tracked Workbook operations lose their source identity
 * along with the rest of their modification tracking (the existing documented caveat).
 *
 * @param commentPathMapping
 *   Sheet name (as read) → comment file path (e.g., "xl/comments1.xml"). Excel numbers comment
 *   files sequentially across only sheets that have comments, NOT by sheet index. This mapping
 *   preserves the original paths to enable correct surgical writes.
 * @param drawingPathMapping
 *   Sheet name (as read) → drawing part path (e.g., "xl/drawings/drawing1.xml") for sheets whose
 *   drawing part was parsed at read time (GH-221).
 * @param drawingSnapshots
 *   As-parsed `Sheet.drawings` vectors by sheet name — the SAME references stored on the sheets, so
 *   the writer's snapshot-equality dirty test hits the reference-equality fast path for untouched
 *   sheets (GH-221).
 * @param chartSnapshots
 *   As-parsed typed-chart provenance by sheet name (GH-222) — feeds the writer's equality-match
 *   part/rel-id reuse for dirty drawing parts.
 * @param sheetPathMapping
 *   Sheet name (as read) → worksheet part path (e.g., "xl/worksheets/sheet2.xml"). The writer
 *   resolves every per-sheet source lookup (preserved metadata, rels, SST reference accounting)
 *   through this mapping, so structural edits never alias a sheet to another sheet's part (GH-315).
 *   Empty for contexts built before the mapping existed — consumers fall back to index-based naming
 *   in that case.
 */
final case class SourceContext(
  sourcePath: Path,
  partManifest: PartManifest,
  modificationTracker: ModificationTracker,
  fingerprint: SourceFingerprint,
  commentPathMapping: Map[SheetName, String] = Map.empty,
  drawingPathMapping: Map[SheetName, String] = Map.empty,
  drawingSnapshots: Map[SheetName, Vector[Drawing]] = Map.empty,
  chartSnapshots: Map[SheetName, Vector[ChartSnapshot]] = Map.empty,
  sheetPathMapping: Map[SheetName, String] = Map.empty
) derives CanEqual:

  /** True when no workbook modifications have been recorded. */
  def isClean: Boolean = modificationTracker.isClean

  /** Mark a sheet as modified. */
  def markSheetModified(sheetIndex: Int): SourceContext =
    copy(modificationTracker = modificationTracker.markSheet(sheetIndex))

  /**
   * Mark a sheet as deleted, dropping its identity-keyed source mappings (GH-315). A sheet added
   * later under the same name is a NEW sheet — it must not inherit the deleted sheet's parts.
   */
  def markSheetDeleted(sheetIndex: Int, name: SheetName): SourceContext =
    copy(
      modificationTracker = modificationTracker.delete(sheetIndex),
      commentPathMapping = commentPathMapping - name,
      drawingPathMapping = drawingPathMapping - name,
      drawingSnapshots = drawingSnapshots - name,
      chartSnapshots = chartSnapshots - name,
      sheetPathMapping = sheetPathMapping - name
    )

  /**
   * Re-key the identity-keyed source mappings after a tracked rename (GH-315): identity follows the
   * sheet, so a renamed sheet keeps its source comment/drawing/worksheet parts. Callers mark the
   * sheet and metadata modified separately (`Workbook.rename` does both).
   */
  def markSheetRenamed(oldName: SheetName, newName: SheetName): SourceContext =
    def rekey[A](m: Map[SheetName, A]): Map[SheetName, A] =
      m.get(oldName) match
        case Some(value) => m - oldName + (newName -> value)
        case None => m
    copy(
      commentPathMapping = rekey(commentPathMapping),
      drawingPathMapping = rekey(drawingPathMapping),
      drawingSnapshots = rekey(drawingSnapshots),
      chartSnapshots = rekey(chartSnapshots),
      sheetPathMapping = rekey(sheetPathMapping)
    )

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
   * All per-sheet mappings are keyed by the sheet NAME as read (stable identity, GH-315); see the
   * [[SourceContext]] scaladoc for the rename/delete semantics.
   *
   * @param commentPathMapping
   *   Sheet name → comment file path. Excel numbers comment files sequentially (comments1.xml,
   *   comments2.xml...) across only sheets that have comments.
   * @param drawingPathMapping
   *   Sheet name → drawing part path (GH-221).
   * @param drawingSnapshots
   *   As-parsed drawings vectors by sheet name (GH-221).
   * @param chartSnapshots
   *   As-parsed typed-chart provenance by sheet name (GH-222).
   * @param sheetPathMapping
   *   Sheet name → worksheet part path (GH-315).
   */
  def fromFile(
    path: Path,
    manifest: PartManifest,
    fingerprint: SourceFingerprint,
    commentPathMapping: Map[SheetName, String] = Map.empty,
    drawingPathMapping: Map[SheetName, String] = Map.empty,
    drawingSnapshots: Map[SheetName, Vector[Drawing]] = Map.empty,
    chartSnapshots: Map[SheetName, Vector[ChartSnapshot]] = Map.empty,
    sheetPathMapping: Map[SheetName, String] = Map.empty
  ): SourceContext =
    SourceContext(
      path,
      manifest,
      ModificationTracker.clean,
      fingerprint,
      commentPathMapping,
      drawingPathMapping,
      drawingSnapshots,
      chartSnapshots,
      sheetPathMapping
    )
