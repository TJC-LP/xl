package com.tjclp.xl.context

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.charts.Chart
import com.tjclp.xl.drawings.Drawing
import com.tjclp.xl.ooxml.PartManifest
import com.tjclp.xl.workbooks.{DefinedName, Workbook}

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
 * Physical source archive a workbook was read from.
 *
 * OnDisk sources are re-opened by path for preserved-part access; InMemory sources (GH-412) carry
 * the full archive bytes, so byte-array reads preserve unknown parts exactly like path reads — a
 * workbook read from bytes and one read from a path over the same content write byte-identically.
 */
enum SourceContent derives CanEqual:
  case OnDisk(path: Path)
  case InMemory(bytes: ArraySeq[Byte])

object SourceContent:
  /**
   * Zero-copy array view over in-memory archive bytes. The reader wraps a private clone at the
   * byte-read boundary, so exposing the backing array within xl never aliases caller-owned memory.
   */
  private[xl] def rawArray(bytes: ArraySeq[Byte]): Array[Byte] = bytes match
    case b: ArraySeq.ofByte => b.unsafeArray
    case other => other.toArray

/**
 * Captures metadata about the physical XLSX that produced a [[Workbook]]. The context enables
 * surgical write operations by preserving the manifest of ZIP entries. Preserved parts are read
 * back from [[SourceContent]] (the source file, or the in-memory archive) when needed for IO
 * operations.
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
 * @param definedNamesAsRead
 *   The `metadata.definedNames` vector exactly as the reader assembled it (GH-470, post print-name
 *   extraction) — the baseline [[reconciledWith]] compares against to detect defined-name edits
 *   made through a raw `Workbook.copy(metadata = ...)` instead of the tracked API. None means "no
 *   snapshot" (contexts built outside the reader): the divergence check is skipped.
 */
final case class SourceContext(
  content: SourceContent,
  partManifest: PartManifest,
  modificationTracker: ModificationTracker,
  fingerprint: SourceFingerprint,
  commentPathMapping: Map[SheetName, String] = Map.empty,
  drawingPathMapping: Map[SheetName, String] = Map.empty,
  drawingSnapshots: Map[SheetName, Vector[Drawing]] = Map.empty,
  chartSnapshots: Map[SheetName, Vector[ChartSnapshot]] = Map.empty,
  sheetPathMapping: Map[SheetName, String] = Map.empty,
  definedNamesAsRead: Option[Vector[DefinedName]] = None
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
   * Mark a sheet removal that was applied OUTSIDE the tracked edit API (GH-470): the sheet is
   * already gone from `Workbook.sheets`, so unlike [[markSheetDeleted]] no live-index shifting may
   * touch the modified-sheet marks — they are in the already-reduced index space. `sourceIndex` is
   * the sheet's best-effort position in the SOURCE package. The identity-keyed mappings drop
   * exactly like a tracked delete, so the writer's surviving-sheet resolution and orphan pruning
   * (GH-417) treat the sheet as removed.
   */
  def markUntrackedSheetRemoval(sourceIndex: Int, name: SheetName): SourceContext =
    copy(
      modificationTracker = modificationTracker.deleteUntracked(sourceIndex),
      commentPathMapping = commentPathMapping - name,
      drawingPathMapping = drawingPathMapping - name,
      drawingSnapshots = drawingSnapshots - name,
      chartSnapshots = chartSnapshots - name,
      sheetPathMapping = sheetPathMapping - name
    )

  /**
   * Reconcile untracked structural divergence between the workbook MODEL and this context (GH-470).
   * Raw `Workbook.copy(sheets = ...)` / `copy(metadata = ...)` edits bypass the tracked API,
   * leaving a clean-looking tracker over a structurally different workbook — a preservation write
   * would then re-emit the SOURCE structure, silently shipping sheets or defined names the caller
   * believes were removed (a confidentiality hazard) and dropping sheets the caller appended. The
   * writer calls this before choosing a write strategy:
   *
   *   - a source sheet (identity-keyed `sheetPathMapping` entry) with no surviving model sheet is
   *     marked deleted exactly like `Workbook.remove` marks it — tracker deletion plus mapping
   *     drops — so surviving sheets regenerate against their own source parts and the removed
   *     sheet's part/rel/media closure is pruned (GH-417);
   *   - a model sheet with no source identity (untracked addition) marks metadata modified, so the
   *     workbook skeleton regenerates from the model instead of a verbatim copy dropping the sheet;
   *   - a defined-name set diverging from [[definedNamesAsRead]] marks metadata modified, so
   *     workbook.xml re-derives from the model.
   *
   * Returns `this` unchanged when nothing diverges — a genuinely clean workbook keeps the verbatim
   * fast path. Contexts without an identity mapping (pre-GH-315) skip the sheet-set checks.
   */
  def reconciledWith(workbook: Workbook): SourceContext =
    val liveNames = workbook.sheets.map(_.name).toSet
    val ghosts = sheetPathMapping.keySet.diff(liveNames).toVector.sortBy(_.value)
    val withRemovals = ghosts.zipWithIndex.foldLeft(this) { case (ctx, (ghost, ordinal)) =>
      val sourceIndex = ctx.sheetPathMapping
        .get(ghost)
        .flatMap(SourceContext.worksheetPartNumber)
        .getOrElse(workbook.sheets.size + ordinal)
      ctx.markUntrackedSheetRemoval(sourceIndex, ghost)
    }
    val untrackedAddition =
      sheetPathMapping.nonEmpty && workbook.sheets.exists(s => !sheetPathMapping.contains(s.name))
    val definedNamesDiverged =
      definedNamesAsRead.exists(_ != workbook.metadata.definedNames)
    if untrackedAddition || definedNamesDiverged then withRemovals.markMetadataModified
    else withRemovals

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

  private val WorksheetPartNumber = raw"xl/worksheets/sheet(\d+)\.xml".r

  /**
   * Zero-based source sheet index implied by the conventional worksheet part name (GH-470:
   * `xl/worksheets/sheet3.xml` → 2). None for foreign packages with unconventional part names — the
   * caller falls back to a synthetic non-colliding index.
   */
  private[context] def worksheetPartNumber(path: String): Option[Int] = path match
    case WorksheetPartNumber(n) => n.toIntOption.map(_ - 1).filter(_ >= 0)
    case _ => None

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
    fromContent(
      SourceContent.OnDisk(path),
      manifest,
      fingerprint,
      commentPathMapping,
      drawingPathMapping,
      drawingSnapshots,
      chartSnapshots,
      sheetPathMapping
    )

  /**
   * Construct a context over any [[SourceContent]] — a file on disk, or the in-memory archive of a
   * byte-array read (GH-412). Same identity-keyed mapping semantics as [[fromFile]].
   */
  def fromContent(
    content: SourceContent,
    manifest: PartManifest,
    fingerprint: SourceFingerprint,
    commentPathMapping: Map[SheetName, String] = Map.empty,
    drawingPathMapping: Map[SheetName, String] = Map.empty,
    drawingSnapshots: Map[SheetName, Vector[Drawing]] = Map.empty,
    chartSnapshots: Map[SheetName, Vector[ChartSnapshot]] = Map.empty,
    sheetPathMapping: Map[SheetName, String] = Map.empty
  ): SourceContext =
    SourceContext(
      content,
      manifest,
      ModificationTracker.clean,
      fingerprint,
      commentPathMapping,
      drawingPathMapping,
      drawingSnapshots,
      chartSnapshots,
      sheetPathMapping
    )
