package com.tjclp.xl.cli.read

import cats.effect.IO
import cats.syntax.all.*
import fs2.{Chunk, Stream}

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue, FormulaKind}
import com.tjclp.xl.cli.ViewFormat
import com.tjclp.xl.cli.contract.{CliError, CliException, ErrorCode, Warning, WarningCode}
import com.tjclp.xl.cli.helpers.Resolve
import com.tjclp.xl.cli.output.RendererCommon
import com.tjclp.xl.cli.raster.{RasterFormat, RasterizerChain}
import com.tjclp.xl.formula.{DependencyGraph, SheetEvaluator}
import com.tjclp.xl.styles.CellStyle

/**
 * The loaded-workbook strategy: every capability. Records are projected straight from the `Sheet`;
 * `view --eval` evaluates the window's formulas (and their closure) with the workbook as
 * cross-sheet context; `cell` walks the bounded cross-sheet graph.
 */
final class InMemorySource(wb: Workbook) extends SheetSource:

  val capabilities: Set[Capability] = Capability.inMemory

  def sheets: IO[Vector[SheetName]] = IO.pure(wb.sheets.map(_.name))

  def usedRange(sheet: SheetName): IO[Option[CellRange]] =
    named(sheet).map(InMemorySource.dimension)

  def rows(sheet: SheetName, window: CellRange): Stream[IO, Vector[CellRecord]] =
    Stream.eval(named(sheet)).flatMap(s => InMemorySource.rows(s, window))

  def hiddenLines(sheet: SheetName, window: CellRange): IO[(Set[Int], Set[Int])] =
    named(sheet).map(InMemorySource.hiddenLines(_, window))

  override def grid(sheet: SheetName, window: CellRange): IO[RecordGrid] =
    named(sheet).map(InMemorySource.grid(_, window))

  def evaluated(
    sheet: SheetName,
    window: CellRange,
    strict: Boolean,
    warn: Warning => IO[Unit]
  ): IO[RecordGrid] =
    named(sheet).flatMap { s =>
      InMemorySource
        .evaluateSheetFormulas(s, Some(wb), Some(window), strict, warn)
        .map(InMemorySource.grid(_, window))
    }

  def occupied(sheet: SheetName): Stream[IO, CellRecord] =
    Stream.eval(named(sheet)).flatMap { s =>
      // A cell that holds no value — a style-only record — is not occupied: the streaming reader
      // never emits one, and `search` has nothing to match in it. Row-major order over a hash map
      // means a sort; the key is one primitive per cell
      val ordered = s.cells.toVector
        .filter { (_, cell) => InMemorySource.holdsValue(cell) }
        .sortBy { (ref, _) => InMemorySource.rowMajor(ref) }
      val merges = InMemorySource.MergeIndex(s.mergedRanges)
      val hidden = InMemorySource.HiddenLines(s)
      Stream.emits(ordered).map { (ref, cell) =>
        CellRecord.of(
          s.name,
          ref,
          cell.value,
          InMemorySource.styleOf(s, cell),
          Some(hidden.contains(ref)),
          merges.at(ref)
        )
      }
    }

  def detail(sheet: SheetName, ref: ARef, withStyle: Boolean): IO[CellDetail] =
    named(sheet).map { s =>
      val cell = s.cells.get(ref)
      val style = if withStyle then cell.flatMap(InMemorySource.styleOf(s, _)) else None
      val record = CellRecord.of(
        s.name,
        ref,
        cell.fold(CellValue.Empty)(_.value),
        style,
        Some(InMemorySource.isHidden(s, ref)),
        s.getMergedRange(ref)
      )
      // ADR-017 §2.10: the bounded cross-sheet graph — precedents at cell granularity (ranges as
      // their occupied cells), dependents through the symbolic range index. Same-sheet refs are
      // unqualified; cross-sheet ones carry the sheet spelled as a formula would (`QualifiedRef`'s
      // own rendering, the printer `deps` uses): `'On-Premise'!G9`, not `On-Premise!G9` (GH-609).
      // Ordered by (sheet, row, column) BEFORE rendering, so the order owes nothing to quoting.
      val current = DependencyGraph.QualifiedRef(s.name, ref)
      val graph = QualifiedGraph.of(wb)
      def show(q: DependencyGraph.QualifiedRef): String =
        if q.sheet == s.name then q.ref.toA1 else q.toString
      def listed(qs: Iterable[DependencyGraph.QualifiedRef]): Vector[String] =
        qs.toVector.sortBy(q => (q.sheet.value, q.ref.row.index0, q.ref.col.index0)).map(show)
      CellDetail(
        record,
        s.getComment(ref),
        cell.flatMap(_.hyperlink),
        Some(listed(graph.precedentsOf(current))),
        Some(listed(graph.dependentsOf(current)))
      )
    }

  def render(
    sheet: SheetName,
    window: CellRange,
    spec: RenderSpec,
    warn: Warning => IO[Unit]
  ): IO[String] =
    named(sheet).flatMap { s =>
      val theme = wb.metadata.theme
      // `gate` is whether an evaluation failure is the --strict gate or an EVAL_FAILED warning
      def evaluated(gate: Boolean): IO[Sheet] =
        if spec.evalFormulas then
          InMemorySource.evaluateSheetFormulas(s, Some(wb), Some(window), gate, warn)
        else IO.pure(s)
      spec.format match
        case ViewFormat.Html =>
          evaluated(spec.strict).map(
            _.toHtml(
              window,
              theme = theme,
              applyPrintScale = spec.printScale,
              showLabels = spec.showLabels
            )
          )
        case ViewFormat.Svg =>
          evaluated(spec.strict).map(
            _.toSvg(
              window,
              theme = theme,
              showGridlines = spec.showGridlines,
              showLabels = spec.showLabels
            )
          )
        case ViewFormat.Png | ViewFormat.Jpeg | ViewFormat.WebP | ViewFormat.Pdf =>
          spec.rasterOutput match
            case None =>
              IO.raiseError(
                CliException(
                  CliError.usage(
                    s"--raster-output required for ${spec.format.toString.toLowerCase} format (binary output cannot go to stdout)",
                    None
                  )
                )
              )
            case Some(outputPath) =>
              // Raster formats have never gated on --strict: an evaluation failure warns
              evaluated(false).flatMap { rendered =>
                val svg = rendered.toSvg(
                  window,
                  theme = theme,
                  showGridlines = spec.showGridlines,
                  showLabels = spec.showLabels
                )
                val rasterFormat = spec.format match
                  case ViewFormat.Jpeg => RasterFormat.Jpeg(spec.quality)
                  case ViewFormat.WebP => RasterFormat.WebP
                  case ViewFormat.Pdf => RasterFormat.Pdf
                  case _ => RasterFormat.Png
                RasterizerChain
                  .convert(svg, outputPath, rasterFormat, spec.dpi, spec.rasterizer)
                  .map { used =>
                    s"Exported: $outputPath (${spec.format.toString.toLowerCase}, ${spec.dpi} DPI, $used)"
                  }
              }
        case ViewFormat.Markdown | ViewFormat.Json | ViewFormat.Csv =>
          IO.raiseError(
            CliException(
              CliError(
                ErrorCode.INTERNAL,
                s"render called for the table format ${spec.format.toString.toLowerCase}"
              )
            )
          )
    }

  private def named(sheet: SheetName): IO[Sheet] =
    IO.fromEither(Resolve.named(wb, sheet).left.map(CliException(_)))

object InMemorySource:

  /** The projection of a loaded sheet's window: what the sheet-based renderer adapters use. */
  def grid(sheet: Sheet, window: CellRange): RecordGrid =
    val (hiddenRows, hiddenCols) = hiddenLines(sheet, window)
    val project = rowProjection(sheet, window, hiddenRows, hiddenCols)
    val rows = (window.start.row.index0 to window.end.row.index0).toVector.map(project)
    RecordGrid(sheet.name, window, rows, hiddenRows, hiddenCols)

  /**
   * [[grid]] with the window's formula cells evaluated one by one (the renderer-level `--eval` the
   * JSON tests pin): the computed value, or the Excel-style error token when the evaluation fails.
   */
  def evaluatedGrid(sheet: Sheet, window: CellRange): RecordGrid =
    val base = grid(sheet, window)
    // GH-430: a dataTable record is never evaluated — its cache IS the value
    def evaluable(kind: FormulaKind): Boolean = kind match
      case _: FormulaKind.DataTable => false
      case _ => true
    base.copy(rows = base.rows.map(_.map { record =>
      record.formula match
        case Some(f) if evaluable(f.kind) =>
          val result = SheetEvaluator
            .evaluateFormula(sheet)(f.text)
            .left
            .map(err => RendererCommon.formatEvalError(err.message))
          CellRecord.evaluated(record, result)
        case _ => record
    }))

  /**
   * The window's dense rows, [[rowChunk]] rows at a time: `stats` over a full column and `filter`
   * fold a million rows through here holding one chunk of them, where a `Vector` of every row
   * (fs2's `emits` takes its rows by value) would be the whole column's records in memory before
   * the first was seen.
   */
  def rows(sheet: Sheet, window: CellRange): Stream[IO, Vector[CellRecord]] =
    val (hiddenRows, hiddenCols) = hiddenLines(sheet, window)
    val project = rowProjection(sheet, window, hiddenRows, hiddenCols)
    val lastRow = window.end.row.index0
    Stream.unfoldChunk(window.start.row.index0) { row =>
      Option.when(row <= lastRow) {
        val until = math.min(row + rowChunk, lastRow + 1)
        (Chunk.from((row until until).map(project)), until)
      }
    }

  /** How many rows of a window [[rows]] holds at once. */
  private[read] val rowChunk: Int = 256

  /**
   * The bounding box of every cell the sheet stores, styled-but-empty ones included — the
   * `<dimension>` the library's writer records, so `view` without a range and `filter` address the
   * same window from both sources for a file it wrote. (`Sheet.usedRange` spans the non-empty cells
   * only; another producer's `<dimension>` is whatever it wrote, see `StreamingSource`.)
   */
  def dimension(sheet: Sheet): Option[CellRange] =
    sheet.cells.keysIterator
      .foldLeft(Option.empty[(Int, Int, Int, Int)]) { (acc, ref) =>
        val (c, r) = (ref.col.index0, ref.row.index0)
        Some(acc.fold((c, r, c, r)) { case (c1, r1, c2, r2) =>
          (c1 min c, r1 min r, c2 max c, r2 max r)
        })
      }
      .map { case (c1, r1, c2, r2) => CellRange(ARef.from0(c1, r1), ARef.from0(c2, r2)) }

  /**
   * One row of the window as dense records, with the state every row shares — the window's columns,
   * its hidden lines and the merges that touch it — computed once, not per row.
   */
  private def rowProjection(
    sheet: Sheet,
    window: CellRange,
    hiddenRows: Set[Int],
    hiddenCols: Set[Int]
  ): Int => Vector[CellRecord] =
    val cols = (window.start.col.index0 to window.end.col.index0).toVector
    // Only the merges that touch the window can contain one of its cells
    val merges = MergeIndex(sheet.mergedRanges.filter(m => intersects(m, window)))
    row =>
      val rowHidden = hiddenRows.contains(row)
      cols.map { col =>
        val ref = ARef.from0(col, row)
        CellRecord.project(
          sheet.name,
          ref,
          sheet.cells.get(ref),
          styleOf(sheet, _),
          Some(rowHidden || hiddenCols.contains(col)),
          merges.at(ref)
        )
      }

  /**
   * Merged ranges indexed by the rows they cover, so the merge containing a cell is found among the
   * few ranges on its row rather than by scanning them all (`search` over 60k cells with 5k merged
   * headers was O(cells × merges)). A range taller than [[MergeIndex.tallRows]] is kept aside and
   * scanned directly, so a whole-column merge does not index a million rows.
   */
  final class MergeIndex private (byRow: Map[Int, Vector[CellRange]], tall: Vector[CellRange]):
    def at(ref: ARef): Option[CellRange] =
      byRow
        .get(ref.row.index0)
        .flatMap(_.find(_.contains(ref)))
        .orElse(tall.find(_.contains(ref)))

    /** The ranges kept aside as tall (the two buckets are the test's business, not a caller's). */
    private[read] def tallRanges: Vector[CellRange] = tall

    /** How many rows the short ranges are indexed under. */
    private[read] def indexedRows: Int = byRow.size

  object MergeIndex:
    val tallRows: Int = 64

    def apply(merges: Iterable[CellRange]): MergeIndex =
      val (tall, short) = merges.toVector.partition(_.height > tallRows)
      val byRow = short
        .flatMap(m => (m.start.row.index0 to m.end.row.index0).map(_ -> m))
        .groupMap(_._1)(_._2)
      new MergeIndex(byRow, tall)

  /** Whether the cell holds a value (a type test — `!= Empty` would run every value's `equals`). */
  private def holdsValue(cell: Cell): Boolean = cell.value match
    case CellValue.Empty => false
    case _ => true

  /** The row-major sort key of a ref — one primitive, so a 60k-cell sort allocates no tuples. */
  private def rowMajor(ref: ARef): Long = (ref.row.index0.toLong << 32) | ref.col.index0.toLong

  /**
   * The sheet's hidden rows and columns, read once: `getRowProperties`/`getColumnProperties` build
   * a default record on every miss, which over a 60k-cell scan is 120k allocations for nothing.
   */
  final class HiddenLines private (val rows: Set[Int], val cols: Set[Int]):
    def contains(ref: ARef): Boolean =
      (rows.nonEmpty && rows.contains(ref.row.index0)) ||
        (cols.nonEmpty && cols.contains(ref.col.index0))

  object HiddenLines:
    def apply(sheet: Sheet): HiddenLines =
      new HiddenLines(
        sheet.rowProperties.collect { case (row, p) if p.hidden => row.index0 }.toSet,
        sheet.columnProperties.collect { case (col, p) if p.hidden => col.index0 }.toSet
      )

  /**
   * The window's hidden rows and columns: the sheet's ([[HiddenLines]], read once off the property
   * maps) clipped to the window — never a `getRowProperties` per row of a full column.
   */
  def hiddenLines(sheet: Sheet, window: CellRange): (Set[Int], Set[Int]) =
    val hidden = HiddenLines(sheet)
    (
      hidden.rows.filter(r => r >= window.start.row.index0 && r <= window.end.row.index0),
      hidden.cols.filter(c => c >= window.start.col.index0 && c <= window.end.col.index0)
    )

  private def isHidden(sheet: Sheet, ref: ARef): Boolean =
    sheet.getRowProperties(ref.row).hidden || sheet.getColumnProperties(ref.col).hidden

  private def styleOf(sheet: Sheet, cell: Cell): Option[CellStyle] =
    cell.styleId.flatMap(sheet.styleRegistry.get)

  private def intersects(a: CellRange, b: CellRange): Boolean =
    a.start.col.index0 <= b.end.col.index0 && b.start.col.index0 <= a.end.col.index0 &&
      a.start.row.index0 <= b.end.row.index0 && b.start.row.index0 <= a.end.row.index0

  /**
   * Evaluate the formula cells of a sheet, replacing them with their computed values so a render
   * shows live numbers (`view --eval`). Dependency-aware: formulas are evaluated in topological
   * order, and with a range only the formulas inside it (plus their transitive dependencies) are
   * evaluated.
   *
   * @param strict
   *   whether an evaluation failure is the `--strict` gate (`RECALC_GATE`, exit 1) or an
   *   `EVAL_FAILED` warning through `warn`, with the original sheet rendered from its caches
   */
  def evaluateSheetFormulas(
    sheet: Sheet,
    workbook: Option[Workbook],
    range: Option[CellRange],
    strict: Boolean,
    warn: Warning => IO[Unit]
  ): IO[Sheet] =
    val evalResult = range match
      case Some(r) => SheetEvaluator.evaluateForRange(sheet)(r, workbook = workbook)
      case None => SheetEvaluator.evaluateWithDependencyCheck(sheet)(workbook = workbook)
    evalResult match
      case Right(results) =>
        // Sheet.put preserves the existing cell styleId
        IO.pure(results.foldLeft(sheet) { case (acc, (ref, value)) => acc.put(ref, value) })
      case Left(error) =>
        // `view --eval --strict` is a user-requested gate (ADR-017 invariant 5): exit 1 with
        // RECALC_GATE, never a failure. The message text is unchanged.
        if strict then
          IO.raiseError(
            CliException(
              CliError(
                ErrorCode.RECALC_GATE,
                s"Formula evaluation failed: ${error.message}",
                hint =
                  Some("drop --strict to render cached values and see the failure as a warning")
              )
            )
          )
        else
          warn(Warning(WarningCode.EVAL_FAILED, s"Formula evaluation failed: ${error.message}"))
            .as(sheet)
