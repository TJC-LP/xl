package com.tjclp.xl.cli.read

import cats.effect.IO
import cats.syntax.all.*
import fs2.Stream

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
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

  def usedRange(sheet: SheetName): IO[Option[CellRange]] = named(sheet).map(_.usedRange)

  def rows(sheet: SheetName, window: CellRange): Stream[IO, Vector[CellRecord]] =
    Stream.eval(named(sheet)).flatMap(s => InMemorySource.rows(s, window))

  def hiddenLines(sheet: SheetName, window: CellRange): IO[(Set[Int], Set[Int])] =
    named(sheet).map(InMemorySource.hiddenLines(_, window))

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
      val ordered = s.cells.toVector.sortBy { (ref, _) => (ref.row.index0, ref.col.index0) }
      Stream.emits(ordered).map { (ref, cell) =>
        CellRecord.of(
          s.name,
          ref,
          cell.value,
          InMemorySource.styleOf(s, cell),
          InMemorySource.isHidden(s, ref),
          s.getMergedRange(ref)
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
        InMemorySource.isHidden(s, ref),
        s.getMergedRange(ref)
      )
      // ADR-017 §2.10: the bounded cross-sheet graph — precedents at cell granularity (ranges as
      // their occupied cells), dependents through the symbolic range index. Same-sheet refs are
      // unqualified, cross-sheet ones carry the sheet.
      val current = DependencyGraph.QualifiedRef(s.name, ref)
      val graph = QualifiedGraph.of(wb)
      def show(q: DependencyGraph.QualifiedRef): String =
        if q.sheet == s.name then q.ref.toA1 else s"${q.sheet.value}!${q.ref.toA1}"
      CellDetail(
        record,
        s.getComment(ref),
        cell.flatMap(_.hyperlink),
        graph.precedentsOf(current).toVector.map(show).sorted,
        Some(graph.dependentsOf(current).toVector.map(show).sorted)
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
    RecordGrid(sheet.name, window, denseRows(sheet, window), hiddenRows, hiddenCols)

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

  def rows(sheet: Sheet, window: CellRange): Stream[IO, Vector[CellRecord]] =
    Stream.emits(denseRows(sheet, window))

  private def denseRows(sheet: Sheet, window: CellRange): Vector[Vector[CellRecord]] =
    val cols = (window.start.col.index0 to window.end.col.index0).toVector
    val (hiddenRows, hiddenCols) = hiddenLines(sheet, window)
    // Only the merges that touch the window can contain one of its cells
    val merges = sheet.mergedRanges.toVector.filter(m => intersects(m, window))
    (window.start.row.index0 to window.end.row.index0).toVector.map { row =>
      val rowHidden = hiddenRows.contains(row)
      cols.map { col =>
        val ref = ARef.from0(col, row)
        CellRecord.project(
          sheet.name,
          ref,
          sheet.cells.get(ref),
          styleOf(sheet, _),
          rowHidden || hiddenCols.contains(col),
          merges.find(_.contains(ref))
        )
      }
    }

  def hiddenLines(sheet: Sheet, window: CellRange): (Set[Int], Set[Int]) =
    val rows = (window.start.row.index0 to window.end.row.index0)
      .filter(r => sheet.getRowProperties(Row.from0(r)).hidden)
      .toSet
    val cols = (window.start.col.index0 to window.end.col.index0)
      .filter(c => sheet.getColumnProperties(Column.from0(c)).hidden)
      .toSet
    (rows, cols)

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
