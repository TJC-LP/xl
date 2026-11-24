package com.tjclp.xl.sheets

import scala.util.Try

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue, FormulaKind}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.ops.{ClearWhat, ColSpan, Edit, FormulaSupport, OffGridRef, RowSpan}
import com.tjclp.xl.render.RenderUtils
import com.tjclp.xl.sheets.styleSyntax.withCellStyle
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * The sheet-local edit semantics the CLI used to own (ADR-017 §2.12, W2.1): fill, copy, sort,
 * clear, column auto-fit, row/column outline grouping and the appearance/print merges. Reached
 * through the `Sheet` methods of the same names; the CLI's `WriteCommands`/`CopyOps`/
 * `GroupingOps`/`AppearanceOps`/`ColumnAutoFit` forward here, so the two paths cannot drift.
 *
 * Formula text is rewritten through [[FormulaSupport.shift]]: a refusal (`UnsupportedCapability`)
 * fails the edit, while text the engine cannot parse is kept as written — an unsupported expression
 * copies verbatim, exactly as the CLI has always done.
 */
private[xl] object SheetEdits:

  // ===== Formula shifting shared by fill / copy / sort =====

  /**
   * The value a cell takes when it moves to `target` by `(colDelta, rowDelta)`, with the references
   * the shift voided (GH-628). A data-table record always copies as its cached constant (GH-430:
   * pasting the TABLE(...) display text would be a `#NAME?` bomb, and the record would claim a
   * table interior that does not exist at the target). A formula is shifted; its cache is dropped
   * because it no longer describes the shifted text.
   *
   * A single-cell array formula (a 1x1 CSE record) stays one, re-anchored at `target`: Excel pastes
   * and fills `{=...}` as `{=...}`, and a plain paste would be a legacy formula computing another
   * value. A multi-cell CSE anchor pastes a plain formula, since its record claims cells the paste
   * does not write. A dynamic-array anchor of any size pastes as a 1x1 dynamic array at `target`
   * (GH-714), unless the copy includes its entire spill. A complete copied spill keeps its shifted
   * extent so Excel recognizes the copied constants as spill members, not occupied cells.
   */
  private def shifted(
    value: CellValue,
    target: ARef,
    colDelta: Int,
    rowDelta: Int,
    copiedSource: Option[CellRange] = None
  )(using fs: FormulaSupport): XLResult[(CellValue, Vector[String])] =
    value match
      case CellValue.Formula(_, cachedOpt, _: FormulaKind.DataTable) =>
        Right((cachedOpt.getOrElse(CellValue.Empty), Vector.empty))
      case CellValue.Formula(expr, _, kind) =>
        val pastedKind = kind match
          case arr: FormulaKind.ArrayFormula if arr.isDynamicArray =>
            val extent = for
              source <- copiedSource
              if source.contains(arr.ref.start) && source.contains(arr.ref.end)
              start <- arr.ref.start.tryShift(colDelta, rowDelta)
              end <- arr.ref.end.tryShift(colDelta, rowDelta)
            yield CellRange(start, end)
            arr.copy(ref = extent.getOrElse(CellRange(target, target)))
          case arr: FormulaKind.ArrayFormula if arr.ref.width == 1 && arr.ref.height == 1 =>
            arr.copy(ref = CellRange(target, target))
          case _ => FormulaKind.Normal()
        fs.shiftReporting(expr, colDelta, rowDelta) match
          case Right(FormulaSupport.Shifted(text, voided)) =>
            Right((CellValue.Formula(text, None, pastedKind), voided))
          case Left(refusal: XLError.UnsupportedCapability) => Left(refusal)
          // unparseable: copied as written
          case Left(_) => Right((CellValue.Formula(expr, None, pastedKind), Vector.empty))
      case other => Right((other, Vector.empty))

  /** `foldLeft` that stops at the first `Left`. */
  private def foldE[A, B](as: Iterator[A])(zero: B)(f: (B, A) => XLResult[B]): XLResult[B] =
    as.foldLeft[XLResult[B]](Right(zero)) {
      case (Right(acc), a) => f(acc, a)
      case (left, _) => left
    }

  /** A sheet being edited plus the off-grid references its formula shifts have produced so far. */
  private type Reporting = (Sheet, Vector[OffGridRef])

  private def recordOffGrid(acc: Vector[OffGridRef], cell: ARef, voided: Vector[String]) =
    if voided.isEmpty then acc else acc :+ OffGridRef(cell, voided)

  // ===== Fill =====

  /** The direction rule: fill down needs matching columns, fill right matching rows. */
  def validateFill(source: CellRange, target: CellRange, direction: Edit.FillDir): XLResult[Unit] =
    direction match
      case Edit.FillDir.Down =>
        if source.colStart == target.colStart && source.colEnd == target.colEnd then Right(())
        else
          Left(
            XLError.InvalidReference(
              s"Fill down requires matching columns. Source: ${source.toA1}, Target: ${target.toA1}"
            )
          )
      case Edit.FillDir.Right =>
        if source.rowStart == target.rowStart && source.rowEnd == target.rowEnd then Right(())
        else
          Left(
            XLError.InvalidReference(
              s"Fill right requires matching rows. Source: ${source.toA1}, Target: ${target.toA1}"
            )
          )

  /**
   * Repeat `source` through `target`, cycling the source rows (down) or columns (right) and
   * shifting relative references by the displacement. Cells are read from the sheet as it is being
   * filled (the CLI's fold), so an overlapping fill sees its own earlier writes; a target cell on a
   * source row/column (zero displacement) is left alone. Empty source cells leave the target
   * unchanged.
   */
  def fill(
    sheet: Sheet,
    source: CellRange,
    target: CellRange,
    direction: Edit.FillDir
  )(using FormulaSupport): XLResult[Sheet] =
    fillReporting(sheet, source, target, direction).map(_._1)

  /**
   * GH-628: [[fill]], also reporting every target cell whose formula gained a `#REF!` because the
   * displacement carried a reference off the grid (Excel writes the `#REF!` silently).
   */
  def fillReporting(
    sheet: Sheet,
    source: CellRange,
    target: CellRange,
    direction: Edit.FillDir
  )(using FormulaSupport): XLResult[(Sheet, Vector[OffGridRef])] =
    validateFill(source, target, direction).flatMap { _ =>
      val zero: Reporting = (sheet, Vector.empty)
      foldE(fillSteps(source, target, direction))(zero) { (acc, step) =>
        copyCell(acc, step.source, step.target, step.colDelta, step.rowDelta)
      }
    }

  /**
   * GH-714: the cells a fill writes — the target of each step whose source cell exists, exactly as
   * [[fillReporting]] walks them (source rows or columns are never written) — so a caller guarding
   * the write cannot drift from the fill.
   */
  def fillTargets(
    sheet: Sheet,
    source: CellRange,
    target: CellRange,
    direction: Edit.FillDir
  ): XLResult[Set[ARef]] =
    validateFill(source, target, direction).map { _ =>
      fillSteps(source, target, direction)
        .filter(step => sheet.cells.contains(step.source))
        .map(_.target)
        .toSet
    }

  /** One fill step: copy `source` onto `target`, shifted by the deltas. */
  private final case class FillStep(source: ARef, target: ARef, colDelta: Int, rowDelta: Int)

  /**
   * The fill's walk: target rows (down) or columns (right) in order, each from its cyclically
   * repeated source row or column, the source's own rows or columns skipped.
   */
  private def fillSteps(
    source: CellRange,
    target: CellRange,
    direction: Edit.FillDir
  ): Iterator[FillStep] =
    direction match
      case Edit.FillDir.Down =>
        val sourceStartRow = source.rowStart.index0
        val sourceRowCount = source.rowEnd.index0 - sourceStartRow + 1
        val targetStartRow = target.rowStart.index0
        val cols = source.colStart.index0 to source.colEnd.index0
        (targetStartRow to target.rowEnd.index0).iterator.flatMap { targetRow =>
          val sourceRow = sourceStartRow + (targetRow - targetStartRow) % sourceRowCount
          val rowDelta = targetRow - sourceRow
          if rowDelta == 0 then Iterator.empty
          else
            cols.iterator.map { col =>
              FillStep(ARef.from0(col, sourceRow), ARef.from0(col, targetRow), 0, rowDelta)
            }
        }
      case Edit.FillDir.Right =>
        val sourceStartCol = source.colStart.index0
        val sourceColCount = source.colEnd.index0 - sourceStartCol + 1
        val targetStartCol = target.colStart.index0
        val rows = source.rowStart.index0 to source.rowEnd.index0
        (targetStartCol to target.colEnd.index0).iterator.flatMap { targetCol =>
          val sourceCol = sourceStartCol + (targetCol - targetStartCol) % sourceColCount
          val colDelta = targetCol - sourceCol
          if colDelta == 0 then Iterator.empty
          else
            rows.iterator.map { row =>
              FillStep(ARef.from0(sourceCol, row), ARef.from0(targetCol, row), colDelta, 0)
            }
        }

  /** One fill step: the source cell's value, shifted, onto the target (styles are not copied). */
  private def copyCell(
    acc: Reporting,
    sourceRef: ARef,
    targetRef: ARef,
    colDelta: Int,
    rowDelta: Int
  )(using FormulaSupport): XLResult[Reporting] =
    val (sheet, offGrid) = acc
    sheet.cells.get(sourceRef) match
      case None => Right(acc)
      case Some(sourceCell) =>
        shifted(sourceCell.value, targetRef, colDelta, rowDelta).map { (v, voided) =>
          (sheet.put(targetRef, v), recordOffGrid(offGrid, targetRef, voided))
        }

  // ===== Copy =====

  /**
   * Copy `sourceRange` of `sourceSheet` into `target` at `targetRange` (same dimensions; the same
   * sheet when the names agree). Source cells are snapshotted before any write, so an overlapping
   * same-sheet copy reads the pre-copy state whatever the iteration order. Empty source cells leave
   * their target unchanged; styles travel with their cells (looked up in the source registry,
   * registered in the target's). A fully included dynamic spill keeps its shifted array extent; an
   * anchor copied without its whole spill keeps a 1x1 record. `valuesOnly` pastes a formula's
   * cached value — an uncached formula pastes `Empty`; callers with an evaluator cache the source
   * first.
   */
  def copyRange(
    target: Sheet,
    sourceSheet: Sheet,
    sourceRange: CellRange,
    targetRange: CellRange,
    valuesOnly: Boolean
  )(using FormulaSupport): XLResult[Sheet] =
    copyRangeReporting(target, sourceSheet, sourceRange, targetRange, valuesOnly).map(_._1)

  /**
   * GH-628: [[copyRange]], also reporting every target cell whose formula gained a `#REF!` because
   * the displacement carried a reference off the grid.
   */
  def copyRangeReporting(
    target: Sheet,
    sourceSheet: Sheet,
    sourceRange: CellRange,
    targetRange: CellRange,
    valuesOnly: Boolean
  )(using FormulaSupport): XLResult[(Sheet, Vector[OffGridRef])] =
    val snapshot: Map[ARef, Cell] =
      sourceRange.cells.flatMap(ref => sourceSheet.cells.get(ref).map(ref -> _)).toMap
    val colDelta = targetRange.colStart.index0 - sourceRange.colStart.index0
    val rowDelta = targetRange.rowStart.index0 - sourceRange.rowStart.index0
    val working = if sourceSheet.name == target.name then sourceSheet else target
    val copiedSource = Some(sourceRange)
    val zero: Reporting = (working, Vector.empty)
    foldE(sourceRange.cells)(zero) { case ((s, offGrid), srcRef) =>
      val tgtRef = ARef.from0(srcRef.col.index0 + colDelta, srcRef.row.index0 + rowDelta)
      snapshot.get(srcRef) match
        case None => Right((s, offGrid))
        case Some(srcCell) =>
          val copied: XLResult[(CellValue, Vector[String])] = srcCell.value match
            case CellValue.Formula(_, cachedOpt, _: FormulaKind.DataTable) =>
              Right((cachedOpt.getOrElse(CellValue.Empty), Vector.empty))
            case CellValue.Formula(_, cachedOpt, _) if valuesOnly =>
              Right((cachedOpt.getOrElse(CellValue.Empty), Vector.empty))
            case other => shifted(other, tgtRef, colDelta, rowDelta, copiedSource)
          copied.map { (v, voided) =>
            val withValue = s.put(tgtRef, v)
            val styled = srcCell.styleId.flatMap(sourceSheet.styleRegistry.get) match
              case Some(style) => withValue.withCellStyle(tgtRef, style)
              case None => withValue
            (styled, recordOffGrid(offGrid, tgtRef, voided))
          }
    }

  // ===== Sort =====

  private enum SortValue:
    case Empty
    case Error
    case Num(value: Double)
    case Str(value: String)

  private def sortableValue(value: CellValue, mode: Edit.SortMode): SortValue =
    value match
      case CellValue.Empty => SortValue.Empty
      case CellValue.Text(s) =>
        if mode == Edit.SortMode.Numeric then
          s.toDoubleOption.map(SortValue.Num(_)).getOrElse(SortValue.Str(s.toLowerCase))
        else SortValue.Str(s.toLowerCase)
      case CellValue.Number(n) => SortValue.Num(n.toDouble)
      case CellValue.Bool(b) => SortValue.Num(if b then 1.0 else 0.0)
      case CellValue.DateTime(dt) => SortValue.Num(CellValue.dateTimeToExcelSerial(dt))
      case CellValue.Formula(_, Some(cached), _) => sortableValue(cached, mode)
      case CellValue.Formula(_, None, _) => SortValue.Str("")
      case CellValue.RichText(rt) => SortValue.Str(rt.toPlainText.toLowerCase)
      case CellValue.Error(_) => SortValue.Error

  /** Empty and error values sort last; in numeric mode numbers come before text. */
  private def compareSortValues(
    a: Option[SortValue],
    b: Option[SortValue],
    mode: Edit.SortMode
  ): Int =
    (a, b) match
      case (None, None) => 0
      case (None, _) => 1
      case (_, None) => -1
      case (Some(va), Some(vb)) =>
        (va, vb) match
          case (SortValue.Empty, SortValue.Empty) => 0
          case (SortValue.Empty, _) => 1
          case (_, SortValue.Empty) => -1
          case (SortValue.Error, SortValue.Error) => 0
          case (SortValue.Error, _) => 1
          case (_, SortValue.Error) => -1
          case (SortValue.Num(na), SortValue.Num(nb)) => na.compare(nb)
          case (SortValue.Str(sa), SortValue.Str(sb)) => sa.compare(sb)
          case (SortValue.Num(n), SortValue.Str(s)) =>
            if mode == Edit.SortMode.Numeric then -1 else n.toString.compare(s)
          case (SortValue.Str(s), SortValue.Num(n)) =>
            if mode == Edit.SortMode.Numeric then 1 else s.compare(n.toString)

  /** An absent key cell or an empty one: Excel excludes blanks from the ordering (GH-596). */
  private def isBlankKey(v: Option[SortValue]): Boolean =
    v.forall(_ == SortValue.Empty)

  private def rowComparator(
    keys: Vector[Edit.SortKeySpec]
  ): (Map[Int, Cell], Map[Int, Cell]) => Int =
    (rowA, rowB) =>
      keys.iterator
        .map { key =>
          val colIdx = key.column.index0
          val valueA = rowA.get(colIdx).map(c => sortableValue(c.value, key.mode))
          val valueB = rowB.get(colIdx).map(c => sortableValue(c.value, key.mode))
          val cmp = compareSortValues(valueA, valueB, key.mode)
          // GH-596: blanks sort LAST in both directions — Excel appends them after the ordered
          // values rather than treating them as the minimum, so the direction flips only the
          // comparison between two non-blank keys.
          if key.direction == Edit.SortDir.Descending && !isBlankKey(valueA) && !isBlankKey(valueB)
          then -cmp
          else cmp
        }
        .find(_ != 0)
        .getOrElse(0)

  /** Every key column must lie inside the sorted range. */
  def validateSortKeys(range: CellRange, keys: Vector[Edit.SortKeySpec]): XLResult[Unit] =
    if keys.isEmpty then Left(XLError.InvalidReference("sort requires at least one key column"))
    else
      keys
        .find(k => k.column.index0 < range.colStart.index0 || k.column.index0 > range.colEnd.index0)
        .map(k =>
          XLError.InvalidReference(
            s"Sort column ${k.column.toLetter} is outside range ${range.toA1}"
          )
        )
        .toLeft(())

  /**
   * Sort the rows of `range` by `keys` (a stable sort, so equal keys keep their order). Only cells
   * within the range's columns move; styles and comments move with their rows; a formula that moves
   * has its relative row references shifted by the displacement and its cache dropped (a
   * single-cell array formula stays one at its new row).
   */
  def sort(
    sheet: Sheet,
    range: CellRange,
    keys: Vector[Edit.SortKeySpec],
    hasHeader: Boolean
  )(using FormulaSupport): XLResult[Sheet] =
    validateSortKeys(range, keys).flatMap { _ =>
      val rowStart = range.rowStart.index0
      val rowEnd = range.rowEnd.index0
      val colStart = range.colStart.index0
      val colEnd = range.colEnd.index0
      val dataRowStart = if hasHeader then rowStart + 1 else rowStart
      val dataRange = Option.when(dataRowStart <= rowEnd)(
        CellRange(ARef.from0(colStart, dataRowStart), ARef.from0(colEnd, rowEnd))
      )
      val blocked = sheet.cells.iterator
        .collect {
          case cell @ Cell(_, CellValue.Formula(_, _, arr: FormulaKind.ArrayFormula), _, _)
              if arr.isDynamicArray && arr.ref.start == cell.ref && arr.ref.cellCount > 1 &&
                dataRange.exists(_.intersects(arr.ref)) =>
            cell.ref -> arr.ref
        }
        .toVector
        .sortBy((anchor, _) => (anchor.row.index0, anchor.col.index0))
        .headOption
      if dataRowStart > rowEnd then Right(sheet)
      else
        blocked match
          case Some((anchor, extent)) =>
            Left(
              XLError.InvalidReference(
                s"cannot sort through the dynamic array at ${anchor.toA1} (${extent.toA1}); " +
                  "clear or overwrite its anchor first, or sort outside its spill range"
              )
            )
          case None =>
            val rowsToSort: Vector[(Int, Map[Int, Cell])] =
              (dataRowStart to rowEnd).toVector.map { rowIdx =>
                val cellMap = (colStart to colEnd).flatMap { colIdx =>
                  sheet.cells.get(ARef.from0(colIdx, rowIdx)).map(colIdx -> _)
                }.toMap
                (rowIdx, cellMap)
              }
            val comparator = rowComparator(keys)
            val sortedRows = rowsToSort.sortWith((a, b) => comparator(a._2, b._2) < 0)
            reconstructWithSortedRows(sheet, range, dataRowStart, sortedRows)
    }

  private def reconstructWithSortedRows(
    sheet: Sheet,
    range: CellRange,
    dataRowStart: Int,
    sortedRows: Vector[(Int, Map[Int, Cell])]
  )(using FormulaSupport): XLResult[Sheet] =
    val rowEnd = range.rowEnd.index0
    val colStart = range.colStart.index0
    val colEnd = range.colEnd.index0

    val withoutDataCells = (dataRowStart to rowEnd).foldLeft(sheet) { (s, rowIdx) =>
      (colStart to colEnd).foldLeft(s)((s2, colIdx) => s2.remove(ARef.from0(colIdx, rowIdx)))
    }

    val refMapping: Map[ARef, ARef] = sortedRows.zipWithIndex.flatMap {
      case ((originalRowIdx, cellMap), newRowOffset) =>
        val newRowIdx = dataRowStart + newRowOffset
        cellMap.keys.map(colIdx =>
          ARef.from0(colIdx, originalRowIdx) -> ARef.from0(colIdx, newRowIdx)
        )
    }.toMap

    val placed = foldE(sortedRows.zipWithIndex.iterator)(withoutDataCells) {
      case (s, ((originalRowIdx, cellMap), newRowOffset)) =>
        val newRowIdx = dataRowStart + newRowOffset
        val rowDelta = newRowIdx - originalRowIdx
        // Columns in ascending order so the result is independent of Map iteration order
        foldE(cellMap.toVector.sortBy(_._1).iterator)(s) { case (s2, (_, cell)) =>
          val newRef = ARef.from0(cell.col.index0, newRowIdx)
          val moved: XLResult[CellValue] =
            if rowDelta == 0 then
              cell.value match
                case CellValue.Formula(_, cachedOpt, _: FormulaKind.DataTable) =>
                  Right(cachedOpt.getOrElse(CellValue.Empty))
                case v => Right(v)
            else shifted(cell.value, newRef, 0, rowDelta).map(_._1)
          moved.map(v => s2.put(Cell(newRef, v, cell.styleId, cell.hyperlink)))
        }
    }

    placed.map { s =>
      val remappedComments =
        s.comments.map((ref, comment) => refMapping.getOrElse(ref, ref) -> comment)
      s.copy(comments = remappedComments)
    }

  // ===== Clear =====

  /**
   * Clear contents, styles and/or comments in `range`.
   *
   * GH-890: a contents-only clear is Excel's Clear Contents: each emptied cell keeps its style but
   * loses its hyperlink, a cell with no style is removed (unless its row or column has a default
   * style, GH-897), and merges stay. The CLI's `clear` verb (`DynamicArrayAuthoring.clearContents`,
   * GH-886) applies the same rule through [[keepsRecordWhenCleared]] (GH-908). A clear that drops
   * the cells as well (`ClearWhat.all`) removes the cell records.
   *
   * GH-909: any clear with `styles` is Excel's Clear Formats (or Clear All): it unmerges every
   * merged region the range touches and resets the styles of each whole merge, since Excel widens
   * the selection to the merges it touches. A cell of a merge that lies outside the range keeps its
   * value but loses its style. The CLI's `clear --styles` / `--all` (GH-894) make the same change.
   */
  def clear(sheet: Sheet, range: CellRange, what: ClearWhat): Sheet =
    val dropCells = what.contents && what.styles
    val touched =
      if what.styles then sheet.mergedRanges.filter(_.intersects(range)).toVector
      else Vector.empty
    val s1 =
      if dropCells then sheet.removeRange(range)
      else if what.contents then clearContentsKeepingStyles(sheet, range)
      else sheet
    val s2 =
      if !what.styles then s1
      else
        val inRange = if dropCells then s1 else s1.clearStylesInRange(range)
        crossingMerges(range, touched).foldLeft(inRange)((s, m) => s.clearStylesInRange(m))
    val s3 = if what.comments then s2.clearCommentsInRange(range) else s2
    // #931: contents clear takes the hyperlinks with it, range hyperlinks' coverage included
    val s4 = if what.contents then s3.removeRangeHyperlinkCoverage(range) else s3
    touched.foldLeft(s4)((s, m) => s.unmerge(m))

  // ===== Merge (GH-935) =====

  /**
   * Merge `range` without a formula parser (GH-935). Excel's Merge Cells (checked against Excel for
   * Mac) keeps a non-empty upper-left cell and clears the rest; when the upper-left cell is empty
   * it MOVES the first non-empty cell of the range (row-major) into it and points every reader of
   * that cell at the upper-left cell. xl-core cannot rewrite those readers, and moving the cell
   * without them would silently change their values (`=B1+1` would read the emptied B1). So here:
   *   - a non-empty upper-left cell keeps its value; every other cell is cleared (as Excel);
   *   - an empty upper-left cell stays empty and the first non-empty cell ([[mergeWinner]]) stays
   *     where it is, untouched; every OTHER cell is cleared. Every reader keeps Excel's value (the
   *     kept cell holds what Excel would move); only the merged cell's display differs from Excel.
   *
   * For Excel's move with reader rewriting use the workbook-level paths: `wb.edit(Edit.Merge(...))`
   * (through `FormulaSupport.relocateCell`) or xl-evaluator's `CellMover.merge`.
   *
   * A cleared cell is cleared as Clear Contents clears it ([[clearContentsKeepingStyles]]): its
   * value or formula and its hyperlink go, a cell that keeps a record keeps its own style, an
   * unstyled cell is removed. Comments stay (Clear Contents keeps them too). The store is walked by
   * its stored cells, so a whole-column merge does not visit a million positions.
   */
  def merge(sheet: Sheet, range: CellRange): Sheet =
    if mergeRefusal(sheet, range).isDefined || mergeIsNoOp(sheet, range) then sheet
    else mergeUnchecked(sheet, range)

  /**
   * GH-935 (Excel oracle): merging inside an existing merged area, or the same range again, or a
   * single cell, changes nothing — the final area ([[mergeTarget]]) is that merge, and Excel leaves
   * its cells, hidden values included, exactly as they were.
   */
  private[xl] def mergeIsNoOp(sheet: Sheet, range: CellRange): Boolean =
    // GH-935 review 4: merging a single cell does nothing in Excel; never a one-cell <mergeCell>
    val target = mergeTarget(sheet, range)
    target.start == target.end || sheet.mergedRanges.contains(target)

  /** [[merge]] once [[mergeRefusal]] has passed. */
  private def mergeUnchecked(sheet: Sheet, range: CellRange): Sheet =
    val target = mergeTarget(sheet, range)
    // Excel widens a selection to the merged areas it touches and merges the whole: the absorbed
    // merges go, so the sheet never holds overlapping merges (Excel repairs a file that does)
    val base = sheet.copy(mergedRanges = sheet.mergedRanges.filterNot(_.intersects(target)))
    // the winner is judged on the sheet as it was (a hidden cell of an absorbed merge never wins)
    val kept = mergeWinner(sheet, range).map(w => w.copy(hyperlink = base.hyperlinkAt(w.ref)))
    // a hidden cell loses its links as it loses its value: its own and range-link coverage
    val cleared = hiddenByMerge(target).foldLeft(base)((s, r) =>
      clearContentsKeepingStyles(s, r).removeRangeHyperlinkCoverage(r)
    )
    val restored = kept.fold(cleared)(cleared.put)
    restored.copy(mergedRanges = restored.mergedRanges + target)

  /**
   * The range a merge of `range` actually merges (GH-935 review): `range` grown, until stable, to
   * the bounding box of every merged area it intersects, as Excel extends a selection over the
   * merged cells it touches. The existing merges inside the result are absorbed by the merge.
   */
  private[xl] def mergeTarget(sheet: Sheet, range: CellRange): CellRange =
    @annotation.tailrec
    def grow(r: CellRange): CellRange =
      val grown = sheet.mergedRanges.filter(_.intersects(r)).foldLeft(r) { (acc, m) =>
        CellRange(
          ARef.from0(
            acc.colStart.index0.min(m.colStart.index0),
            acc.rowStart.index0.min(m.rowStart.index0)
          ),
          ARef.from0(acc.colEnd.index0.max(m.colEnd.index0), acc.rowEnd.index0.max(m.rowEnd.index0))
        )
      }
      if grown == r then r else grow(grown)
    grow(range)

  /**
   * The cell Excel moves into the upper-left cell of a merge of `range` (GH-935), judged over the
   * range the merge actually covers ([[mergeTarget]]): `None` when the upper-left cell holds a
   * value (it wins) or nothing in the range does; otherwise the first non-empty cell in row-major
   * order. A spill child is never merge content (GH-935 review): a cell inside an array formula's
   * result range other than its anchor (a dynamic spill's cached value, a legacy array member) is
   * skipped, as Excel leaves the spill to its anchor (which then shows #SPILL!). [[merge]] keeps
   * the winner in place; the workbook-level merges move it ([[moveCell]]) after rewriting its
   * readers.
   */
  private[xl] def mergeWinner(sheet: Sheet, range: CellRange): Option[Cell] =
    val target = mergeTarget(sheet, range)
    val anchor = target.start
    val arrays = sheet.cells.iterator.flatMap { c =>
      c.value match
        case CellValue.Formula(_, _, k: FormulaKind.ArrayFormula) if k.ref.intersects(target) =>
          Some((c.ref, k.ref))
        case _ => None
    }.toVector
    def spillChild(ref: ARef): Boolean = arrays.exists((a, r) => a != ref && r.contains(ref))
    // GH-935 (Excel oracle): a hidden cell of a merged area the merge absorbs is never a candidate
    // (its value is cleared, nothing that reads it is rewritten)
    val absorbed = sheet.mergedRanges.filter(_.intersects(target)).toVector
    def hiddenBefore(ref: ARef): Boolean = absorbed.exists(m => m.contains(ref) && m.start != ref)
    // GH-935 review 3: a spilled value is never merge content, the upper-left cell's included:
    // the upper-left cell wins only when it holds real content of its own
    def content(c: Cell): Boolean =
      c.value != CellValue.Empty && !spillChild(c.ref) && !hiddenBefore(c.ref)
    if sheet.cells.get(anchor).exists(content) then None
    else
      sheet.cells
        .inRange(target)
        .filter(c => c.ref != anchor && content(c))
        .minByOption(c => (c.ref.row.index0, c.ref.col.index0))

  /**
   * Why Excel refuses to merge `range` (GH-935 review), judged over the range the merge actually
   * covers ([[mergeTarget]]): it intersects a legacy (CSE) array formula's range or a data table's
   * range — Excel cannot change part of an array or a data table, and array formulas are not valid
   * in merged cells. `None` when the merge may proceed. A dynamic array is not a refusal: merging
   * over its spill range blocks it (`#SPILL!`), as in Excel; merging over its anchor discards the
   * anchor like any other hidden content, or keeps it as the merged cell's value.
   */
  private[xl] def mergeRefusal(sheet: Sheet, range: CellRange): Option[XLError] =
    val target = mergeTarget(sheet, range)
    sheet.cells.iterator
      .flatMap { c =>
        c.value match
          case CellValue.Formula(_, _, k: FormulaKind.ArrayFormula)
              if !k.isDynamicArray && k.ref.intersects(target) =>
            Some(("an array formula", k.ref))
          case CellValue.Formula(_, _, k: FormulaKind.DataTable) if k.ref.intersects(target) =>
            Some(("a data table", k.ref))
          case _ => None
      }
      .toVector
      .sortBy((_, r) => (r.start.row.index0, r.start.col.index0))
      .headOption
      .map { (what, r) =>
        XLError.InvalidRange(
          target.toA1,
          s"the merge range ${target.toA1} intersects $what at ${r.toA1}; Excel cannot merge " +
            s"cells of $what"
        )
      }

  /**
   * Whether the merge of `range` touches a dynamic array's spill (its anchor's declared range
   * intersects what the merge covers): the workbook-level merge releases such a spill, which the
   * kernel cannot, so the Edit/Patch lowerings leave these merges alone.
   */
  private[xl] def mergeTouchesSpill(sheet: Sheet, range: CellRange): Boolean =
    val target = mergeTarget(sheet, range)
    sheet.cells.iterator.exists { c =>
      c.value match
        case CellValue.Formula(_, _, k: FormulaKind.ArrayFormula) =>
          k.isDynamicArray && k.ref.cellCount > 1 && k.ref.intersects(target)
        case _ => false
    }

  /** Whether `Edit.Merge`/`Patch.Merge` of `range` lower exactly to `Sheet.merge` (no move). */
  private[xl] def mergeLowers(sheet: Sheet, range: CellRange): Boolean =
    mergeRefusal(sheet, range).isEmpty && mergeWinner(sheet, range).isEmpty &&
      !mergeTouchesSpill(sheet, range)

  /** The upper-left cell of the range a merge of `range` covers ([[mergeTarget]]). */
  private[xl] def mergeAnchor(sheet: Sheet, range: CellRange): ARef =
    mergeTarget(sheet, range).start

  /**
   * Cut the cell at `from` and paste it at `to` on the same sheet: the record (value or formula
   * verbatim, style, effective hyperlink) lands at `to`, replacing what was there, and `from` is
   * left with no record and no link. References to `from` are not rewritten here (no parser in
   * xl-core). No record at `from` leaves the sheet unchanged.
   */
  private[xl] def moveCell(sheet: Sheet, from: ARef, to: ARef): Sheet =
    if from == to then sheet
    else
      sheet.cells.get(from).fold(sheet) { cell =>
        // the cut carries the cell's EFFECTIVE link (its own, else the range link covering it) and
        // replaces whatever linked the destination; the source is left unlinked
        val link = sheet.hyperlinkAt(from)
        sheet
          .remove(from)
          .removeRangeHyperlinkCoverage(CellRange(from, from))
          .removeRangeHyperlinkCoverage(CellRange(to, to))
          .put(cell.copy(ref = to, value = movedGeometry(cell.value, from, to), hyperlink = link))
      }

  /**
   * A moved record's own geometry follows it: an array formula's (legacy or dynamic) result range
   * and a data table's range, when anchored at the moved cell, shift by the move's displacement, as
   * a cut and paste carries them. Inputs a data table names (`r1`/`r2`) are other cells and stay.
   */
  private def movedGeometry(value: CellValue, from: ARef, to: ARef): CellValue =
    val dc = to.col.index0 - from.col.index0
    val dr = to.row.index0 - from.row.index0
    // Only a range the record anchors (its top-left is the moved cell, as OOXML writes array and
    // data-table masters) moves with it; the move is then up-left inside the merge, never off-grid
    def shifted(r: CellRange): CellRange =
      def at(a: ARef): ARef = ARef.from0(a.col.index0 + dc, a.row.index0 + dr)
      if r.start == from then r.copy(start = at(r.start), end = at(r.end)) else r
    value match
      case f @ CellValue.Formula(_, _, k: FormulaKind.ArrayFormula) =>
        f.copy(kind = k.copy(ref = shifted(k.ref)))
      case f @ CellValue.Formula(_, _, k: FormulaKind.DataTable) =>
        f.copy(kind = k.copy(ref = shifted(k.ref)))
      case other => other

  /**
   * The cells a merge of `range` hides — everything but the upper-left cell — as at most two
   * rectangles: the rest of the first row, then every row below it.
   */
  private[xl] def hiddenByMerge(range: CellRange): Vector[CellRange] =
    val start = range.start
    val end = range.end
    val restOfRow =
      Option.when(end.col.index0 > start.col.index0)(
        CellRange(
          ARef.from0(start.col.index0 + 1, start.row.index0),
          ARef.from0(end.col.index0, start.row.index0)
        )
      )
    val below =
      Option.when(end.row.index0 > start.row.index0)(
        CellRange(ARef.from0(start.col.index0, start.row.index0 + 1), end)
      )
    (restOfRow.toList ++ below.toList).toVector

  /**
   * The merges in `touched` that cross the edge of `range`: their cells outside it need their own
   * style reset (GH-909). A merge wholly inside `range` is already covered by the range's own.
   */
  private[xl] def crossingMerges(range: CellRange, touched: Vector[CellRange]): Vector[CellRange] =
    touched.filterNot(m => range.contains(m.start) && range.contains(m.end))

  /**
   * Clear Contents: a cell in `range` that keeps a record is emptied and loses its hyperlink; any
   * other is removed. A cell keeps its record when it has a style, or when its row or column has a
   * default style ([[keepsRecordWhenCleared]]).
   */
  private def clearContentsKeepingStyles(sheet: Sheet, range: CellRange): Sheet =
    // the store walks the range and merges the kept records back once (#891)
    val kept = sheet.cells
      .inRange(range)
      .filter(keepsRecordWhenCleared(sheet, _))
      .map(_.copy(value = CellValue.Empty, hyperlink = None))
      .toVector
    sheet.copy(cells = sheet.cells.removedIn(range).putAll(kept))

  /**
   * Whether a contents clear keeps `cell` as an empty record: it has a style, or its row or column
   * has a default style (GH-897), which a cell with no record would take instead of its own.
   */
  private[xl] def keepsRecordWhenCleared(sheet: Sheet, cell: Cell): Boolean =
    cell.styleId.isDefined ||
      sheet.rowProperties.get(cell.ref.row).exists(_.styleId.isDefined) ||
      sheet.columnProperties.get(cell.ref.col).exists(_.styleId.isDefined)

  // ===== Column auto-fit (GH-156) =====

  /** Excel's default column width in character units, used for empty columns. */
  val DefaultColumnWidth: Double = 8.43

  /**
   * Lower bound on an auto-fit width (Excel allows narrower columns; the CLI never produced one).
   */
  private val MinAutoFitWidth: Double = 5.0

  /** Max digit width of Calibri 11 at 96 DPI (the OOXML/POI column-width convention). */
  private val MaxDigitWidthPx: Double = 7.0

  /** Cell padding in px: 2 left margin + 2 right margin + 1 gridline (OOXML/POI convention). */
  private val CellPaddingPx: Double = 5.0

  /**
   * The width a column needs for its formatted content, in Excel character units. Each cell's
   * display text is measured with its resolved font (`RenderUtils.measureTextWidth`, which
   * estimates from character counts when no graphics context exists) and converted with
   * `width = (textPx + 5) / 7` — the Calibri-11 max-digit-width convention that matches Excel's own
   * auto-fit (`RenderUtils.excelColWidthToPixels` deliberately uses a wider factor for SVG
   * fidelity). If measurement throws (no fontconfig), the pre-0.12 `chars × 0.90 + 1.5` heuristic
   * is used for that cell. A formula cell is measured by its cached value; an UNCACHED formula
   * contributes nothing (GH-613) — evaluate or recalculate first when the fit must reflect it.
   */
  def autoFitWidth(sheet: Sheet, col: Column): Double =
    fitWidth(sheet, sheet.cells.iterator.filter(_.ref.col == col).toVector)

  /**
   * [[autoFitWidth]] for every column of `columns`, in order, from ONE pass over the sheet's cells:
   * the cells are grouped by column once and each width is computed from its group, so fitting
   * every column of a wide sheet is O(cells + columns) rather than O(cells x columns). A column
   * without cells gets [[DefaultColumnWidth]]; a repeated column is reported each time.
   */
  def autoFitWidths(sheet: Sheet, columns: Iterable[Column]): Vector[(Column, Double)] =
    val byColumn: Map[Column, Vector[Cell]] = sheet.cells.iterator.toVector.groupBy(_.ref.col)
    columns.iterator
      .map(col => col -> fitWidth(sheet, byColumn.getOrElse(col, Vector.empty)))
      .toVector

  /** The width the cells of one column need; [[DefaultColumnWidth]] when there are none. */
  private def fitWidth(sheet: Sheet, cellsInColumn: Vector[Cell]): Double =
    if cellsInColumn.isEmpty then DefaultColumnWidth
    else
      val maxWidth = cellsInColumn.map(cellWidth(_, sheet)).maxOption.getOrElse(0.0)
      val rounded = BigDecimal(maxWidth).setScale(2, BigDecimal.RoundingMode.HALF_UP).toDouble
      math.max(rounded, MinAutoFitWidth)

  private def cellWidth(cell: Cell, sheet: Sheet): Double =
    val styleOpt = cell.styleId.flatMap(sheet.styleRegistry.get)
    val numFmt = styleOpt.map(_.numFmt).getOrElse(NumFmt.General)
    // GH-613: a column is fitted to what it DISPLAYS. A formula with a cache displays that value;
    // one without has nothing to measure yet — its text (`=COUNTIF('Deal Pipeline'!$E$2:$E$1000,A2)`)
    // is not what Excel shows, so it contributes nothing rather than a 44-character width.
    val text = cell.value match
      case CellValue.Formula(_, None, _) => ""
      case value => RenderUtils.cellValueToText(value, numFmt)
    if text.isEmpty then 0.0
    else
      Try(RenderUtils.measureTextWidth(text, styleOpt.map(_.font))).fold(
        _ => heuristicWidth(text, styleOpt),
        px => (px + CellPaddingPx) / MaxDigitWidthPx
      )

  private def heuristicWidth(text: String, styleOpt: Option[CellStyle]): Double =
    val boldFactor = if styleOpt.exists(_.font.bold) then 1.1 else 1.0
    text.length * boldFactor * 0.90 + 1.5

  /**
   * Set every column's width to its [[autoFitWidth]], measured in one pass ([[autoFitWidths]]).
   * Widths are read from the input sheet: setting a column's properties changes neither the cells
   * nor the style registry, so this equals the per-column fold cell for cell.
   */
  def autoFit(sheet: Sheet, columns: Iterable[Column]): Sheet =
    autoFitWidths(sheet, columns).foldLeft(sheet) { case (s, (col, width)) =>
      s.setColumnProperties(col, s.getColumnProperties(col).copy(width = Some(width)))
    }

  // ===== Row/column outline grouping (GH-421) =====

  def validateLevel(level: Int): XLResult[Unit] =
    if level >= 1 && level <= 7 then Right(())
    else Left(XLError.InvalidArgument("outline", s"level must be 1-7, got: $level"))

  /** Excel's ceiling on a column width, in character units. */
  val MaxColumnWidth: Double = 255.0

  /** Excel's ceiling on a row height, in points. */
  val MaxRowHeight: Double = 409.0

  /**
   * The column-width guard (GH-617), `INVALID_ARGUMENT` naming `op` — shared by `Edit.validate`
   * (`col-width`) and the CLI's `col --width` / batch `colwidth`, so no path writes a width Excel
   * repairs.
   */
  def validateColumnWidth(op: String, width: Double): XLResult[Unit] =
    if width >= 0 && width <= MaxColumnWidth then Right(())
    else
      Left(
        XLError.InvalidArgument(
          op,
          s"width must be 0-${MaxColumnWidth.toInt} character units, got $width"
        )
      )

  /** The row-height guard (GH-617), the twin of [[validateColumnWidth]]. */
  def validateRowHeight(op: String, height: Double): XLResult[Unit] =
    if height >= 0 && height <= MaxRowHeight then Right(())
    else
      Left(
        XLError.InvalidArgument(op, s"height must be 0-${MaxRowHeight.toInt} points, got $height")
      )

  /**
   * Apply `f` to the row's properties, keeping the entry even when the result is all-default: an
   * explicit entry is authoritative on write (actively clearing preserved source attributes),
   * whereas a missing entry lets a preserved `<row>` ride through verbatim — pruning would
   * resurrect the very collapsed/outlineLevel attrs an ungroup just cleared. A row without an entry
   * gains one, as `Edit.lower`'s `SetRowProperties` per member does (the lowering-coherence law).
   */
  private def updateRow(sheet: Sheet, row: Row)(f: RowProperties => RowProperties): Sheet =
    sheet.setRowProperties(row, f(sheet.getRowProperties(row)))

  private def updateCol(sheet: Sheet, col: Column)(f: ColumnProperties => ColumnProperties): Sheet =
    sheet.setColumnProperties(col, f(sheet.getColumnProperties(col)))

  /**
   * [[updateRow]] for an edit that is a no-op wherever nothing was set (GH-589 expand): an existing
   * entry is updated and kept as above, but a row WITHOUT an entry that `f` leaves at the default
   * stays absent — nothing changed, so nothing is emitted. Expanding rows that were never hidden
   * leaves the file byte-identical instead of adding an empty `<row>` per member.
   */
  private def touchRow(sheet: Sheet, row: Row)(f: RowProperties => RowProperties): Sheet =
    val current = sheet.getRowProperties(row)
    val next = f(current)
    if next == current && !sheet.rowProperties.contains(row) then sheet
    else sheet.setRowProperties(row, next)

  private def touchCol(sheet: Sheet, col: Column)(f: ColumnProperties => ColumnProperties): Sheet =
    val current = sheet.getColumnProperties(col)
    val next = f(current)
    if next == current && !sheet.columnProperties.contains(col) then sheet
    else sheet.setColumnProperties(col, next)

  /**
   * The one outline composition (GH-421 grouping, GH-465/GH-589 collapse and expand): `f` on every
   * row of the span, then the summary row AFTER it — Excel's `summaryBelow` default — gets
   * `collapsed = marker` when a marker is given (a span ending on the last row has no summary row).
   * [[groupRows]]/[[ungroupRows]] here and `Sheet.collapseRows`/`expandRows` (`outlineSyntax`) are
   * all this fold, so the summary-marker rule cannot drift between the CLI and scripts. `sparse`
   * selects [[touchRow]] over [[updateRow]]: rows without an entry stay absent when `f` would leave
   * them at the default (expand); grouping keeps [[updateRow]] to stay coherent with `Edit.lower`.
   */
  private[xl] def outlineRows(
    sheet: Sheet,
    rows: RowSpan,
    marker: Option[Boolean],
    sparse: Boolean
  )(
    f: RowProperties => RowProperties
  ): Sheet =
    def step(s: Sheet, r: Row)(g: RowProperties => RowProperties): Sheet =
      if sparse then touchRow(s, r)(g) else updateRow(s, r)(g)
    val members = rows.rows.foldLeft(sheet)((s, r) => step(s, r)(f))
    marker match
      case Some(collapsed) if rows.end.index0 < Row.MaxIndex0 =>
        step(members, rows.end + 1)(_.copy(collapsed = collapsed))
      case _ => members

  /** The column twin of [[outlineRows]] (`summaryRight`). */
  private[xl] def outlineCols(
    sheet: Sheet,
    cols: ColSpan,
    marker: Option[Boolean],
    sparse: Boolean
  )(
    f: ColumnProperties => ColumnProperties
  ): Sheet =
    def step(s: Sheet, c: Column)(g: ColumnProperties => ColumnProperties): Sheet =
      if sparse then touchCol(s, c)(g) else updateCol(s, c)(g)
    val members = cols.columns.foldLeft(sheet)((s, c) => step(s, c)(f))
    marker match
      case Some(collapsed) if cols.end.index0 < Column.MaxIndex0 =>
        step(members, cols.end + 1)(_.copy(collapsed = collapsed))
      case _ => members

  /**
   * Group rows into a collapsible outline: every row gets `level`; `collapsed` hides the members
   * and marks the summary row AFTER the group (Excel's summaryBelow default) so the "+" button
   * draws there.
   */
  def groupRows(sheet: Sheet, rows: RowSpan, level: Int, collapsed: Boolean): XLResult[Sheet] =
    validateLevel(level).map { _ =>
      outlineRows(sheet, rows, Option.when(collapsed)(true), sparse = false) { p =>
        p.copy(outlineLevel = Some(level), hidden = collapsed || p.hidden)
      }
    }

  def groupCols(sheet: Sheet, cols: ColSpan, level: Int, collapsed: Boolean): XLResult[Sheet] =
    validateLevel(level).map { _ =>
      outlineCols(sheet, cols, Option.when(collapsed)(true), sparse = false) { p =>
        p.copy(outlineLevel = Some(level), hidden = collapsed || p.hidden)
      }
    }

  /**
   * Clear the outline level and collapse markers of the rows and of the group's summary row.
   * Members a collapse hid stay hidden, like Excel (show them with a row-show edit).
   */
  def ungroupRows(sheet: Sheet, rows: RowSpan): Sheet =
    outlineRows(sheet, rows, Some(false), sparse = false)(
      _.copy(outlineLevel = None, collapsed = false)
    )

  def ungroupCols(sheet: Sheet, cols: ColSpan): Sheet =
    outlineCols(sheet, cols, Some(false), sparse = false)(
      _.copy(outlineLevel = None, collapsed = false)
    )

  // ===== Appearance & print setup (GH-358) =====

  /** Excel accepts a zoom of 10-400 percent. */
  def validateZoom(zoom: Option[Int]): XLResult[Unit] =
    zoom.filter(z => z < 10 || z > 400) match
      case Some(z) =>
        Left(XLError.InvalidArgument("sheet-view", s"zoom scale must be 10-400, got: $z"))
      case None => Right(())

  /** Merge view options into the sheet's current view; unspecified fields are preserved. */
  def mergeSheetView(
    sheet: Sheet,
    gridlines: Option[Boolean],
    zoom: Option[Int],
    tabSelected: Option[Boolean]
  ): XLResult[Sheet] =
    validateZoom(zoom).map { _ =>
      val current = sheet.viewSettings.getOrElse(SheetView.default)
      sheet.withViewSettings(
        current.copy(
          showGridLines = gridlines.getOrElse(current.showGridLines),
          zoomScale = zoom.orElse(current.zoomScale),
          tabSelected = tabSelected.orElse(current.tabSelected)
        )
      )
    }

  /** The page-setup guards, checked ahead of the model's `require`s. */
  def validatePageSetup(
    orientation: Option[String],
    scale: Option[Int],
    fitToWidth: Option[Int],
    fitToHeight: Option[Int]
  ): XLResult[Unit] =
    for
      _ <- orientation
        .filterNot(o => o == "portrait" || o == "landscape")
        .map(o => XLError.Other(s"Orientation must be 'portrait' or 'landscape', got: $o"))
        .toLeft(())
      _ <- scale
        .filter(sc => sc < 10 || sc > 400)
        .map(sc => XLError.Other(s"Scale must be 10-400, got: $sc"))
        .toLeft(())
      // GH-463: 0 is Excel's "automatic" (as many pages as needed on that axis)
      _ <- fitToWidth
        .filter(_ < 0)
        .map(n => XLError.Other(s"fit-to-width must be >= 0 (0 = automatic), got: $n"))
        .toLeft(())
      _ <- fitToHeight
        .filter(_ < 0)
        .map(n => XLError.Other(s"fit-to-height must be >= 0 (0 = automatic), got: $n"))
        .toLeft(())
    yield ()

  /** Merge print options into the sheet's page setup; unspecified fields are preserved. */
  def mergePageSetup(
    sheet: Sheet,
    orientation: Option[String],
    scale: Option[Int],
    fitToWidth: Option[Int],
    fitToHeight: Option[Int],
    fitToPage: Option[Boolean]
  ): XLResult[Sheet] =
    validatePageSetup(orientation, scale, fitToWidth, fitToHeight).map { _ =>
      val current = sheet.pageSetup.getOrElse(PageSetup.default)
      sheet.withPageSetup(
        current.copy(
          scale = scale.getOrElse(current.scale),
          orientation = orientation.orElse(current.orientation),
          fitToWidth = fitToWidth.orElse(current.fitToWidth),
          fitToHeight = fitToHeight.orElse(current.fitToHeight),
          // Tri-state (GH-284): None derives from fitToWidth/fitToHeight and preserves any flag
          // in the source file; Some(false) actively strips a preserved flag.
          fitToPage = fitToPage.orElse(current.fitToPage)
        )
      )
    }

  /**
   * Merge header/footer text into the page setup. Even-page text sets `differentOddEven` and
   * first-page text sets `differentFirst` (Excel ignores the text while the flag is off); the
   * explicit flags force them on without text.
   */
  def mergeHeaderFooter(
    sheet: Sheet,
    oddHeader: Option[String],
    oddFooter: Option[String],
    evenHeader: Option[String],
    evenFooter: Option[String],
    firstHeader: Option[String],
    firstFooter: Option[String],
    differentOddEven: Boolean,
    differentFirst: Boolean
  ): Sheet =
    val setup = sheet.pageSetup.getOrElse(PageSetup.default)
    val hf = setup.headerFooter.getOrElse(HeaderFooter())
    val updated = hf.copy(
      oddHeader = oddHeader.orElse(hf.oddHeader),
      oddFooter = oddFooter.orElse(hf.oddFooter),
      evenHeader = evenHeader.orElse(hf.evenHeader),
      evenFooter = evenFooter.orElse(hf.evenFooter),
      firstHeader = firstHeader.orElse(hf.firstHeader),
      firstFooter = firstFooter.orElse(hf.firstFooter),
      differentOddEven =
        hf.differentOddEven || differentOddEven || evenHeader.isDefined || evenFooter.isDefined,
      differentFirst =
        hf.differentFirst || differentFirst || firstHeader.isDefined || firstFooter.isDefined
    )
    sheet.withPageSetup(setup.copy(headerFooter = Some(updated)))
