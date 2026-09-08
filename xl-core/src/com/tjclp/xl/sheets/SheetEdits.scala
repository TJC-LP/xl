package com.tjclp.xl.sheets

import scala.util.Try

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue, FormulaKind}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.ops.{ClearWhat, ColSpan, Edit, FormulaSupport, RowSpan}
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
   * The value a cell takes when it moves by `(colDelta, rowDelta)`. A data-table record always
   * copies as its cached constant (GH-430: pasting the TABLE(...) display text would be a `#NAME?`
   * bomb, and the record would claim a table interior that does not exist at the target). A formula
   * is shifted; its cache is dropped because it no longer describes the shifted text.
   */
  private def shifted(
    value: CellValue,
    colDelta: Int,
    rowDelta: Int
  )(using fs: FormulaSupport): XLResult[CellValue] =
    value match
      case CellValue.Formula(_, cachedOpt, _: FormulaKind.DataTable) =>
        Right(cachedOpt.getOrElse(CellValue.Empty))
      case CellValue.Formula(expr, _, _) =>
        fs.shift(expr, colDelta, rowDelta) match
          case Right(text) => Right(CellValue.Formula(text, None))
          case Left(refusal: XLError.UnsupportedCapability) => Left(refusal)
          case Left(_) => Right(CellValue.Formula(expr, None)) // unparseable: copied as written
      case other => Right(other)

  /** `foldLeft` that stops at the first `Left`. */
  private def foldE[A, B](as: Iterator[A])(zero: B)(f: (B, A) => XLResult[B]): XLResult[B] =
    as.foldLeft[XLResult[B]](Right(zero)) {
      case (Right(acc), a) => f(acc, a)
      case (left, _) => left
    }

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
    validateFill(source, target, direction).flatMap { _ =>
      direction match
        case Edit.FillDir.Down =>
          val sourceStartRow = source.rowStart.index0
          val sourceRowCount = source.rowEnd.index0 - sourceStartRow + 1
          val targetStartRow = target.rowStart.index0
          val cols = source.colStart.index0 to source.colEnd.index0
          foldE((targetStartRow to target.rowEnd.index0).iterator)(sheet) { (s, targetRow) =>
            val sourceRow = sourceStartRow + (targetRow - targetStartRow) % sourceRowCount
            val rowDelta = targetRow - sourceRow
            if rowDelta == 0 then Right(s)
            else
              foldE(cols.iterator)(s) { (s2, col) =>
                copyCell(s2, ARef.from0(col, sourceRow), ARef.from0(col, targetRow), 0, rowDelta)
              }
          }
        case Edit.FillDir.Right =>
          val sourceStartCol = source.colStart.index0
          val sourceColCount = source.colEnd.index0 - sourceStartCol + 1
          val targetStartCol = target.colStart.index0
          val rows = source.rowStart.index0 to source.rowEnd.index0
          foldE((targetStartCol to target.colEnd.index0).iterator)(sheet) { (s, targetCol) =>
            val sourceCol = sourceStartCol + (targetCol - targetStartCol) % sourceColCount
            val colDelta = targetCol - sourceCol
            if colDelta == 0 then Right(s)
            else
              foldE(rows.iterator)(s) { (s2, row) =>
                copyCell(s2, ARef.from0(sourceCol, row), ARef.from0(targetCol, row), colDelta, 0)
              }
          }
    }

  /** One fill step: the source cell's value, shifted, onto the target (styles are not copied). */
  private def copyCell(
    sheet: Sheet,
    sourceRef: ARef,
    targetRef: ARef,
    colDelta: Int,
    rowDelta: Int
  )(using FormulaSupport): XLResult[Sheet] =
    sheet.cells.get(sourceRef) match
      case None => Right(sheet)
      case Some(sourceCell) =>
        shifted(sourceCell.value, colDelta, rowDelta).map(v => sheet.put(targetRef, v))

  // ===== Copy =====

  /**
   * Copy `sourceRange` of `sourceSheet` into `target` at `targetRange` (same dimensions; the same
   * sheet when the names agree). Source cells are snapshotted before any write, so an overlapping
   * same-sheet copy reads the pre-copy state whatever the iteration order. Empty source cells leave
   * their target unchanged; styles travel with their cells (looked up in the source registry,
   * registered in the target's). `valuesOnly` pastes a formula's cached value — an uncached formula
   * pastes `Empty`; callers with an evaluator cache the source first.
   */
  def copyRange(
    target: Sheet,
    sourceSheet: Sheet,
    sourceRange: CellRange,
    targetRange: CellRange,
    valuesOnly: Boolean
  )(using FormulaSupport): XLResult[Sheet] =
    val snapshot: Map[ARef, Cell] =
      sourceRange.cells.flatMap(ref => sourceSheet.cells.get(ref).map(ref -> _)).toMap
    val colDelta = targetRange.colStart.index0 - sourceRange.colStart.index0
    val rowDelta = targetRange.rowStart.index0 - sourceRange.rowStart.index0
    val working = if sourceSheet.name == target.name then sourceSheet else target
    foldE(sourceRange.cells)(working) { (s, srcRef) =>
      val tgtRef = ARef.from0(srcRef.col.index0 + colDelta, srcRef.row.index0 + rowDelta)
      snapshot.get(srcRef) match
        case None => Right(s)
        case Some(srcCell) =>
          val copied: XLResult[CellValue] = srcCell.value match
            case CellValue.Formula(_, cachedOpt, _: FormulaKind.DataTable) =>
              Right(cachedOpt.getOrElse(CellValue.Empty))
            case CellValue.Formula(_, cachedOpt, _) if valuesOnly =>
              Right(cachedOpt.getOrElse(CellValue.Empty))
            case other => shifted(other, colDelta, rowDelta)
          copied.map { v =>
            val withValue = s.put(tgtRef, v)
            srcCell.styleId.flatMap(sourceSheet.styleRegistry.get) match
              case Some(style) => withValue.withCellStyle(tgtRef, style)
              case None => withValue
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
          if key.direction == Edit.SortDir.Descending then -cmp else cmp
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
   * has its relative row references shifted by the displacement and its cache dropped.
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
      if dataRowStart > rowEnd then Right(sheet)
      else
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
            else shifted(cell.value, 0, rowDelta)
          moved.map(v => s2.put(Cell(newRef, v, cell.styleId, None, cell.hyperlink)))
        }
    }

    placed.map { s =>
      val remappedComments =
        s.comments.map((ref, comment) => refMapping.getOrElse(ref, ref) -> comment)
      s.copy(comments = remappedComments)
    }

  // ===== Clear =====

  /**
   * Clear contents, styles and/or comments in `range`. Clearing contents also unmerges every merged
   * region overlapping the range, so no merge is left describing cleared cells.
   */
  def clear(sheet: Sheet, range: CellRange, what: ClearWhat): Sheet =
    val s1 = if what.contents then sheet.removeRange(range) else sheet
    val s2 = if what.styles then s1.clearStylesInRange(range) else s1
    val s3 = if what.comments then s2.clearCommentsInRange(range) else s2
    if what.contents then
      s3.mergedRanges.filter(_.intersects(range)).foldLeft(s3)((s, mr) => s.unmerge(mr))
    else s3

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
   * is used for that cell.
   */
  def autoFitWidth(sheet: Sheet, col: Column): Double =
    fitWidth(sheet, sheet.cells.valuesIterator.filter(_.ref.col == col).toVector)

  /**
   * [[autoFitWidth]] for every column of `columns`, in order, from ONE pass over the sheet's cells:
   * the cells are grouped by column once and each width is computed from its group, so fitting
   * every column of a wide sheet is O(cells + columns) rather than O(cells x columns). A column
   * without cells gets [[DefaultColumnWidth]]; a repeated column is reported each time.
   */
  def autoFitWidths(sheet: Sheet, columns: Iterable[Column]): Vector[(Column, Double)] =
    val byColumn: Map[Column, Vector[Cell]] = sheet.cells.valuesIterator.toVector.groupBy(_.ref.col)
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
    val text = RenderUtils.cellValueToText(cell.value, numFmt)
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
    else Left(XLError.Other(s"Outline level must be 1-7, got: $level"))

  /**
   * Apply `f` to the row's properties, keeping the entry even when the result is all-default: an
   * explicit entry is authoritative on write (actively clearing preserved source attributes),
   * whereas a missing entry lets a preserved `<row>` ride through verbatim — pruning would
   * resurrect the very collapsed/outlineLevel attrs an ungroup just cleared.
   */
  private def updateRow(sheet: Sheet, row: Row)(f: RowProperties => RowProperties): Sheet =
    sheet.setRowProperties(row, f(sheet.getRowProperties(row)))

  private def updateCol(sheet: Sheet, col: Column)(f: ColumnProperties => ColumnProperties): Sheet =
    sheet.setColumnProperties(col, f(sheet.getColumnProperties(col)))

  /**
   * Group rows into a collapsible outline: every row gets `level`; `collapsed` hides the members
   * and marks the summary row AFTER the group (Excel's summaryBelow default) so the "+" button
   * draws there.
   */
  def groupRows(sheet: Sheet, rows: RowSpan, level: Int, collapsed: Boolean): XLResult[Sheet] =
    validateLevel(level).map { _ =>
      val withMembers = rows.rows.foldLeft(sheet) { (s, r) =>
        updateRow(s, r)(p => p.copy(outlineLevel = Some(level), hidden = collapsed || p.hidden))
      }
      if collapsed && rows.end.index0 < Row.MaxIndex0 then
        updateRow(withMembers, rows.end + 1)(_.copy(collapsed = true))
      else withMembers
    }

  def groupCols(sheet: Sheet, cols: ColSpan, level: Int, collapsed: Boolean): XLResult[Sheet] =
    validateLevel(level).map { _ =>
      val withMembers = cols.columns.foldLeft(sheet) { (s, c) =>
        updateCol(s, c)(p => p.copy(outlineLevel = Some(level), hidden = collapsed || p.hidden))
      }
      if collapsed && cols.end.index0 < Column.MaxIndex0 then
        updateCol(withMembers, cols.end + 1)(_.copy(collapsed = true))
      else withMembers
    }

  /**
   * Clear the outline level and collapse markers of the rows and of the group's summary row.
   * Members a collapse hid stay hidden, like Excel (show them with a row-show edit).
   */
  def ungroupRows(sheet: Sheet, rows: RowSpan): Sheet =
    val cleared = rows.rows.foldLeft(sheet) { (s, r) =>
      updateRow(s, r)(_.copy(outlineLevel = None, collapsed = false))
    }
    if rows.end.index0 < Row.MaxIndex0 then
      updateRow(cleared, rows.end + 1)(_.copy(collapsed = false))
    else cleared

  def ungroupCols(sheet: Sheet, cols: ColSpan): Sheet =
    val cleared = cols.columns.foldLeft(sheet) { (s, c) =>
      updateCol(s, c)(_.copy(outlineLevel = None, collapsed = false))
    }
    if cols.end.index0 < Column.MaxIndex0 then
      updateCol(cleared, cols.end + 1)(_.copy(collapsed = false))
    else cleared

  // ===== Appearance & print setup (GH-358) =====

  /** Excel accepts a zoom of 10-400 percent. */
  def validateZoom(zoom: Option[Int]): XLResult[Unit] =
    zoom.filter(z => z < 10 || z > 400) match
      case Some(z) => Left(XLError.Other(s"Zoom scale must be 10-400, got: $z"))
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
