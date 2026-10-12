package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.error.XLError
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.sheets.styleSyntax.{getCellStyle, withCellStyle}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-714: author an Excel 365 dynamic-array formula — what typing `=SORT(B1:B3)` into a cell of
 * Excel 365 stores.
 *
 * The anchor holds the formula as a dynamic-array record whose `ref` is the spill extent, cached at
 * the array's top-left value; every other cell of the extent holds its element as a plain constant
 * (Excel's own on-disk shape). The extent comes from evaluating the formula as an array at the
 * anchor — a scalar result is a 1x1 record, which is how Excel stores `=SUM(A1:A10*B1:B10)` typed
 * into one cell (no implicit intersection).
 *
 * A spill Excel would block (`#SPILL!`) is refused and nothing is written: xl does not model a
 * blocked spill as a value, and a later recalculation would overwrite a written `#SPILL!` with the
 * array's top-left value. The refusal is typed ([[SpillBlocked]]), so a caller derives both its
 * message and its remedy from the blocker, never from English text. Pure and total.
 */
private[xl] object DynamicArrayAuthoring:

  /**
   * The authored sheet, the record's extent, every cell whose value the authoring changed (the old
   * extent it cleared plus the new one — the seeds of the dependent refresh), and whether the
   * anchor carries a cache (false when the formula could not be evaluated: it is then written as an
   * uncached 1x1 record, like any putf whose evaluation fails), and the parsed formula (so a caller
   * that inspects it — the CLI's date auto-format — does not parse it again).
   */
  final case class Authored(
    sheet: Sheet,
    extent: CellRange,
    touched: Set[ARef],
    cached: Boolean,
    expr: TExpr[?]
  ):

    /**
     * `numFmt` merged into the style of every cell of the extent (font, fill and borders kept) —
     * the one place the CLI's date auto-format and batch's `format` cover a spill.
     */
    def withNumFmt(numFmt: NumFmt): Authored =
      copy(sheet = extent.cells.foldLeft(sheet) { (s, r) =>
        s.withCellStyle(r, s.getCellStyle(r).getOrElse(CellStyle.default).withNumFmt(numFmt))
      })

  /** What stands in a dynamic array's way, with the blockers; each is a `#SPILL!` in Excel. */
  enum SpillBlocker derives CanEqual:
    /** Cells of the extent other than the anchor that hold a value or formula. */
    case Occupied(cells: Vector[ARef])

    /** Merged ranges the extent touches. */
    case Merged(ranges: Vector[CellRange])

    /** Other array records (anchor and range) reaching into the extent. */
    case ArrayRecord(records: Vector[(ARef, CellRange)])

    /** Tables (name and range) the extent overlaps. */
    case Table(tables: Vector[(String, CellRange)])

    /** A `rows`x`cols` spill that runs past XFD1048576. */
    case OffGrid(rows: Int, cols: Int)

    /**
     * A `rows`x`cols` spill of `limit` cells or more ([[ArrayResult.MaxCells]], one full column):
     * xl's own guard, not Excel's — Excel would spill it — refused rather than materialise that
     * many constants (and their styles) in memory.
     */
    case TooLarge(rows: Int, cols: Int, limit: Long)

    /**
     * A plain write into cells of an existing anchor's spill, leaving the anchor in place: Excel
     * keeps the written values and shows `#SPILL!` at the anchor.
     */
    case Overwritten(cells: Vector[ARef])

    /**
     * A clear of cells of an existing anchor's spill, leaving the anchor in place: Excel keeps a
     * live spill's cells, so the clear would change nothing there.
     */
    case Cleared(cells: Vector[ARef])

  /**
   * A spill Excel would block: the dynamic array `formula` (with its `=`) at `anchor`, the extent
   * it needs (None when it runs off the grid) and why. Message and hint both derive from `blocker`.
   */
  final case class SpillBlocked(
    formula: String,
    anchor: ARef,
    extent: Option[CellRange],
    blocker: SpillBlocker
  ) derives CanEqual:
    private def span: String = extent.fold(anchor.toA1)(_.toA1)

    /** The refusal as prose: the blockers, `(Excel shows #SPILL!)`, then the remedy in short. */
    def message: String = blocker match
      case SpillBlocker.Occupied(cells) =>
        val them = if cells.size == 1 then "it" else "them"
        s"spill range $span is blocked by ${listed(cells.map(_.toA1))} " +
          s"(Excel shows #SPILL!); clear $them or choose another anchor"
      case SpillBlocker.Merged(ranges) =>
        s"spill range $span is blocked by merged cells ${listed(ranges.map(_.toA1))} " +
          "(Excel shows #SPILL!); unmerge them or choose another anchor"
      case SpillBlocker.ArrayRecord(records) =>
        val named = records.map((ref, r) => s"${ref.toA1} (${r.toA1})")
        s"spill range $span is blocked by the array formula at ${listed(named)} " +
          "(Excel shows #SPILL!); choose another anchor"
      case SpillBlocker.Table(tables) =>
        val named = tables.map((name, r) => s"$name (${r.toA1})")
        s"spill range $span overlaps table ${listed(named)}; " +
          "a dynamic array cannot spill into a table (Excel shows #SPILL!); choose another anchor"
      case SpillBlocker.OffGrid(rows, cols) =>
        s"a ${rows}x$cols spill from ${anchor.toA1} runs past the sheet's last cell " +
          "XFD1048576 (Excel shows #SPILL!); choose another anchor"
      case SpillBlocker.TooLarge(rows, cols, limit) =>
        s"a ${rows}x$cols spill from ${anchor.toA1} is ${count(rows.toLong * cols)} cells; xl " +
          s"authors a dynamic array of fewer than ${count(limit)} cells (one full column, the " +
          "SEQUENCE and INDIRECT guard), so nothing was written"
      case SpillBlocker.Overwritten(cells) =>
        val they = if cells.size == 1 then "it is" else "they are"
        s"${listed(cells.map(_.toA1))}: $they in the spill range $span of the dynamic array " +
          s"$formula at ${anchor.toA1}; writing there would block that spill " +
          "(Excel shows #SPILL!)"
      case SpillBlocker.Cleared(cells) =>
        val they = if cells.size == 1 then "it is" else "they are"
        s"${listed(cells.map(_.toA1))}: $they in the spill range $span of the dynamic array " +
          s"$formula at ${anchor.toA1}; Excel keeps a spill's cells while its anchor stays, " +
          "so clearing them alone would change nothing"

    /** The next step for this blocker, naming what to act on. */
    def hint: String = blocker match
      case SpillBlocker.Occupied(cells) =>
        s"clear ${listed(cells.map(_.toA1))} or choose another anchor"
      case SpillBlocker.Merged(ranges) =>
        s"unmerge ${listed(ranges.map(_.toA1))} or choose another anchor"
      case SpillBlocker.ArrayRecord(records) =>
        s"clear or move the array formula at ${listed(records.map(_._1.toA1))}, " +
          "or choose another anchor"
      case SpillBlocker.Table(tables) =>
        s"choose an anchor whose spill misses table ${listed(tables.map(_._1))}, " +
          "or convert the table to a range"
      case SpillBlocker.OffGrid(rows, cols) =>
        s"choose an anchor with $rows row(s) and $cols column(s) of room before XFD1048576"
      case SpillBlocker.TooLarge(_, _, _) =>
        "bound the formula's ranges to the data (A1:A1000, not A:A), or write a plain formula"
      case SpillBlocker.Overwritten(_) =>
        s"overwrite or clear ${anchor.toA1} first (its spill goes with it), " +
          s"or write outside $span"
      case SpillBlocker.Cleared(_) =>
        s"clear ${anchor.toA1} to remove the whole spill, or clear outside $span"

    /** The domain error, for callers without a typed channel. */
    def toXLError: XLError = XLError.FormulaError(formula, message)

  /** Why an authoring wrote nothing: the formula does not parse, or its spill is blocked. */
  enum Refusal derives CanEqual:
    case Invalid(error: XLError)
    case Blocked(spill: SpillBlocked)

    def toXLError: XLError = this match
      case Invalid(error) => error
      case Blocked(spill) => spill.toXLError

    def message: String = toXLError.message

  private val maxListed = 10

  private def count(n: Long): String = String.format(java.util.Locale.ROOT, "%,d", n)

  def author(
    wb: Workbook,
    sheet: Sheet,
    anchor: ARef,
    formula: String,
    clock: Clock,
    maxCells: Long = ArrayResult.MaxCells
  ): Either[Refusal, Authored] =
    val text = CellValue.canonicalFormulaText(formula)
    val full = s"=$text"
    SheetEvaluator.parseFormula(full).left.map(Refusal.Invalid(_)).flatMap { expr =>
      val (cleared, clearedRefs) = clearOwnSpill(sheet, anchor)
      val single = CellRange(anchor, anchor)
      val placed =
        cleared.put(anchor, CellValue.Formula(text, None, FormulaKind.dynamicArray(single)))
      SheetEvaluator.evaluateArrayValues(
        placed,
        full,
        anchor,
        Evaluator.instance,
        clock,
        Some(wb.put(placed))
      ) match
        // a host failure (not an error value): an uncached 1x1 record, reported by the refresh
        case Left(_) => Right(Authored(placed, single, clearedRefs + anchor, cached = false, expr))
        case Right(grid0) =>
          val grid =
            if grid0.isEmpty || grid0.exists(_.isEmpty) then
              Vector(Vector(CellValue.Error(CellError.Calc)))
            else grid0
          val rows = grid.size
          val cols = grid.map(_.size).maxOption.getOrElse(1)
          def refuse(extent: Option[CellRange], blocker: SpillBlocker) =
            Left(Refusal.Blocked(SpillBlocked(full, anchor, extent, blocker)))
          // checked from the result's dimensions, before a cell of the spill is written
          if rows.toLong * cols.toLong >= maxCells then
            refuse(None, SpillBlocker.TooLarge(rows, cols, maxCells))
          else
            spillExtent(anchor, rows, cols) match
              case None => refuse(None, SpillBlocker.OffGrid(rows, cols))
              case Some(extent) =>
                blocked(placed, anchor, extent) match
                  case Some(blocker) => refuse(Some(extent), blocker)
                  case None =>
                    val kind = FormulaKind.dynamicArray(extent)
                    val written = grid.iterator.zipWithIndex.foldLeft(placed) {
                      case (s, (row, r)) =>
                        row.iterator.zipWithIndex.foldLeft(s) { case (acc, (v, c)) =>
                          if r == 0 && c == 0 then
                            acc.put(anchor, CellValue.Formula(text, Some(v), kind))
                          else acc.put(anchor.shift(c, r), v)
                        }
                    }
                    val touched = clearedRefs ++ extent.cells.toSet
                    Right(Authored(written, extent, touched, cached = true, expr))
    }

  /**
   * A plain (non-array) write of `targets` on `sheet`, applied to its dynamic arrays as Excel 365
   * applies it. A target that is a dynamic anchor releases its spill: the other cells of its extent
   * are cleared first ([[clearOwnSpill]]; the write then fills whichever of them it covers). A
   * target inside another anchor's spill whose anchor the write does not replace would block that
   * spill — Excel keeps the value and shows `#SPILL!` at the anchor, which xl cannot write (a
   * recalculation would overwrite it) — so it is refused, the first such anchor in reading order
   * named. Returns the released sheet and the cells it cleared (seeds of the dependent refresh).
   * `clearing` (a contents clear, not a write) only changes the refusal's wording: Excel keeps a
   * live spill's cells, so clearing them alone would do nothing ([[SpillBlocker.Cleared]]).
   *
   * Costs one pass over the sheet's cells; a sheet without dynamic arrays is returned unchanged.
   */
  def releaseForWrite(
    sheet: Sheet,
    targets: Set[ARef],
    clearing: Boolean = false
  ): Either[SpillBlocked, (Sheet, Set[ARef])] =
    if targets.isEmpty then Right((sheet, Set.empty))
    else releaseCells(sheet, spillsOf(sheet), targets, clearing)

  /**
   * A multi-cell dynamic-array anchor of a sheet: the anchor, its formula text (without `=`) and
   * its spill extent. A 1x1 record has no spill to release or intrude on.
   */
  final case class Spill(anchor: ARef, formula: String, extent: CellRange) derives CanEqual

  /**
   * Every [[Spill]] of `sheet`, in reading order — one pass over its cells. A caller applying many
   * writes to one sheet (a batch) computes this once and keeps it current instead of rescanning.
   */
  def spillsOf(sheet: Sheet): Vector[Spill] =
    sheet.cells.valuesIterator
      .flatMap(cell =>
        cell.value match
          case CellValue.Formula(text, _, arr: FormulaKind.ArrayFormula)
              if arr.isDynamicArray && arr.ref.start == cell.ref && arr.ref.cellCount > 1 =>
            Some(Spill(cell.ref, text, arr.ref))
          case _ => None
      )
      .toVector
      .sortBy(s => (s.anchor.row.index0, s.anchor.col.index0))

  /** [[releaseForWrite]] against `spills` already known for `sheet` (see [[spillsOf]]). */
  def releaseCells(
    sheet: Sheet,
    spills: Vector[Spill],
    targets: Set[ARef],
    clearing: Boolean
  ): Either[SpillBlocked, (Sheet, Set[ARef])] =
    // the targets, not the extent, are walked: a spill can be a whole column
    release(sheet, spills, targets.contains, extent => targets.iterator.filter(extent.contains))(
      clearing
    )

  /**
   * [[releaseForWrite]] for targets given as ranges, against `spills` already known for `sheet`: no
   * range is expanded into its cells unless a refusal must list them, so a batch op over a large
   * range costs a pass over the sheet's spills, not over the range.
   */
  def releaseRanges(
    sheet: Sheet,
    spills: Vector[Spill],
    targets: Vector[CellRange],
    clearing: Boolean
  ): Either[SpillBlocked, (Sheet, Set[ARef])] =
    release(
      sheet,
      spills,
      r => targets.exists(_.contains(r)),
      extent => targets.iterator.flatMap(overlap(_, extent)).flatMap(_.cellsRowMajor)
    )(clearing)

  /** The cells two ranges share, as a range; None when they do not meet. */
  private def overlap(a: CellRange, b: CellRange): Option[CellRange] =
    Option.when(a.intersects(b))(
      CellRange(
        ARef.from0(
          a.colStart.index0.max(b.colStart.index0),
          a.rowStart.index0.max(b.rowStart.index0)
        ),
        ARef.from0(a.colEnd.index0.min(b.colEnd.index0), a.rowEnd.index0.min(b.rowEnd.index0))
      )
    )

  private def release(
    sheet: Sheet,
    spills: Vector[Spill],
    targeted: ARef => Boolean,
    targetedIn: CellRange => Iterator[ARef]
  )(clearing: Boolean): Either[SpillBlocked, (Sheet, Set[ARef])] =
    if spills.isEmpty then Right((sheet, Set.empty))
    else
      val intrusion = spills.iterator
        .filterNot(s => targeted(s.anchor))
        .flatMap { s =>
          val hit = targetedIn(s.extent)
            .filter(_ != s.anchor)
            .toVector
            .distinct
            .sortBy(r => (r.row.index0, r.col.index0))
          val blocker =
            if clearing then SpillBlocker.Cleared(hit) else SpillBlocker.Overwritten(hit)
          Option.when(hit.nonEmpty)(
            SpillBlocked(s"=${s.formula}", s.anchor, Some(s.extent), blocker)
          )
        }
        .nextOption()
      intrusion.toLeft {
        spills
          .filter(s => targeted(s.anchor))
          .foldLeft((sheet, Set.empty[ARef])) { case ((current, acc), s) =>
            val (next, cleared) = clearOwnSpill(current, s.anchor)
            (next, acc ++ cleared)
          }
      }

  /**
   * Releasing a dynamic anchor's spill clears the cells its spill filled (values only; styles
   * stay), so re-authoring a shrinking spill or overwriting the anchor leaves no stale children. A
   * legacy (CSE) extent is left alone: its cells are the user's to clear.
   */
  private def clearOwnSpill(sheet: Sheet, anchor: ARef): (Sheet, Set[ARef]) =
    sheet.cells.get(anchor).map(_.value) match
      case Some(CellValue.Formula(_, _, arr: FormulaKind.ArrayFormula))
          if arr.isDynamicArray && arr.ref.start == anchor =>
        val children = arr.ref.cells
          .filter(_ != anchor)
          .filter { r =>
            sheet.cells.get(r).exists(_.value != CellValue.Empty)
          }
          .toSet
        (clearValues(sheet, children), children)
      case _ => (sheet, Set.empty)

  /**
   * `refs` emptied of their values; a cell that keeps a style or hyperlink stays as an empty cell,
   * any other is removed. What releasing a spill does to its cells, for a caller that releases
   * after its own write (copy and fill read the cells a release would clear).
   */
  def clearValues(sheet: Sheet, refs: Set[ARef]): Sheet =
    val cells = refs.foldLeft(sheet.cells) { (acc, r) =>
      acc.get(r) match
        case Some(cell) if cell.styleId.isDefined || cell.hyperlink.isDefined =>
          acc.updated(r, cell.withValue(CellValue.Empty))
        case Some(_) => acc.removed(r)
        case None => acc
    }
    sheet.copy(cells = cells)

  /** The extent of a `rows`x`cols` spill from `anchor`; None when it runs past XFD1048576. */
  private def spillExtent(anchor: ARef, rows: Int, cols: Int): Option[CellRange] =
    val lastCol = anchor.col.index0.toLong + cols - 1
    val lastRow = anchor.row.index0.toLong + rows - 1
    Option.when(lastCol <= Column.MaxIndex0 && lastRow <= Row.MaxIndex0)(
      CellRange(anchor, ARef.from0(lastCol.toInt, lastRow.toInt))
    )

  /**
   * Why Excel would show `#SPILL!` for `extent`, in Excel's own order of checks; None when clear.
   */
  private def blocked(sheet: Sheet, anchor: ARef, extent: CellRange): Option[SpillBlocker] =
    val occupied = extent.cellsRowMajor
      .filter(_ != anchor)
      .filter(r => sheet.cells.get(r).exists(_.value != CellValue.Empty))
      .toVector
    lazy val merges = sheet.mergedRanges.filter(_.intersects(extent)).toVector.sortBy(rowMajor)
    lazy val arrays = sheet.cells.valuesIterator
      .filter(_.ref != anchor)
      .flatMap(cell =>
        cell.value match
          case CellValue.Formula(_, _, arr: FormulaKind.ArrayFormula)
              if arr.ref.intersects(extent) =>
            Some(cell.ref -> arr.ref)
          case _ => None
      )
      .toVector
      .sortBy((ref, _) => (ref.row.index0, ref.col.index0))
    lazy val tables = sheet.tables.values
      .filter(_.range.intersects(extent))
      .toVector
      .sortBy(_.name)
      .map(t => t.name -> t.range)
    if occupied.nonEmpty then Some(SpillBlocker.Occupied(occupied))
    else if merges.nonEmpty then Some(SpillBlocker.Merged(merges))
    else if arrays.nonEmpty then Some(SpillBlocker.ArrayRecord(arrays))
    else if tables.nonEmpty then Some(SpillBlocker.Table(tables))
    else None

  private def rowMajor(r: CellRange): (Int, Int) = (r.start.row.index0, r.start.col.index0)

  /** At most ten names, then `… (+N more)`. */
  private def listed(names: Vector[String]): String =
    val shown = names.take(maxListed).mkString(", ")
    if names.size > maxListed then s"$shown, … (+${names.size - maxListed} more)" else shown
