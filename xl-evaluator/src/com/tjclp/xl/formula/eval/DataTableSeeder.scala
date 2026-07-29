package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * Explicit cache seeding for data-table interiors (GH-419).
 *
 * xl never evaluates `TABLE(...)` implicitly — a data-table record's cache is pinned as its only
 * truthful value (GH-430). Under `calcMode="autoNoTable"` (the house dialect) Excel itself never
 * recomputes tables on open either, so an uncached authored table would open BLANK. This seeder is
 * the explicit lever: it replays Excel's what-if substitution — overlay each axis value into the
 * input cell(s), evaluate the source formula, cache the result per interior cell.
 *
 * Semantics (tolerant, deterministic — never all-or-nothing):
 *   - `Right(v)` caches `v`, INCLUDING error VALUES (`#NUM!` interiors are legitimate data);
 *   - a `Left` (unsupported function, unresolvable axis) leaves that cell untouched and continues;
 *   - an empty or non-formula source cell caches `Number(0)` (Excel-verified bytes);
 *   - SHAPE-PRESERVING: DataTable-kind cells get `cachedValue` refreshed KEEPING their kind
 *     (foreign every-interior books keep per-cell faithfulness — only re-authoring normalizes);
 *     plain interior cells get their values replaced; real (non-record) formulas are never
 *     overwritten;
 *   - records with `del1`/`del2`, missing required inputs, no room for the corner/axes, or an
 *     interior larger than [[MaxInteriorArea]] cells (the record ref is file-controlled — a
 *     corrupt/hostile extent is a malformed record, never materialized) are skipped untouched;
 *   - a seed that changes NOTHING never marks a sheet modified, so a pristine book keeps the
 *     writer's verbatim-copy fidelity loop.
 *
 * Upstream formulas keep the landed pinned-cache doctrine: a cached precedent reads via its cache.
 * Cells are computed row-major against the group's base sheet, so results are order-independent
 * within a group; groups seed in row-major record-ref order, sheets in workbook order.
 */
object DataTableSeeder:

  /**
   * Interior-area skip cap, in cells. The record ref comes from the file, so a crafted
   * `B2:XFD1048576` interior (~1.7e10 cells) must skip like any other malformed record instead of
   * being iterated. 1,000,000 cells (a 1000x1000 grid) is orders beyond any real sensitivity table.
   */
  private val MaxInteriorArea: Long = 1000000L

  extension (wb: Workbook)

    /** Seed every data table on every sheet (system clock). */
    def seedDataTables(): XLResult[Workbook] =
      Right(wb.sheets.foldLeft(wb)((acc, sheet) => seedSheet(acc, sheet.name, None, Clock.system)))

    /** Seed every data table on one sheet: `Left(SheetNotFound)` for a bad name. */
    @annotation.targetName("seedDataTablesSheet")
    def seedDataTables(sheetName: SheetName): XLResult[Workbook] =
      seedDataTables(sheetName, None, Clock.system)

    /**
     * Seed the data tables on one sheet whose record ref intersects `only` (`None` = all), with an
     * explicit clock for volatile source formulas.
     */
    @annotation.targetName("seedDataTablesScoped")
    def seedDataTables(
      sheetName: SheetName,
      only: Option[CellRange],
      clock: Clock
    ): XLResult[Workbook] =
      if wb.sheets.exists(_.name == sheetName) then Right(seedSheet(wb, sheetName, only, clock))
      else Left(XLError.SheetNotFound(sheetName.value))

  private def seedSheet(
    wb: Workbook,
    sheetName: SheetName,
    only: Option[CellRange],
    clock: Clock
  ): Workbook =
    wb.sheets.find(_.name == sheetName) match
      case None => wb
      case Some(sheet) =>
        val groups = recordGroups(sheet).filter { case (ref, _) => only.forall(_.intersects(ref)) }
        if groups.isEmpty then wb
        else
          val seeded = groups.foldLeft(sheet) { case (acc, (_, kind)) =>
            seedGroup(wb.put(acc), acc, kind, clock)
          }
          // Only a real change may mark the sheet modified — a no-op seed (already-correct
          // caches, or nothing but skips) must keep the verbatim-copy write fidelity.
          if seeded == sheet then wb else wb.put(seeded)

  /**
   * Record groups by record ref, row-major; the representative kind comes from the group's
   * row-major-first cell (deterministic under any Map iteration order).
   */
  private def recordGroups(sheet: Sheet): Vector[(CellRange, FormulaKind.DataTable)] =
    sheet.cells.values.toVector
      .flatMap { cell =>
        cell.value match
          case CellValue.Formula(_, _, dt: FormulaKind.DataTable) => Vector((cell, dt))
          case _ => Vector.empty
      }
      .sortBy { case (cell, _) => (cell.ref.row.index0, cell.ref.col.index0) }
      .distinctBy { case (_, dt) => dt.ref }
      .map { case (_, dt) => (dt.ref, dt) }
      .sortBy { case (ref, _) => (ref.start.row.index0, ref.start.col.index0) }

  private def seedGroup(
    wb: Workbook,
    sheet: Sheet,
    kind: FormulaKind.DataTable,
    clock: Clock
  ): Sheet =
    val interior = kind.ref
    val startRow = interior.start.row.index0
    val startCol = interior.start.col.index0
    val oversized = interior.width.toLong * interior.height.toLong > MaxInteriorArea
    val inputs: Option[(ARef, Option[ARef])] =
      if kind.del1 || kind.del2 || startRow < 1 || startCol < 1 || oversized then None
      else
        (kind.r1, kind.r2) match
          case (Some(r1), Some(r2)) if kind.dt2D => Some((r1, Some(r2)))
          case (Some(r1), _) if !kind.dt2D => Some((r1, None))
          case _ => None // missing required input(s)
    inputs match
      case None => sheet // skip untouched
      case Some((input1, input2)) =>
        // Fold the interior iterator directly — O(1) beyond the accumulating sheet, never a
        // materialized extent. Every cell computes against the group's base `sheet`, so
        // interleaving the puts stays order-independent.
        interior.cellsRowMajor.foldLeft(sheet) { (acc, cellRef) =>
          computeCell(wb, sheet, kind, cellRef, input1, input2, clock) match
            case None => acc // tolerant: leave the cell untouched
            case Some(value) =>
              acc.cells.get(cellRef).map(_.value) match
                case Some(record @ CellValue.Formula(_, _, _: FormulaKind.DataTable)) =>
                  acc.put(cellRef, record.copy(cachedValue = Some(value))) // shape-preserving
                case Some(_: CellValue.Formula) => acc // never overwrite a real formula
                case _ => acc.put(cellRef, value)
        }

  /** One interior cell's what-if value; `None` leaves the cell untouched (tolerant). */
  private def computeCell(
    wb: Workbook,
    sheet: Sheet,
    kind: FormulaKind.DataTable,
    cellRef: ARef,
    input1: ARef,
    input2: Option[ARef],
    clock: Clock
  ): Option[CellValue] =
    val interior = kind.ref
    val startRow = interior.start.row.index0
    val startCol = interior.start.col.index0
    val col = cellRef.col.index0
    val row = cellRef.row.index0
    val topAxis = ARef.from0(col, startRow - 1)
    val leftAxis = ARef.from0(startCol - 1, row)

    // D4 geometry: the source formula and which axis feeds which input.
    val sourceRef =
      if kind.dt2D then ARef.from0(startCol - 1, startRow - 1) // corner
      else if kind.dtr then ARef.from0(startCol - 1, row) // row-oriented: left of THIS row
      else ARef.from0(col, startRow - 1) // column-oriented: above THIS column
    val overlayRefs: Vector[(ARef, ARef)] =
      input2 match
        case Some(colInput) => Vector(input1 -> topAxis, colInput -> leftAxis) // 2-D
        case None if kind.dtr => Vector(input1 -> topAxis) // 1-D row-oriented
        case None => Vector(input1 -> leftAxis) // 1-D column-oriented

    sheet.cells.get(sourceRef).map(_.value) match
      case Some(CellValue.Formula(expression, _, _)) =>
        val overlaid = overlayRefs.foldLeft(Option(sheet)) { case (acc, (inputRef, axisRef)) =>
          acc.flatMap { s =>
            SheetEvaluator
              .evaluateCell(sheet)(axisRef, clock, Some(wb))
              .toOption
              .map(axisValue => s.put(inputRef, axisValue))
          }
        }
        overlaid.flatMap { s =>
          SheetEvaluator
            .evaluateFormula(s)(expression, clock, Some(wb.put(s)), None)
            .toOption
        }
      case _ =>
        // Empty or non-formula source cell: Excel caches 0 (fixture-verified bytes).
        Some(CellValue.Number(BigDecimal(0)))
