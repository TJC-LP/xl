package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{Cell, CellValue, FormulaKind}
import com.tjclp.xl.error.{XLError, XLResult}

// ========== Data Table Authoring (GH-419) ==========

/**
 * Native what-if Data Table authoring — the house sensitivity engine.
 *
 * Geometry (uniform across shapes): `interior` is the RESULT GRID = the emitted record `ref`, NOT
 * Excel's UI selection block. Worked example for `interior = D5:F6`:
 *   - corner formula: C4 (one-up-one-left of the interior start),
 *   - row axis (values substituted into the row input): D4:F4 (the row directly above),
 *   - column axis (values substituted into the column input): C5:C6 (the column directly left).
 *
 * Only the interior start receives the `<f t="dataTable"/>` record cell (corner-only, exactly what
 * Excel writes); every other interior cell stays a plain cached value. Authoring emits Excel's own
 * record dialect: `dt2D`/`dtr` explicit, `ca="1"`, the single 1-D input always riding `r1`.
 *
 * Calc-mode doctrine (empirically proven): `fullCalcOnLoad` does NOT recompute data tables under
 * `calcMode="autoNoTable"` (the TJC house dialect) — an uncached table opens BLANK. Ship `seeds`
 * (or run `wb.seedDataTables()` from xl-evaluator) when targeting autoNoTable books; plain
 * `calcMode="auto"` books self-heal on open. Calc levers stay on `Workbook.withCalcPr` (GH-400) —
 * authoring never mutates calcPr.
 *
 * Absorb rule: when the corner cell of an unseeded table already holds a plain scalar (e.g. a
 * `fillBy` result), that scalar is ABSORBED as the record's cached value — `fillBy` then
 * `dataTable` composes into a seeded grid. Re-authoring over an existing table consumes its records
 * and keeps their Excel-written caches.
 */
object dataTableSyntax:

  extension (sheet: Sheet)

    /**
     * Author a two-variable what-if data table over `interior` (see [[dataTableSyntax]] for the
     * geometry). `rowInput` receives the top-axis values (`r1`), `colInput` the left-axis values
     * (`r2`); both must live outside the table block. `seeds` (row-major, interior-shaped) become
     * the cached interior values; `seeds(0)(0)` becomes the corner record's cache.
     */
    def dataTable(
      interior: CellRange,
      rowInput: ARef,
      colInput: ARef,
      seeds: Seq[Seq[CellValue]] = Nil
    ): XLResult[Sheet] =
      DataTableAuthoring.author(
        sheet,
        interior,
        DataTableAuthoring.Shape.TwoD(rowInput, colInput),
        seeds
      )

    /** Runtime-parsed [[dataTable]]: `Left` on an unparseable range or input reference. */
    @annotation.targetName("dataTableString")
    def dataTable(interior: String, rowInput: String, colInput: String): XLResult[Sheet] =
      for
        parsedInterior <- DataTableAuthoring.parseRange(interior)
        parsedRow <- DataTableAuthoring.parseRef(rowInput)
        parsedCol <- DataTableAuthoring.parseRef(colInput)
        authored <- dataTable(parsedInterior, parsedRow, parsedCol, Nil)
      yield authored

    /**
     * Author a 1-D ROW-ORIENTED data table: axis values ride the row directly above `interior`; the
     * source formulas live in the column directly left (one per interior row). The single input
     * rides `r1` (`dtr="1"`).
     */
    def dataTableRow(
      interior: CellRange,
      rowInput: ARef,
      seeds: Seq[Seq[CellValue]] = Nil
    ): XLResult[Sheet] =
      DataTableAuthoring.author(sheet, interior, DataTableAuthoring.Shape.OneDRow(rowInput), seeds)

    /** Runtime-parsed [[dataTableRow]]: `Left` on an unparseable range or input reference. */
    @annotation.targetName("dataTableRowString")
    def dataTableRow(interior: String, rowInput: String): XLResult[Sheet] =
      for
        parsedInterior <- DataTableAuthoring.parseRange(interior)
        parsedRow <- DataTableAuthoring.parseRef(rowInput)
        authored <- dataTableRow(parsedInterior, parsedRow, Nil)
      yield authored

    /**
     * Author a 1-D COLUMN-ORIENTED data table: axis values ride the column directly left of
     * `interior`; the source formulas live in the row directly above (one per interior column —
     * multi-result-column tables are legal). The single input rides `r1` (`dtr="0"`).
     */
    def dataTableCol(
      interior: CellRange,
      colInput: ARef,
      seeds: Seq[Seq[CellValue]] = Nil
    ): XLResult[Sheet] =
      DataTableAuthoring.author(sheet, interior, DataTableAuthoring.Shape.OneDCol(colInput), seeds)

    /** Runtime-parsed [[dataTableCol]]: `Left` on an unparseable range or input reference. */
    @annotation.targetName("dataTableColString")
    def dataTableCol(interior: String, colInput: String): XLResult[Sheet] =
      for
        parsedInterior <- DataTableAuthoring.parseRange(interior)
        parsedCol <- DataTableAuthoring.parseRef(colInput)
        authored <- dataTableCol(parsedInterior, parsedCol, Nil)
      yield authored

/** Validation + materialization engine behind [[dataTableSyntax]] (GH-419). */
private[sheets] object DataTableAuthoring:

  enum Shape derives CanEqual:
    case TwoD(rowInput: ARef, colInput: ARef)
    case OneDRow(rowInput: ARef)
    case OneDCol(colInput: ARef)

  def parseRange(s: String): XLResult[CellRange] =
    CellRange.parse(s).left.map(reason => XLError.InvalidRange(s, reason))

  def parseRef(s: String): XLResult[ARef] =
    ARef.parse(s).left.map(reason => XLError.InvalidCellRef(s, reason))

  def author(
    sheet: Sheet,
    interior: CellRange,
    shape: Shape,
    seeds: Seq[Seq[CellValue]]
  ): XLResult[Sheet] =
    for
      _ <- validGeometry(interior) // V1
      _ <- validDistinctInputs(shape) // V2
      _ <- validInputsOutsideBlock(interior, shape) // V3
      _ <- validSourceFormulas(sheet, interior, shape) // V4
      consumed <- validOverlap(sheet, interior) // V5
      _ <- validInteriorFormulaFree(sheet, interior) // V6
      _ <- validUnmerged(sheet, interior) // V7
      _ <- validSeeds(interior, seeds) // V8
    yield materialize(sheet, interior, shape, seeds, consumed)

  private def anchorOf(interior: CellRange): ARef =
    ARef.from0(interior.start.col.index0 - 1, interior.start.row.index0 - 1)

  /** The Excel UI selection block: corner + axis row/column + interior. */
  private def blockOf(interior: CellRange): CellRange =
    CellRange(anchorOf(interior), interior.end)

  // V1: room for the corner formula and both axes.
  private def validGeometry(interior: CellRange): XLResult[Unit] =
    if interior.start.row.index0 >= 1 && interior.start.col.index0 >= 1 then Right(())
    else
      Left(
        XLError.InvalidRange(
          interior.toA1,
          "a data table interior needs the corner formula one-up-one-left, the axis row above " +
            "and the axis column left (interior D5:F6 -> corner C4, row axis D4:F4, column axis " +
            "C5:C6), so it cannot start in row 1 or column A"
        )
      )

  // V2: a 2-D table substitutes two DIFFERENT input cells.
  private def validDistinctInputs(shape: Shape): XLResult[Unit] =
    shape match
      case Shape.TwoD(rowInput, colInput) if rowInput == colInput =>
        Left(
          XLError.InvalidReference(
            s"data table row input and column input must be two different cells, both were " +
              s"${rowInput.toA1}"
          )
        )
      case _ => Right(())

  // V3: Excel refuses input cells anywhere in the table block. Cross-sheet inputs are
  // unrepresentable by construction (ARef is sheet-local), matching Excel.
  private def validInputsOutsideBlock(interior: CellRange, shape: Shape): XLResult[Unit] =
    val block = blockOf(interior)
    val inputs = shape match
      case Shape.TwoD(rowInput, colInput) => Vector(rowInput, colInput)
      case Shape.OneDRow(rowInput) => Vector(rowInput)
      case Shape.OneDCol(colInput) => Vector(colInput)
    inputs.find(block.contains) match
      case Some(inside) =>
        Left(
          XLError.InvalidCellRef(
            inside.toA1,
            "input cell reference is not valid — Excel refuses inputs anywhere in the table block"
          )
        )
      case None => Right(())

  /** A cell value that can drive a data table: any formula that is not itself a record. */
  private def isSourceFormula(value: CellValue): Boolean =
    value match
      case CellValue.Formula(_, _, _: FormulaKind.DataTable) => false
      case _: CellValue.Formula => true
      case _ => false

  // V4: source formulas present (deliberately stricter than Excel's empty-corner tolerance —
  // an empty corner authors a dead grid that computes all zeros).
  private def validSourceFormulas(
    sheet: Sheet,
    interior: CellRange,
    shape: Shape
  ): XLResult[Unit] =
    val startCol = interior.start.col.index0
    val startRow = interior.start.row.index0
    val sourceRefs: Vector[ARef] = shape match
      case _: Shape.TwoD => Vector(anchorOf(interior))
      case _: Shape.OneDCol =>
        // One source formula above EVERY interior column (multi-result-column tables are legal).
        (startCol to interior.end.col.index0).toVector.map(ARef.from0(_, startRow - 1))
      case _: Shape.OneDRow =>
        // One source formula left of EVERY interior row.
        (startRow to interior.end.row.index0).toVector.map(ARef.from0(startCol - 1, _))
    sourceRefs.find(r => !sheet.cells.get(r).map(_.value).exists(isSourceFormula)) match
      case Some(missing) =>
        Left(
          XLError.FormulaError(
            missing.toA1,
            s"data table over ${interior.toA1} needs a source formula at ${missing.toA1} — put " +
              "the corner formula first, e.g. sheet.put(ref\"C4\", fx\"=B14\"); the low-level " +
              "escape hatch is CellValue.dataTable + put"
          )
        )
      case None => Right(())

  /** Every DataTable-record cell on the sheet, first offender selection in row-major order. */
  private def recordCells(sheet: Sheet): Vector[(Cell, FormulaKind.DataTable)] =
    sheet.cells.values.toVector
      .flatMap { cell =>
        cell.value match
          case CellValue.Formula(_, _, dt: FormulaKind.DataTable) => Vector((cell, dt))
          case _ => Vector.empty
      }
      .sortBy { case (cell, _) => (cell.ref.row.index0, cell.ref.col.index0) }

  private def containsRange(outer: CellRange, inner: CellRange): Boolean =
    outer.contains(inner.start) && outer.contains(inner.end)

  // V5: an existing record intersecting the interior must be fully contained (then it is
  // CONSUMED — the refresh/re-author/normalize case, incl. foreign every-interior books).
  // Partial overlap from outside would tear that table. Record cells physically inside the
  // interior are consumed regardless of their record ref, so the authored grid can never end
  // up carrying a foreign record.
  private def validOverlap(sheet: Sheet, interior: CellRange): XLResult[Vector[Cell]] =
    val records = recordCells(sheet)
    val overlapping = records.filter { case (_, dt) => dt.ref.intersects(interior) }
    overlapping.find { case (_, dt) => !containsRange(interior, dt.ref) } match
      case Some((_, dt)) =>
        Left(
          XLError.InvalidRange(
            dt.ref.toA1,
            s"authoring a data table over ${interior.toA1} would tear data table at " +
              s"${dt.ref.toA1}; Excel: cannot change part of a data table"
          )
        )
      case None =>
        val inside = records.filter { case (cell, _) => interior.contains(cell.ref) }
        Right((overlapping ++ inside).map(_._1).distinctBy(_.ref))

  // V6: no silent destruction of real formulas inside the interior (plain scalars are fine —
  // they are seeds).
  private def validInteriorFormulaFree(sheet: Sheet, interior: CellRange): XLResult[Unit] =
    val offender = interior.cellsRowMajor
      .flatMap(sheet.cells.get)
      .collectFirst {
        case cell if isSourceFormula(cell.value) =>
          cell.value match
            case CellValue.Formula(expression, _, _) => (cell.ref, expression)
            case _ => (cell.ref, "")
      }
    offender match
      case Some((ref, expression)) =>
        Left(
          XLError.FormulaError(
            expression,
            s"data table interior ${interior.toA1} would overwrite the formula at " +
              s"${ref.toA1} — clear ${ref.toA1} first (removeRange) or bake values"
          )
        )
      case None => Right(())

  // V7: merged cells inside a data table interior are conservative-refused.
  private def validUnmerged(sheet: Sheet, interior: CellRange): XLResult[Unit] =
    sheet.mergedRanges.toVector
      .sortBy(m => (m.start.row.index0, m.start.col.index0))
      .find(_.intersects(interior)) match
      case Some(merge) =>
        Left(
          XLError.InvalidRange(
            merge.toA1,
            s"data table interior ${interior.toA1} intersects merged range ${merge.toA1} — " +
              "unmerge first"
          )
        )
      case None => Right(())

  // V8: seeds are optional; when present they must be interior-shaped plain values. Validated
  // BEFORE construction so CellValue.formula's require can never throw.
  private def validSeeds(interior: CellRange, seeds: Seq[Seq[CellValue]]): XLResult[Unit] =
    val context = s"data table interior ${interior.toA1}"
    if seeds.isEmpty then Right(())
    else if seeds.length != interior.height then
      Left(XLError.ValueCountMismatch(interior.height, seeds.length, context))
    else
      seeds.find(_.length != interior.width) match
        case Some(row) => Left(XLError.ValueCountMismatch(interior.width, row.length, context))
        case None =>
          val formulaSeed = seeds.iterator.zipWithIndex
            .flatMap { case (row, r) => row.iterator.zipWithIndex.map { case (v, c) => (v, r, c) } }
            .collectFirst { case (formula: CellValue.Formula, r, c) => (formula, r, c) }
          formulaSeed match
            case Some((formula, r, c)) =>
              Left(
                XLError.FormulaError(
                  formula.expression,
                  s"$context seeds must be plain cached values — the seed at row $r column $c " +
                    "is a formula"
                )
              )
            case None => Right(())

  private def kindOf(interior: CellRange, shape: Shape): FormulaKind.DataTable =
    shape match
      case Shape.TwoD(rowInput, colInput) =>
        FormulaKind.DataTable(
          ref = interior,
          dt2D = true,
          dtr = true,
          r1 = Some(rowInput),
          r2 = Some(colInput),
          del1 = false,
          del2 = false,
          ca = true
        )
      case Shape.OneDRow(rowInput) =>
        FormulaKind.DataTable(
          ref = interior,
          dt2D = false,
          dtr = true,
          r1 = Some(rowInput),
          r2 = None,
          del1 = false,
          del2 = false,
          ca = true
        )
      case Shape.OneDCol(colInput) =>
        // The single input ALWAYS rides r1 (Excel-verified; matches FormulaKind.displayExpression).
        FormulaKind.DataTable(
          ref = interior,
          dt2D = false,
          dtr = false,
          r1 = Some(colInput),
          r2 = None,
          del1 = false,
          del2 = false,
          ca = true
        )

  private def materialize(
    sheet: Sheet,
    interior: CellRange,
    shape: Shape,
    seeds: Seq[Seq[CellValue]],
    consumed: Vector[Cell]
  ): Sheet =
    val corner = interior.start

    // 1. Consumed records revert to their cached value as a plain value (cache-less ones drop) —
    //    except the new corner, whose cell is rebuilt as the record below.
    val reverted = consumed.foldLeft(sheet) { (acc, cell) =>
      if cell.ref == corner then acc
      else
        cell.value match
          case CellValue.Formula(_, Some(cached), _) => acc.put(cell.ref, cached)
          case _ => acc.remove(cell.ref)
    }

    // 2. Corner record cache: seeds(0)(0) when seeded; else the old corner record's Excel-written
    //    cache (re-author keeps caches); else an absorbed plain scalar previously at the corner
    //    (fillBy-then-dataTable composes); else None.
    val seedCache = seeds.headOption.flatMap(_.headOption).filter(_ != CellValue.Empty)
    val cache = seedCache.orElse {
      sheet.cells.get(corner).map(_.value) match
        case Some(CellValue.Formula(_, cached, _: FormulaKind.DataTable)) => cached
        case Some(_: CellValue.Formula) => None // unreachable: V6 rejected real formulas
        case Some(CellValue.Empty) | None => None
        case Some(scalar) => Some(scalar) // absorb rule
    }
    val withCorner = reverted.put(corner, CellValue.dataTable(kindOf(interior, shape), cache))

    // 3. Non-corner interiors: seeded -> plain seed values (style-preserving put); unseeded ->
    //    untouched (reverted caches from step 1 stay).
    if seeds.isEmpty then withCorner
    else
      interior.cellsRowMajor
        .zip(seeds.iterator.flatten)
        .foldLeft(withCorner) { case (acc, (ref, value)) =>
          if ref == corner then acc else acc.put(ref, value)
        }
