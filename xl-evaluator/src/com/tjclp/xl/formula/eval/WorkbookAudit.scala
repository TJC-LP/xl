package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.formula.ast.{BinarySpine, TExpr}
import com.tjclp.xl.formula.functions.{ArgValue, FunctionRegistry}
import com.tjclp.xl.formula.graph.{DependencyGraph, QualifiedGraph}
import com.tjclp.xl.formula.graph.DependencyGraph.{QualifiedRef, Scc}
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{CalcPr, Workbook}

/**
 * #678: one data table whose interior caches disagree with the corner formula re-evaluated at each
 * cell's input pair — the evaluation `xl recalc --tables` seeds with. The file is valid (Excel
 * never recomputes tables under `calcMode="autoNoTable"`, and xl's own edits leave interiors as
 * they were, the same parity), so this is an audit NOTE, not a finding: the numbers in the grid
 * describe an older model.
 *
 * @param sampled
 *   the interior cells re-evaluated, row-major: every cell of a small table, otherwise
 *   [[WorkbookAudit.DataTableSampleSize]] cells spread evenly over the interior (first and last
 *   included) — the check is bounded per table, and says which cells it looked at
 * @param stale
 *   the sampled cells whose cache differs from the re-evaluation: (cell, cached, recomputed)
 */
final case class StaleDataTable(
  sheet: SheetName,
  ref: CellRange,
  sampled: Vector[ARef],
  stale: Vector[(ARef, CellValue, CellValue)]
) derives CanEqual:

  /**
   * e.g. `3 of 8 sampled interior cells differ from the corner re-evaluated at their inputs (F10:
   * cached 2, recomputed 3; sampled F10, F11, …) — run xl recalc --tables`
   */
  def render: String =
    val shownSample = sampled.take(4).map(_.toA1).mkString(", ") +
      (if sampled.sizeIs > 4 then ", …" else "")
    val first = stale.headOption.fold("") { (r, cached, recomputed) =>
      s"${r.toA1}: cached ${StaleDataTable.show(cached)}, recomputed ${StaleDataTable.show(recomputed)}; "
    }
    s"${stale.size} of ${sampled.size} sampled interior cell(s) differ from the corner " +
      s"re-evaluated at their inputs (${first}sampled $shownSample) — run xl recalc --tables"

object StaleDataTable:
  /** A cell value as the audit prints it: a number plainly, an error as Excel spells it. */
  def show(value: CellValue): String = value match
    case CellValue.Number(n) => n.bigDecimal.stripTrailingZeros.toPlainString
    case CellValue.Text(t) => s"\"$t\""
    case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
    case CellValue.Error(e) => e.toExcel
    case other => other.toString

/**
 * ADR-017 §2.10: every reason a number in the workbook can be wrong, bucketed in one pass — what
 * `xl audit` prints and `wb.audit` returns.
 *
 * Findings (they make [[isClean]] false):
 *   - `errorCells` — a cached Excel error on a formula, or a bare error value
 *   - `uncachedFormulas` — formulas with no cached value (`xl recalc` fills them)
 *   - `unparseable` — formulas this evaluator cannot parse, each with the parser's one-line reason
 *   - `cycles` — the cyclic strongly connected components of the formula graph, when the file's
 *     `calcPr` does NOT enable iterative calculation (Excel shows such a book with a circular
 *     reference warning and zeros)
 *   - `unresolvedReaders` — formulas whose name references the static graph cannot resolve
 *     (GH-507); an unparseable formula is reported once, in `unparseable`
 *
 * Notes (reported, never findings): `iterativeCycles` (the same cyclic components when
 * `calcPr.iterativeCalculation` is on — an intentional circular model converges by design, so
 * `audit --fail-on-findings` must not fail it forever; every cycle lands in exactly one of `cycles`
 * and `iterativeCycles`), `volatile` (a call to a function flagged `FunctionFlags.volatile` —
 * TODAY, NOW, RAND, RANDBETWEEN — read off the parsed call, GH-588), `dynamic` (INDIRECT/OFFSET
 * readers, [[DependencyGraph.dynamicCells]]), `externalRefs` (formulas touching another workbook,
 * whose caches are pinned) and the file's `calcPr`.
 *
 * Every bucket is in workbook order — sheet position, then row, then column — so two runs on the
 * same file print the same report. Pure and total.
 */
final case class WorkbookAudit(
  errorCells: Vector[(QualifiedRef, CellError)],
  uncachedFormulas: Vector[QualifiedRef],
  unparseable: Vector[(QualifiedRef, String)],
  volatile: Vector[QualifiedRef],
  dynamic: Vector[QualifiedRef],
  cycles: Vector[Scc],
  externalRefs: Vector[QualifiedRef],
  unresolvedReaders: Vector[QualifiedRef],
  calcPr: Option[CalcPr],
  iterativeCycles: Vector[Scc],
  staleDataTables: Vector[StaleDataTable] = Vector.empty
) derives CanEqual:

  /** The number of findings: error cells, uncached and unparseable formulas, cycles, unresolved. */
  def findings: Int =
    errorCells.size + uncachedFormulas.size + unparseable.size + cycles.size +
      unresolvedReaders.size

  /**
   * No findings. Iterative cycles, volatile, dynamic and external references and `calcPr` do not
   * count.
   */
  def isClean: Boolean = findings == 0

  /**
   * The same audit seen from one sheet (`xl audit -s`): cell buckets keep that sheet's cells, a
   * cycle stays when any member is on the sheet, `calcPr` is workbook-wide and stays as it is.
   */
  def restrictTo(sheet: SheetName): WorkbookAudit =
    copy(
      errorCells = errorCells.filter(_._1.sheet == sheet),
      uncachedFormulas = uncachedFormulas.filter(_.sheet == sheet),
      unparseable = unparseable.filter(_._1.sheet == sheet),
      volatile = volatile.filter(_.sheet == sheet),
      dynamic = dynamic.filter(_.sheet == sheet),
      cycles = cycles.filter(_.members.exists(_.sheet == sheet)),
      externalRefs = externalRefs.filter(_.sheet == sheet),
      unresolvedReaders = unresolvedReaders.filter(_.sheet == sheet),
      iterativeCycles = iterativeCycles.filter(_.members.exists(_.sheet == sheet)),
      staleDataTables = staleDataTables.filter(_.sheet == sheet)
    )

object WorkbookAudit:

  /**
   * Upper-case names of the functions flagged `FunctionFlags.volatile` (GH-588) — informational;
   * the audit itself reads the flag off each parsed call, never this set.
   */
  lazy val volatileFunctions: Set[String] = FunctionRegistry.volatileFunctionNames.toSet

  val clean: WorkbookAudit = WorkbookAudit(
    errorCells = Vector.empty,
    uncachedFormulas = Vector.empty,
    unparseable = Vector.empty,
    volatile = Vector.empty,
    dynamic = Vector.empty,
    cycles = Vector.empty,
    externalRefs = Vector.empty,
    unresolvedReaders = Vector.empty,
    calcPr = None,
    iterativeCycles = Vector.empty
  )

  /**
   * #678: the most interior cells the data-table-stale note re-evaluates per table. Eight cells
   * spread over the interior catch a table left behind by a model change (every cell moves) at a
   * bounded cost on a 1000×1000 sensitivity grid.
   */
  val DataTableSampleSize: Int = 8

  /**
   * The audit with the system clock — the one a data table whose corner reads TODAY/NOW is
   * re-evaluated at (see [[of(wb:com\.tjclp\.xl\.workbooks\.Workbook,clock* the clocked variant]]).
   */
  def of(wb: Workbook): WorkbookAudit = of(wb, Clock.system)

  /** The audit, re-evaluating data-table corners (the stale-interior note) at `clock`. */
  def of(wb: Workbook, clock: Clock): WorkbookAudit =
    // Reverse insertion so a duplicated sheet name keeps its FIRST position, like indexWhere.
    val position: Map[SheetName, Int] =
      wb.sheets.zipWithIndex.reverseIterator.map((sheet, i) => sheet.name -> i).toMap
    given Ordering[QualifiedRef] =
      Ordering.by(q =>
        (position.getOrElse(q.sheet, Int.MaxValue), q.ref.row.index0, q.ref.col.index0)
      )

    val scanned: Vector[Finding] = wb.sheets.iterator.flatMap { sheet =>
      sheet.cells.iterator.flatMap { (ref, cell) =>
        classify(QualifiedRef(sheet.name, ref), cell.value)
      }
    }.toVector
    val unparseable = scanned.collect { case Finding.Unparseable(q, message) => (q, message) }
    val unparseableRefs = unparseable.iterator.map(_._1).toSet
    val graph = QualifiedGraph.of(wb)
    // sccs come dependency-first; the report reads in workbook order like every other bucket
    val cyclic = graph.sccs.filter(_.cyclic).sortBy(scc => scc.members.sorted.headOption)
    // With iterative calculation on, a circular model is the file's declared intent, not a finding
    val iterative = wb.metadata.calcPr.exists(_.iterativeCalculation)

    WorkbookAudit(
      errorCells = scanned.collect { case Finding.ErrorValue(q, e) => (q, e) }.sortBy(_._1),
      uncachedFormulas = scanned.collect { case Finding.Uncached(q) => q }.sorted,
      unparseable = unparseable.sortBy(_._1),
      volatile = scanned.collect { case Finding.Volatile(q) => q }.sorted,
      dynamic = DependencyGraph.dynamicCells(wb).toVector.sorted,
      cycles = if iterative then Vector.empty else cyclic,
      externalRefs = scanned.collect { case Finding.External(q) => q }.sorted,
      unresolvedReaders =
        DependencyGraph.unresolvedReaders(wb).iterator.filterNot(unparseableRefs).toVector.sorted,
      calcPr = wb.metadata.calcPr,
      iterativeCycles = if iterative then cyclic else Vector.empty,
      staleDataTables = staleDataTablesOf(wb, clock)
    )

  /**
   * #678: every data table whose sampled interior disagrees with the `recalc --tables` evaluation,
   * in workbook order then record-ref row-major. Cells the evaluation cannot settle (a failing
   * corner, a skipped or oversized table) come back unchanged and never read as stale; an uncached
   * cell is `uncachedFormulas`' to report, not this note's.
   */
  private def staleDataTablesOf(wb: Workbook, clock: Clock): Vector[StaleDataTable] =
    val tables = wb.sheets.flatMap(sheet => DataTableSeeder.dataTables(sheet).map(sheet -> _))
    if tables.isEmpty then Vector.empty
    else
      val seeded = DataTableSeeder.seedSample(wb, clock, interior => sampleOf(interior).iterator)
      val after = seeded.sheets.iterator.map(s => s.name -> s).toMap
      tables.flatMap { case (sheet, (interior, _)) =>
        val sampled = sampleOf(interior)
        val recomputedSheet = after.getOrElse(sheet.name, sheet)
        val stale = sampled.flatMap { ref =>
          (interiorCache(sheet, ref), interiorCache(recomputedSheet, ref)) match
            case (Some(cached), Some(recomputed)) if !sameValue(cached, recomputed) =>
              Vector((ref, cached, recomputed))
            case _ => Vector.empty
        }
        Option.when(stale.nonEmpty)(StaleDataTable(sheet.name, interior, sampled, stale))
      }

  /**
   * The cells the note re-evaluates: all of a small interior, else [[DataTableSampleSize]] cells at
   * evenly spaced row-major positions, first and last included.
   */
  private def sampleOf(interior: CellRange): Vector[ARef] =
    val width = interior.width.toLong
    val count = width * interior.height.toLong
    if count <= DataTableSampleSize then interior.cellsRowMajor.toVector
    else
      val startCol = interior.start.col.index0
      val startRow = interior.start.row.index0
      (0 until DataTableSampleSize).toVector
        .map(i => i.toLong * (count - 1) / (DataTableSampleSize - 1))
        .distinct
        .map(idx => ARef.from0(startCol + (idx % width).toInt, startRow + (idx / width).toInt))

  /**
   * An interior cell's cached value: a record's cache, a plain value; nothing for anything else.
   */
  private def interiorCache(sheet: Sheet, ref: ARef): Option[CellValue] =
    sheet.cells.get(ref).map(_.value).flatMap {
      case CellValue.Formula(_, cached, _: FormulaKind.DataTable) => cached
      case _: CellValue.Formula => None // a real formula in the interior: data-table-torn's class
      case CellValue.Empty => None
      case value => Some(value)
    }

  /**
   * Equal as a grid reader sees them: numbers within a relative 1e-9 (Excel's cache and xl's
   * re-evaluation may differ in the last binary digits — `0.1+0.2`), everything else exactly.
   */
  private def sameValue(a: CellValue, b: CellValue): Boolean = (a, b) match
    case (CellValue.Number(x), CellValue.Number(y)) =>
      (x - y).abs <= RelativeTolerance * (BigDecimal(1).max(x.abs).max(y.abs))
    case _ => a == b

  private val RelativeTolerance = BigDecimal("1E-9")

  /** One observation about one cell; a cell can yield several (an uncached volatile formula). */
  private enum Finding derives CanEqual:
    case ErrorValue(ref: QualifiedRef, error: CellError)
    case Uncached(ref: QualifiedRef)
    case Unparseable(ref: QualifiedRef, message: String)
    case Volatile(ref: QualifiedRef)
    case External(ref: QualifiedRef)

  private def classify(q: QualifiedRef, value: CellValue): List[Finding] = value match
    case CellValue.Error(e) => List(Finding.ErrorValue(q, e))
    // A data-table record is a pinned value source: its cache is audited, its text never parsed
    case CellValue.Formula(_, cached, _: FormulaKind.DataTable) => cacheFindings(q, cached)
    case CellValue.Formula(text, cached, _) =>
      val parsed = FormulaParser.parse(text) match
        case Left(err) => List(Finding.Unparseable(q, unparseableMessage(text, err)))
        case Right(expr) =>
          val volatile = if callsVolatile(expr) then List(Finding.Volatile(q)) else Nil
          val external = if TExpr.containsExternalRef(expr) then List(Finding.External(q)) else Nil
          volatile ++ external
      cacheFindings(q, cached) ++ parsed
    case _ => Nil

  /**
   * Cap on the formula text quoted into an unparseable entry — the same 80 characters `xl lint`'s
   * `formula-unparseable` finding quotes, so the two tools print one formula the same way.
   */
  private val UnparseableTextSample = 80

  /**
   * #676: `SUM(A1:A2: Unexpected end of formula at position 10` — the formula (capped at
   * [[UnparseableTextSample]] characters) and the parser's one-line reason, the form lint uses. The
   * multi-line caret rendering stays the parser's own (`ParseError.formatWithContext`): a report
   * listing many cells reads one line per cell, and an 8193-character formula is never echoed.
   */
  private def unparseableMessage(text: String, err: ParseError): String =
    val shown =
      if text.length > UnparseableTextSample then text.take(UnparseableTextSample) + "…" else text
    s"$shown: ${ParseError.describe(err)}"

  private def cacheFindings(q: QualifiedRef, cached: Option[CellValue]): List[Finding] =
    cached match
      case Some(CellValue.Error(e)) => List(Finding.ErrorValue(q, e))
      case Some(_) => Nil
      case None => List(Finding.Uncached(q))

  /** Whether the expression calls a function flagged `FunctionFlags.volatile`, at any depth. */
  private def callsVolatile(expr: TExpr[?]): Boolean = expr match
    case call: TExpr.Call[?] =>
      call.spec.flags.volatile ||
      call.spec.argSpec.toValues(call.args).exists {
        case ArgValue.Expr(e) => callsVolatile(e)
        case _ => false
      }
    // GH-680: a chain's left spine in one loop, not one recursion per operator
    case chain @ (_: TExpr.Add | _: TExpr.Sub | _: TExpr.Mul | _: TExpr.Div | _: TExpr.Pow |
        _: TExpr.Concat | _: TExpr.Eq[?] | _: TExpr.Neq[?] | _: TExpr.Lt[?] | _: TExpr.Lte[?] |
        _: TExpr.Gt[?] | _: TExpr.Gte[?]) =>
      BinarySpine.existsOperand(chain)(callsVolatile(_))
    case TExpr.ToInt(e) => callsVolatile(e)
    case TExpr.UnaryPlus(e) => callsVolatile(e)
    case TExpr.Percent(e) => callsVolatile(e)
    case TExpr.DateToSerial(e) => callsVolatile(e)
    case TExpr.DateTimeToSerial(e) => callsVolatile(e)
    case TExpr.Let(bindings, body) =>
      bindings.exists((_, value) => callsVolatile(value)) || callsVolatile(body)
    case TExpr.Coerced(inner, _) => callsVolatile(inner)
    case _ => false
