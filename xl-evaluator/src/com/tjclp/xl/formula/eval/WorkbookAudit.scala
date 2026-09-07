package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.{ArgValue, FunctionRegistry}
import com.tjclp.xl.formula.graph.{DependencyGraph, QualifiedGraph}
import com.tjclp.xl.formula.graph.DependencyGraph.{QualifiedRef, Scc}
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.workbooks.{CalcPr, Workbook}

/**
 * ADR-017 §2.10: every reason a number in the workbook can be wrong, bucketed in one pass — what
 * `xl audit` prints and `wb.audit` returns.
 *
 * Findings (they make [[isClean]] false):
 *   - `errorCells` — a cached Excel error on a formula, or a bare error value
 *   - `uncachedFormulas` — formulas with no cached value (`xl recalc` fills them)
 *   - `unparseable` — formulas this evaluator cannot parse, with the parser's diagnostic in context
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
  iterativeCycles: Vector[Scc]
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
      iterativeCycles = iterativeCycles.filter(_.members.exists(_.sheet == sheet))
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

  def of(wb: Workbook): WorkbookAudit =
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
      iterativeCycles = if iterative then cyclic else Vector.empty
    )

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
        case Left(err) => List(Finding.Unparseable(q, ParseError.formatWithContext(err, text)))
        case Right(expr) =>
          val volatile = if callsVolatile(expr) then List(Finding.Volatile(q)) else Nil
          val external = if TExpr.containsExternalRef(expr) then List(Finding.External(q)) else Nil
          volatile ++ external
      cacheFindings(q, cached) ++ parsed
    case _ => Nil

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
    case TExpr.Add(l, r) => callsVolatile(l) || callsVolatile(r)
    case TExpr.Sub(l, r) => callsVolatile(l) || callsVolatile(r)
    case TExpr.Mul(l, r) => callsVolatile(l) || callsVolatile(r)
    case TExpr.Div(l, r) => callsVolatile(l) || callsVolatile(r)
    case TExpr.Pow(l, r) => callsVolatile(l) || callsVolatile(r)
    case TExpr.Concat(l, r) => callsVolatile(l) || callsVolatile(r)
    case TExpr.Eq(l, r) => callsVolatile(l) || callsVolatile(r)
    case TExpr.Neq(l, r) => callsVolatile(l) || callsVolatile(r)
    case TExpr.Lt(l, r) => callsVolatile(l) || callsVolatile(r)
    case TExpr.Lte(l, r) => callsVolatile(l) || callsVolatile(r)
    case TExpr.Gt(l, r) => callsVolatile(l) || callsVolatile(r)
    case TExpr.Gte(l, r) => callsVolatile(l) || callsVolatile(r)
    case TExpr.ToInt(e) => callsVolatile(e)
    case TExpr.UnaryPlus(e) => callsVolatile(e)
    case TExpr.Percent(e) => callsVolatile(e)
    case TExpr.DateToSerial(e) => callsVolatile(e)
    case TExpr.DateTimeToSerial(e) => callsVolatile(e)
    case TExpr.Let(bindings, body) =>
      bindings.exists((_, value) => callsVolatile(value)) || callsVolatile(body)
    case TExpr.Coerced(inner, _) => callsVolatile(inner)
    case _ => false
