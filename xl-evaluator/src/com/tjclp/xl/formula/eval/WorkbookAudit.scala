package com.tjclp.xl.formula.eval

import java.util.Locale

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.ArgValue
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
 *   - `cycles` — the cyclic strongly connected components of the formula graph
 *   - `unresolvedReaders` — formulas whose name references the static graph cannot resolve
 *     (GH-507); an unparseable formula is reported once, in `unparseable`
 *
 * Notes (reported, never findings): `volatile` (a call to TODAY/NOW/RAND/RANDBETWEEN — the registry
 * has no volatility flag yet, so these are matched by name), `dynamic` (INDIRECT/OFFSET readers,
 * [[DependencyGraph.dynamicCells]]), `externalRefs` (formulas touching another workbook, whose
 * caches are pinned) and the file's `calcPr`.
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
  calcPr: Option[CalcPr]
) derives CanEqual:

  /** The number of findings: error cells, uncached and unparseable formulas, cycles, unresolved. */
  def findings: Int =
    errorCells.size + uncachedFormulas.size + unparseable.size + cycles.size +
      unresolvedReaders.size

  /** No findings. Volatile, dynamic and external references and `calcPr` do not count. */
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
      unresolvedReaders = unresolvedReaders.filter(_.sheet == sheet)
    )

object WorkbookAudit:

  /** Matched by (case-insensitive) function name: there is no `FunctionFlags.volatile` yet. */
  val volatileFunctions: Set[String] = Set("TODAY", "NOW", "RAND", "RANDBETWEEN")

  val clean: WorkbookAudit = WorkbookAudit(
    errorCells = Vector.empty,
    uncachedFormulas = Vector.empty,
    unparseable = Vector.empty,
    volatile = Vector.empty,
    dynamic = Vector.empty,
    cycles = Vector.empty,
    externalRefs = Vector.empty,
    unresolvedReaders = Vector.empty,
    calcPr = None
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

    WorkbookAudit(
      errorCells = scanned.collect { case Finding.ErrorValue(q, e) => (q, e) }.sortBy(_._1),
      uncachedFormulas = scanned.collect { case Finding.Uncached(q) => q }.sorted,
      unparseable = unparseable.sortBy(_._1),
      volatile = scanned.collect { case Finding.Volatile(q) => q }.sorted,
      dynamic = DependencyGraph.dynamicCells(wb).toVector.sorted,
      // sccs come dependency-first; the report reads in workbook order like every other bucket
      cycles = graph.sccs.filter(_.cyclic).sortBy(scc => scc.members.sorted.headOption),
      externalRefs = scanned.collect { case Finding.External(q) => q }.sorted,
      unresolvedReaders =
        DependencyGraph.unresolvedReaders(wb).iterator.filterNot(unparseableRefs).toVector.sorted,
      calcPr = wb.metadata.calcPr
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
          val volatile =
            if callsAny(expr, volatileFunctions) then List(Finding.Volatile(q)) else Nil
          val external = if TExpr.containsExternalRef(expr) then List(Finding.External(q)) else Nil
          volatile ++ external
      cacheFindings(q, cached) ++ parsed
    case _ => Nil

  private def cacheFindings(q: QualifiedRef, cached: Option[CellValue]): List[Finding] =
    cached match
      case Some(CellValue.Error(e)) => List(Finding.ErrorValue(q, e))
      case Some(_) => Nil
      case None => List(Finding.Uncached(q))

  /** Whether the expression calls any of `names` (upper-case), at any depth. */
  private def callsAny(expr: TExpr[?], names: Set[String]): Boolean = expr match
    case call: TExpr.Call[?] =>
      names.contains(call.spec.name.toUpperCase(Locale.ROOT)) ||
      call.spec.argSpec.toValues(call.args).exists {
        case ArgValue.Expr(e) => callsAny(e, names)
        case _ => false
      }
    case TExpr.Add(l, r) => callsAny(l, names) || callsAny(r, names)
    case TExpr.Sub(l, r) => callsAny(l, names) || callsAny(r, names)
    case TExpr.Mul(l, r) => callsAny(l, names) || callsAny(r, names)
    case TExpr.Div(l, r) => callsAny(l, names) || callsAny(r, names)
    case TExpr.Pow(l, r) => callsAny(l, names) || callsAny(r, names)
    case TExpr.Concat(l, r) => callsAny(l, names) || callsAny(r, names)
    case TExpr.Eq(l, r) => callsAny(l, names) || callsAny(r, names)
    case TExpr.Neq(l, r) => callsAny(l, names) || callsAny(r, names)
    case TExpr.Lt(l, r) => callsAny(l, names) || callsAny(r, names)
    case TExpr.Lte(l, r) => callsAny(l, names) || callsAny(r, names)
    case TExpr.Gt(l, r) => callsAny(l, names) || callsAny(r, names)
    case TExpr.Gte(l, r) => callsAny(l, names) || callsAny(r, names)
    case TExpr.ToInt(e) => callsAny(e, names)
    case TExpr.UnaryPlus(e) => callsAny(e, names)
    case TExpr.Percent(e) => callsAny(e, names)
    case TExpr.DateToSerial(e) => callsAny(e, names)
    case TExpr.DateTimeToSerial(e) => callsAny(e, names)
    case TExpr.Let(bindings, body) =>
      bindings.exists((_, value) => callsAny(value, names)) || callsAny(body, names)
    case TExpr.Coerced(inner, _) => callsAny(inner, names)
    case _ => false
