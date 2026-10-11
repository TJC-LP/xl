package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.formula.ast.{BinarySpine, TExpr}
import com.tjclp.xl.formula.functions.{ArgValue, FunctionRegistry, ReferenceOperators}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-714: plain formula cells whose value depends on implicit intersection.
 *
 * Excel 365 evaluates a plain (`Normal`-kind) formula cell as a legacy formula: a multi-cell
 * reference in a value position is intersected with the cell's row or column. Typed into Excel 365
 * the same text is an array formula, so an agent writing `=SUM(A1:A10*B1:B10)` in row 5 or
 * `=A1:A10*2` usually meant the array. [[check]] evaluates each cell both ways — as the plain cell
 * and as an array at the cell — and reports where they differ: the top-left value, or the shape (a
 * spill the plain cell cannot show). Informational: a caller warns, it never gates.
 *
 * Both evaluations share one clock and a fixed-seed randomness source, so a volatile formula
 * (`RAND()`, `NOW()`) compares equal to itself, while an array generator (`SEQUENCE(3)`) still
 * differs by shape. A host failure on either side is no finding. Pure and total.
 *
 * Cost: only a cell whose parsed formula [[mayDiverge]] is evaluated, and at most `budget` of those
 * (the first in reading order), so a long drag of `=SUM($A$1:A1)` costs a parse per cell.
 */
private[xl] object ImplicitIntersection:

  /** One cell whose plain value differs from its array value (`rows`x`cols`, top-left shown). */
  final case class Divergence(
    sheet: SheetName,
    cell: ARef,
    formula: String,
    plain: CellValue,
    arrayTopLeft: CellValue,
    rows: Int,
    cols: Int
  ) derives CanEqual:
    def isShape: Boolean = rows != 1 || cols != 1

  /**
   * What [[check]] found: the diverging cells among the first `checked` of the `candidates` cells
   * whose formula [[mayDiverge]]; `sampled` when the budget stopped it before the last.
   */
  final case class Report(hits: Vector[Divergence], checked: Int, candidates: Int) derives CanEqual:
    def sampled: Boolean = checked < candidates

  /** How many candidate cells one check evaluates both ways at most. */
  val DefaultBudget: Int = 1000

  /**
   * The cells of `cells` that still hold a plain formula in `wb` and diverge, in sheet then
   * row-major order, checking at most `budget` candidates.
   */
  def check(
    wb: Workbook,
    cells: Vector[(SheetName, ARef)],
    clock: Clock,
    budget: Int = DefaultBudget
  ): Report =
    val sheets = wb.sheets.map(s => s.name -> s).toMap
    val sheetOrder = wb.sheets.map(_.name).zipWithIndex.toMap
    val candidates = cells.distinct
      .flatMap { (name, ref) => sheetOrder.get(name).map(i => (i, name, ref)) }
      .sortBy((i, _, ref) => (i, ref.row.index0, ref.col.index0))
      .flatMap { (_, name, ref) =>
        sheets.get(name).flatMap { sheet =>
          sheet.cells.get(ref).map(_.value) match
            case Some(CellValue.Formula(text, _, _: FormulaKind.Normal))
                if FormulaParser.parse(s"=$text").exists(mayDiverge) =>
              Some((sheet, ref, text))
            case _ => None
        }
      }
    val checked = candidates.take(math.max(budget, 0))
    Report(
      checked.flatMap((sheet, ref, text) => divergenceAt(wb, sheet, ref, text, clock)),
      checked.size,
      candidates.size
    )

  private def divergenceAt(
    wb: Workbook,
    sheet: Sheet,
    ref: ARef,
    text: String,
    clock: Clock
  ): Option[Divergence] =
    val full = s"=$text"
    val plain = SheetEvaluator.evaluatePlainAt(
      sheet,
      full,
      ref,
      Evaluator.instance(Rng.seeded(0)),
      clock,
      Some(wb)
    )
    val array = SheetEvaluator.evaluateArrayValues(
      sheet,
      full,
      ref,
      Evaluator.instance(Rng.seeded(0)),
      clock,
      Some(wb)
    )
    (plain, array) match
      case (Right(p), Right(grid)) =>
        val rows = grid.size
        val cols = grid.map(_.size).maxOption.getOrElse(0)
        grid.headOption.flatMap(_.headOption).flatMap { topLeft =>
          val shape = rows != 1 || cols != 1
          Option.when(shape || !sameValue(p, topLeft))(
            Divergence(sheet.name, ref, text, p, topLeft, rows, cols)
          )
        }
      case _ => None

  /**
   * Whether a parsed formula can evaluate differently as a plain cell than as an array: an array
   * source — a multi-cell reference, a name (it may be one), an array constant, a call whose
   * declared result is an array ([[FunctionRegistry.arrayResultNames]]) or a reference operator —
   * in a position that intersects or broadcasts: an operand, a scalar argument, the result. A range
   * slot (`ArgValue.Range`: SUM's and COUNTIF's ranges, VLOOKUP's table, SUMPRODUCT's arrays) takes
   * its reference whole in both modes, so `SUM($A$1:A1)` is no candidate. Conservative otherwise: a
   * LET binding or an argument expression holding a source counts.
   */
  def mayDiverge(expr: TExpr[?]): Boolean = expr match
    case _: TExpr.RangeRef | _: TExpr.SheetRange | _: TExpr.ExternalRange | _: TExpr.NameRef |
        _: TExpr.SheetNameRef =>
      true
    case TExpr.Lit(value) =>
      value match
        case _: ArrayResult => true
        case _ => false
    case _: TExpr.Ref[?] | _: TExpr.PolyRef | _: TExpr.SheetRef[?] | _: TExpr.SheetPolyRef |
        _: TExpr.ExternalRef | _: TExpr.ErrorLit | TExpr.Missing | _: TExpr.Aggregate |
        _: TExpr.BindingRef | _: TExpr.CoercedBindingRef[?] =>
      false
    case chain @ (_: TExpr.Add | _: TExpr.Sub | _: TExpr.Mul | _: TExpr.Div | _: TExpr.Pow |
        _: TExpr.Concat | _: TExpr.Eq[?] | _: TExpr.Neq[?] | _: TExpr.Lt[?] | _: TExpr.Lte[?] |
        _: TExpr.Gt[?] | _: TExpr.Gte[?]) =>
      BinarySpine.operandList(chain).exists(mayDiverge)
    case TExpr.UnaryPlus(inner) => mayDiverge(inner)
    case TExpr.Percent(inner) => mayDiverge(inner)
    case TExpr.ToInt(inner) => mayDiverge(inner)
    case TExpr.DateToSerial(inner) => mayDiverge(inner)
    case TExpr.DateTimeToSerial(inner) => mayDiverge(inner)
    case TExpr.Coerced(inner, _) => mayDiverge(inner)
    case TExpr.Let(bindings, body) =>
      bindings.exists((_, value) => mayDiverge(value)) || mayDiverge(body)
    case call: TExpr.Call[?] =>
      FunctionRegistry.arrayResultNames.contains(call.spec.name.toUpperCase) ||
      ReferenceOperators.isOperatorCall(call) ||
      call.spec.argSpec.toValues(call.args).exists {
        case ArgValue.Expr(arg) => mayDiverge(arg)
        case _: ArgValue.Range | _: ArgValue.Cells => false
      }

  /** Numbers compare by value, Empty reads 0, anything else by equality (errors by code). */
  private def sameValue(a: CellValue, b: CellValue): Boolean =
    (normalized(a), normalized(b)) match
      case (CellValue.Number(x), CellValue.Number(y)) => x.compare(y) == 0
      case (x, y) => x == y

  private def normalized(v: CellValue): CellValue = v match
    case CellValue.Empty => CellValue.Number(BigDecimal(0))
    case CellValue.Formula(_, Some(cached), _) => normalized(cached)
    case other => other
