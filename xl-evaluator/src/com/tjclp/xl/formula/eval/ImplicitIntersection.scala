package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.formula.{Clock, Rng}
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
   * The cells of `cells` that still hold a plain formula in `wb` and diverge, in sheet then
   * row-major order.
   */
  def check(wb: Workbook, cells: Vector[(SheetName, ARef)], clock: Clock): Vector[Divergence] =
    val sheetOrder = wb.sheets.map(_.name).zipWithIndex.toMap
    cells.distinct
      .flatMap { (name, ref) => sheetOrder.get(name).map(i => (i, name, ref)) }
      .sortBy((i, _, ref) => (i, ref.row.index0, ref.col.index0))
      .flatMap((_, name, ref) => divergenceAt(wb, name, ref, clock))

  private def divergenceAt(
    wb: Workbook,
    name: SheetName,
    ref: ARef,
    clock: Clock
  ): Option[Divergence] =
    wb.sheets.find(_.name == name).flatMap { sheet =>
      sheet.cells.get(ref).map(_.value) match
        case Some(CellValue.Formula(text, _, _: FormulaKind.Normal)) if mayIntersect(text) =>
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
                  Divergence(name, ref, text, p, topLeft, rows, cols)
                )
              }
            case _ => None
        case _ => None
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

  private val quoted = "\"(?:[^\"]|\"\")*\"".r
  private val identifier = """\$?[A-Za-z_\\][A-Za-z0-9_.$]*""".r
  private val cellRef = """\$?[A-Za-z]{1,3}\$?[0-9]+""".r

  /**
   * A cheap, conservative prefilter: a formula with no range (`:`), array literal, function call,
   * whole-axis or structured reference, and no identifier but cell references and booleans cannot
   * intersect anything, so it is not evaluated twice. Anything else is (a name may be a range).
   */
  private def mayIntersect(text: String): Boolean =
    val code = quoted.replaceAllIn(text, "\"\"")
    code.exists(c => c == ':' || c == '{' || c == '(' || c == '[' || c == '#') ||
    identifier.findAllIn(code).exists { id =>
      !cellRef.matches(id) && !id.equalsIgnoreCase("TRUE") && !id.equalsIgnoreCase("FALSE")
    }
