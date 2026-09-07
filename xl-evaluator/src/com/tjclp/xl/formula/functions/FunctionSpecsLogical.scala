package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, BindingCoercion}
import com.tjclp.xl.formula.eval.{EvalError, ArrayArithmetic, ArrayResult, ScalarCoercion}
import com.tjclp.xl.formula.Arity

import com.tjclp.xl.cells.{CellError, CellValue}
import scala.util.boundary
import boundary.break

trait FunctionSpecsLogical extends FunctionSpecsBase:

  /**
   * GH-338: evaluate one AND/OR argument to its boolean contribution — `None` when the argument
   * contributes nothing (a range with no logical values).
   *
   * Array-shaped arguments (range comparisons like A1:A10>0, array-returning calls, NOT over an
   * array) fold elementwise from `seed` with `combine` (AND = TRUE/&&, OR = FALSE/||) using the
   * broadcastIf condition conventions — GH-344: an error element propagates as its Excel error
   * VALUE, text elements refuse with a loud Left. Scalars follow the shared Excel-truthiness table
   * (numbers zero/non-zero, empty FALSE, error values propagate, text refuses).
   *
   * GH-564: bare ranges (`=AND(K9:K21)`, the house check-row idiom) follow Excel's REFERENCE rule
   * instead: logical cells count as themselves, numeric cells as zero/non-zero, text and blank
   * cells are ignored, and an error cell propagates. A range with nothing to fold contributes
   * `None`; when no argument contributes at all the function is #VALUE!, as in Excel.
   */
  private def conditionArg(fnName: String, ctx: EvalContext, expr: TExpr[?])(
    seed: Boolean,
    combine: (Boolean, Boolean) => Boolean
  ): Either[EvalError, Option[Boolean]] =
    val label = s"$fnName condition"
    expr match
      case _: TExpr.RangeRef | _: TExpr.SheetRange =>
        evalMaybeArrayArg(ctx, expr).flatMap {
          case arr: ArrayResult => rangeFold(label, arr, combine)
          case scalar => scalarCondition(label, scalar).map(Some(_))
        }
      case other =>
        evalMaybeArrayArg(ctx, other).flatMap {
          case arr: ArrayResult =>
            ArrayArithmetic.truthyElements(label, arr).map(v => Some(v.foldLeft(seed)(combine)))
          case scalar => scalarCondition(label, scalar).map(Some(_))
        }

  private def scalarCondition(label: String, scalar: Any): Either[EvalError, Boolean] =
    ScalarCoercion.coerce(label, scalar, BindingCoercion.Bool).flatMap {
      case b: Boolean => Right(b)
      case other => Left(EvalError.TypeMismatch(label, "boolean", other.toString))
    }

  /**
   * GH-564: fold the logical values of a range argument in row-major order with O(1) state (an
   * `=AND(A:A)` check row is a million cells; nothing is materialized beyond the ArrayResult the
   * range reader already produced). Booleans as-is, numbers zero/non-zero (Excel: "if an array or
   * reference argument contains text or empty cells, those values are ignored"); the first error
   * cell stops the walk and propagates as its Excel error VALUE (GH-344). `None` when no cell
   * contributed.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def rangeFold(
    label: String,
    arr: ArrayResult,
    combine: (Boolean, Boolean) => Boolean
  ): Either[EvalError, Option[Boolean]] =
    boundary:
      var acc: Option[Boolean] = None
      val rows = arr.values
      var r = 0
      while r < rows.length do
        val row = rows(r)
        var c = 0
        while c < row.length do
          rangeLogical(label, row(c)) match
            case Left(err) => break(Left(err))
            case Right(Some(b)) => acc = Some(acc.fold(b)(combine(_, b)))
            case Right(None) => ()
          c += 1
        r += 1
      Right(acc)

  /**
   * One range cell's logical contribution. The ArrayResult arrives with formula cells already
   * resolved to their values (GH-499's evaluating range reader), so the cached-formula arm is a
   * defensive unwrap; text, rich text and blanks contribute nothing.
   */
  private def rangeLogical(label: String, cv: CellValue): Either[EvalError, Option[Boolean]] =
    cv match
      case CellValue.Bool(b) => Right(Some(b))
      case CellValue.Number(n) => Right(Some(n.signum != 0))
      case CellValue.Error(err) => Left(EvalError.ErrorValue(err, Some(label)))
      case CellValue.Formula(_, Some(cached), _) => rangeLogical(label, cached)
      case _ => Right(None)

  /**
   * GH-344/GH-564: fold every argument left-to-right (Excel does NOT short-circuit: the first
   * failure wins), combining contributions with `combine`; #VALUE! when nothing contributed.
   */
  private def foldLogical(
    fnName: String,
    args: List[TExpr[?]],
    ctx: EvalContext,
    seed: Boolean,
    combine: (Boolean, Boolean) => Boolean
  ): Either[EvalError, Boolean] =
    args
      .foldLeft[Either[EvalError, Option[Boolean]]](Right(None)) {
        case (Left(err), _) => Left(err)
        case (Right(acc), arg) =>
          conditionArg(fnName, ctx, arg)(seed, combine).map {
            case Some(b) => Some(acc.fold(b)(combine(_, b)))
            case None => acc
          }
      }
      .flatMap {
        case Some(b) => Right(b)
        case None =>
          Left(EvalError.ErrorValue(CellError.Value, Some(s"$fnName: no logical values")))
      }

  val and: FunctionSpec[Boolean] { type Args = BooleanList } =
    FunctionSpec.simple[Boolean, BooleanList]("AND", Arity.atLeastOne) { (args, ctx) =>
      // GH-338: array arguments aggregate across every element (Excel AND is n-ary over
      // arrays). GH-344: Excel does NOT short-circuit logical functions — EVERY argument
      // evaluates left-to-right, the first failure (error value or refusal) wins, and only
      // then does the decisive fold apply: =AND(FALSE,1/0) is #DIV/0!, not FALSE. Known
      // residual: =AND(FALSE,"abc") is a loud Left (full #VALUE! arrives with the deferred
      // TypeMismatch boundary demotion follow-up). GH-564: bare ranges fold their logical cells.
      foldLogical("AND", args, ctx, seed = true, _ && _)
    }

  val or: FunctionSpec[Boolean] { type Args = BooleanList } =
    FunctionSpec.simple[Boolean, BooleanList]("OR", Arity.atLeastOne) { (args, ctx) =>
      // GH-338: array arguments aggregate across every element (Excel OR is n-ary over
      // arrays). GH-344: eager like AND — see the AND note; =OR(TRUE,#N/A-cell) is #N/A.
      // GH-564: bare ranges fold their logical cells.
      foldLogical("OR", args, ctx, seed = false, _ || _)
    }

  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  val not: FunctionSpec[Boolean] { type Args = UnaryBoolean } =
    FunctionSpec.simple[Boolean, UnaryBoolean]("NOT", Arity.one) { (expr, ctx) =>
      // GH-338: an array condition broadcasts elementwise (Excel NOT is elementwise, not an
      // aggregate); the ArrayResult travels through the erased Boolean slot exactly like
      // comparison results do — cast the Either container, not the value. Scalar conditions
      // keep Excel truthiness; bare ranges keep their pre-existing "must be used within a
      // function" error (NOT is elementwise, so there is no reference fold to apply).
      val evaluated = expr match
        case _: TExpr.RangeRef | _: TExpr.SheetRange => evalAny(ctx, expr)
        case other => evalMaybeArrayArg(ctx, other)
      evaluated.flatMap {
        case arr: ArrayResult =>
          ArrayArithmetic
            .broadcastNot("NOT condition", arr)
            .asInstanceOf[Either[EvalError, Boolean]]
        case scalar =>
          ScalarCoercion.coerce("NOT condition", scalar, BindingCoercion.Bool).flatMap {
            case b: Boolean => Right(!b)
            case other => Left(EvalError.TypeMismatch("NOT condition", "boolean", other.toString))
          }
      }
    }
