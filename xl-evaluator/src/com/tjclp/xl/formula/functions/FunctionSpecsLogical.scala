package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, BindingCoercion}
import com.tjclp.xl.formula.eval.{EvalError, ArrayArithmetic, ArrayResult, ScalarCoercion}
import com.tjclp.xl.formula.Arity

trait FunctionSpecsLogical extends FunctionSpecsBase:

  /**
   * GH-338: evaluate one AND/OR argument to its boolean contribution.
   *
   * Array-shaped arguments (range comparisons like A1:A10>0, array-returning calls, NOT over an
   * array) fold elementwise with `foldElems` (AND = forall, OR = exists) using the broadcastIf
   * condition conventions — text or error elements refuse with a clean Left. Scalars follow the
   * shared Excel-truthiness table (numbers zero/non-zero, empty FALSE, text refuses). Bare ranges
   * keep their pre-existing "must be used within a function" error: raw-range coercion parity
   * (Excel skips text/blanks over untyped ranges) is documented out of scope in GH-338.
   */
  private def conditionArg(fnName: String, ctx: EvalContext, expr: TExpr[?])(
    foldElems: Vector[Boolean] => Boolean
  ): Either[EvalError, Boolean] =
    val label = s"$fnName condition"
    val evaluated = expr match
      case _: TExpr.RangeRef | _: TExpr.SheetRange => evalAny(ctx, expr)
      case other => evalMaybeArrayArg(ctx, other)
    evaluated.flatMap {
      case arr: ArrayResult => ArrayArithmetic.truthyElements(label, arr).map(foldElems)
      case scalar =>
        ScalarCoercion.coerce(label, scalar, BindingCoercion.Bool).flatMap {
          case b: Boolean => Right(b)
          case other => Left(EvalError.TypeMismatch(label, "boolean", other.toString))
        }
    }

  val and: FunctionSpec[Boolean] { type Args = BooleanList } =
    FunctionSpec.simple[Boolean, BooleanList]("AND", Arity.atLeastOne) { (args, ctx) =>
      // GH-338: array arguments aggregate across every element (Excel AND is n-ary over
      // arrays); a decisive FALSE still short-circuits the remaining arguments.
      @annotation.tailrec
      def loop(remaining: List[TExpr[Boolean]]): Either[EvalError, Boolean] =
        remaining match
          case Nil => Right(true)
          case head :: tail =>
            conditionArg("AND", ctx, head)(_.forall(identity)) match
              case Left(err) => Left(err)
              case Right(false) => Right(false)
              case Right(true) => loop(tail)
      loop(args)
    }

  val or: FunctionSpec[Boolean] { type Args = BooleanList } =
    FunctionSpec.simple[Boolean, BooleanList]("OR", Arity.atLeastOne) { (args, ctx) =>
      // GH-338: array arguments aggregate across every element (Excel OR is n-ary over
      // arrays); a decisive TRUE still short-circuits the remaining arguments.
      @annotation.tailrec
      def loop(remaining: List[TExpr[Boolean]]): Either[EvalError, Boolean] =
        remaining match
          case Nil => Right(false)
          case head :: tail =>
            conditionArg("OR", ctx, head)(_.exists(identity)) match
              case Left(err) => Left(err)
              case Right(true) => Right(true)
              case Right(false) => loop(tail)
      loop(args)
    }

  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  val not: FunctionSpec[Boolean] { type Args = UnaryBoolean } =
    FunctionSpec.simple[Boolean, UnaryBoolean]("NOT", Arity.one) { (expr, ctx) =>
      // GH-338: an array condition broadcasts elementwise (Excel NOT is elementwise, not an
      // aggregate); the ArrayResult travels through the erased Boolean slot exactly like
      // comparison results do — cast the Either container, not the value. Scalar conditions
      // keep Excel truthiness; bare ranges keep their pre-existing error (see conditionArg).
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
