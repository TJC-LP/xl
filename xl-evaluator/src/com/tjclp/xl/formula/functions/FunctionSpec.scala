package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.CellRange
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * Metadata flags for function behavior that impacts formatting or inference.
 */
final case class FunctionFlags(
  returnsDate: Boolean = false,
  returnsTime: Boolean = false
)

final case class ArgPrinter(
  expr: TExpr[?] => String,
  location: TExpr.RangeLocation => String,
  cellRange: CellRange => String
)

final case class EvalContext(
  sheet: Sheet,
  clock: Clock,
  workbook: Option[Workbook],
  evalExpr: [A] => TExpr[A] => Either[EvalError, A],
  /** Current cell being evaluated. Used by ROW() and COLUMN() with no arguments. */
  currentCell: Option[ARef] = None,
  /** Recursion depth for cross-sheet formula evaluation. */
  depth: Int = 0
)

sealed trait ArgValue
object ArgValue:
  final case class Expr(value: TExpr[?]) extends ArgValue
  final case class Range(value: TExpr.RangeLocation) extends ArgValue
  final case class Cells(value: CellRange) extends ArgValue

/**
 * Describes how to parse, render, and transform a function's arguments.
 */
trait ArgSpec[A]:
  def describeParts: List[String]
  final def describe: String = describeParts.mkString(", ")

  def parse(
    args: List[TExpr[?]],
    pos: Int,
    fnName: String
  ): Either[ParseError, (A, List[TExpr[?]])]

  def toValues(args: A): List[ArgValue]

  def map(
    args: A
  )(
    mapExpr: TExpr[?] => TExpr[?],
    mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
    mapCells: CellRange => CellRange
  ): A

  def render(args: A, printer: ArgPrinter): List[String] =
    toValues(args).map {
      case ArgValue.Expr(expr) => printer.expr(expr)
      case ArgValue.Range(location) => printer.location(location)
      case ArgValue.Cells(range) => printer.cellRange(range)
    }

trait FunctionSpec[A]:
  type Args
  def name: String
  def arity: Arity
  def argSpec: ArgSpec[Args]
  def eval(args: Args, ctx: EvalContext): Either[EvalError, A]
  def flags: FunctionFlags = FunctionFlags()

  def render(args: Args, printer: ArgPrinter): String =
    val rendered = argSpec.render(args, printer)
    s"${name}(${rendered.mkString(", ")})"

object FunctionSpec:
  final case class Simple[A, A0](
    name: String,
    arity: Arity,
    argSpec: ArgSpec[A0],
    evalFn: (A0, EvalContext) => Either[EvalError, A],
    override val flags: FunctionFlags = FunctionFlags(),
    renderFn: Option[(A0, ArgPrinter) => String] = None
  ) extends FunctionSpec[A]:
    type Args = A0
    def eval(args: A0, ctx: EvalContext): Either[EvalError, A] = evalFn(args, ctx)
    override def render(args: A0, printer: ArgPrinter): String =
      renderFn.map(_(args, printer)).getOrElse(super.render(args, printer))

  def simple[A, A0](
    name: String,
    arity: Arity,
    flags: FunctionFlags = FunctionFlags(),
    renderFn: Option[(A0, ArgPrinter) => String] = None
  )(evalFn: (A0, EvalContext) => Either[EvalError, A])(using
    spec: ArgSpec[A0]
  ): FunctionSpec[A] { type Args = A0 } =
    Simple(name, arity, spec, evalFn, flags, renderFn)

trait ExprCoercer[A]:
  def label: String
  def coerce(expr: TExpr[?]): TExpr[A]

object ExprCoercer:
  given numeric: ExprCoercer[BigDecimal] with
    val label = "number"
    def coerce(expr: TExpr[?]): TExpr[BigDecimal] = TExpr.asNumericExpr(expr)

  given boolean: ExprCoercer[Boolean] with
    val label = "boolean"
    def coerce(expr: TExpr[?]): TExpr[Boolean] = TExpr.asBooleanExpr(expr)

  given text: ExprCoercer[String] with
    val label = "text"
    def coerce(expr: TExpr[?]): TExpr[String] = TExpr.asStringExpr(expr)

  given intExpr: ExprCoercer[Int] with
    val label = "integer"
    def coerce(expr: TExpr[?]): TExpr[Int] = TExpr.asIntExpr(expr)

  given cellValue: ExprCoercer[CellValue] with
    val label = "cell"
    def coerce(expr: TExpr[?]): TExpr[CellValue] = TExpr.asCellValueExpr(expr)

  given dateExpr: ExprCoercer[java.time.LocalDate] with
    val label = "date"
    def coerce(expr: TExpr[?]): TExpr[java.time.LocalDate] = TExpr.asDateExpr(expr)

@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
object ArgSpec:
  def expr[A](using coercer: ExprCoercer[A]): ArgSpec[TExpr[A]] =
    new ArgSpec[TExpr[A]]:
      def describeParts: List[String] = List(coercer.label)

      def parse(
        args: List[TExpr[?]],
        pos: Int,
        fnName: String
      ): Either[ParseError, (TExpr[A], List[TExpr[?]])] =
        args match
          case head :: tail => Right((coercer.coerce(head), tail))
          case Nil =>
            Left(ParseError.InvalidArguments(fnName, pos, describe, "0 arguments"))

      def toValues(args: TExpr[A]): List[ArgValue] =
        List(ArgValue.Expr(args))

      def map(
        args: TExpr[A]
      )(
        mapExpr: TExpr[?] => TExpr[?],
        mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
        mapCells: CellRange => CellRange
      ): TExpr[A] =
        mapExpr(args).asInstanceOf[TExpr[A]]

  given rangeLocation: ArgSpec[TExpr.RangeLocation] with
    def describeParts: List[String] = List("range")

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (TExpr.RangeLocation, List[TExpr[?]])] =
      args match
        case TExpr.RangeRef(range) :: tail =>
          Right((TExpr.RangeLocation.Local(range), tail))
        case TExpr.SheetRange(sheet, range) :: tail =>
          Right((TExpr.RangeLocation.CrossSheet(sheet, range), tail))
        case _ =>
          Left(ParseError.InvalidArguments(fnName, pos, describe, s"${args.length} arguments"))

    def toValues(args: TExpr.RangeLocation): List[ArgValue] =
      List(ArgValue.Range(args))

    def map(
      args: TExpr.RangeLocation
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): TExpr.RangeLocation =
      mapRange(args)

  given cellRange: ArgSpec[CellRange] with
    def describeParts: List[String] = List("range")

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (CellRange, List[TExpr[?]])] =
      args match
        case TExpr.RangeRef(range) :: tail =>
          Right((range, tail))
        case _ =>
          Left(ParseError.InvalidArguments(fnName, pos, describe, s"${args.length} arguments"))

    def toValues(args: CellRange): List[ArgValue] =
      List(ArgValue.Cells(args))

    def map(
      args: CellRange
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): CellRange =
      mapCells(args)

  given option[A](using inner: ArgSpec[A]): ArgSpec[Option[A]] with
    def describeParts: List[String] = List(s"optional ${inner.describe}")

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (Option[A], List[TExpr[?]])] =
      args match
        case Nil => Right((None, Nil))
        case _ =>
          inner.parse(args, pos, fnName).map { case (value, rest) => (Some(value), rest) }

    def toValues(args: Option[A]): List[ArgValue] =
      args.toList.flatMap(inner.toValues)

    def map(
      args: Option[A]
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): Option[A] =
      args.map(inner.map(_)(mapExpr, mapRange, mapCells))

  given list[A](using inner: ArgSpec[A]): ArgSpec[List[A]] with
    def describeParts: List[String] = List(s"${inner.describe}...")

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (List[A], List[TExpr[?]])] =
      @annotation.tailrec
      def loop(
        remaining: List[TExpr[?]],
        acc: List[A]
      ): Either[ParseError, (List[A], List[TExpr[?]])] =
        remaining match
          case Nil => Right((acc.reverse, Nil))
          case _ =>
            inner.parse(remaining, pos, fnName) match
              case Right((value, rest)) => loop(rest, value :: acc)
              case Left(err) => Left(err)
      loop(args, Nil)

    def toValues(args: List[A]): List[ArgValue] =
      args.flatMap(inner.toValues)

    def map(
      args: List[A]
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): List[A] =
      args.map(inner.map(_)(mapExpr, mapRange, mapCells))

  given emptyTuple: ArgSpec[EmptyTuple] with
    def describeParts: List[String] = Nil

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (EmptyTuple, List[TExpr[?]])] =
      Right((EmptyTuple, args))

    def toValues(args: EmptyTuple): List[ArgValue] = Nil

    def map(
      args: EmptyTuple
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): EmptyTuple =
      args

  given tuple[H, T <: Tuple](using head: ArgSpec[H], tail: ArgSpec[T]): ArgSpec[H *: T] with
    def describeParts: List[String] =
      head.describeParts ++ tail.describeParts

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (H *: T, List[TExpr[?]])] =
      for
        (h, rest) <- head.parse(args, pos, fnName)
        (t, rest2) <- tail.parse(rest, pos, fnName)
      yield (h *: t, rest2)

    def toValues(args: H *: T): List[ArgValue] =
      head.toValues(args.head) ++ tail.toValues(args.tail)

    def map(
      args: H *: T
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): H *: T =
      head.map(args.head)(mapExpr, mapRange, mapCells) *: tail
        .map(args.tail)(mapExpr, mapRange, mapCells)

  /** Variadic numeric arg: either a range location or a numeric expression. */
  type NumericArg = Either[TExpr.RangeLocation, TExpr[BigDecimal]]

  given numericArg: ArgSpec[NumericArg] with
    def describeParts: List[String] = List("number or range")

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (NumericArg, List[TExpr[?]])] =
      args match
        case TExpr.RangeRef(range) :: tail =>
          Right((Left(TExpr.RangeLocation.Local(range)), tail))
        case TExpr.SheetRange(sheet, range) :: tail =>
          Right((Left(TExpr.RangeLocation.CrossSheet(sheet, range)), tail))
        case head :: tail =>
          Right((Right(TExpr.asNumericExpr(head)), tail))
        case Nil =>
          Left(ParseError.InvalidArguments(fnName, pos, describe, "0 arguments"))

    def toValues(args: NumericArg): List[ArgValue] =
      args match
        case Left(loc) => List(ArgValue.Range(loc))
        case Right(expr) => List(ArgValue.Expr(expr))

    def map(
      args: NumericArg
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): NumericArg =
      args match
        case Left(loc) => Left(mapRange(loc))
        case Right(expr) => Right(mapExpr(expr).asInstanceOf[TExpr[BigDecimal]])
