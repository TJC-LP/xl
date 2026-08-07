package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.cells.{CellError, CellValue}

trait FunctionSpecsTypeCheck extends FunctionSpecsBase:
  val iferror: FunctionSpec[CellValue] { type Args = IfErrorArgs } =
    FunctionSpec.simple[CellValue, IfErrorArgs]("IFERROR", Arity.two) { (args, ctx) =>
      val (valueExpr, valueIfErrorExpr) = args
      evalValue(ctx, valueExpr) match
        case Left(_) =>
          evalValue(ctx, valueIfErrorExpr).map(toCellValue)
        case Right(ExprValue.Cell(cv)) =>
          cv match
            case CellValue.Error(_) =>
              evalValue(ctx, valueIfErrorExpr).map(toCellValue)
            case _ => Right(cv)
        case Right(other) =>
          Right(toCellValue(other))
    }

  val iserror: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISERROR", Arity.one) { (expr, ctx) =>
      evalValue(ctx, expr) match
        case Left(_) => Right(true)
        case Right(ExprValue.Cell(CellValue.Error(_))) => Right(true)
        case Right(_) => Right(false)
    }

  val iserr: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISERR", Arity.one) { (expr, ctx) =>
      evalValue(ctx, expr) match
        // GH-344: Excel's ISERR excludes #N/A on the Left channel too (`=ISERR(1<na-cell)` is
        // FALSE); every other failure — error value or host — stays TRUE like ISERROR.
        case Left(EvalError.ErrorValue(CellError.NA, _)) => Right(false)
        case Left(_) => Right(true)
        case Right(ExprValue.Cell(CellValue.Error(err))) => Right(err != CellError.NA)
        case Right(_) => Right(false)
    }

  val isnumber: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISNUMBER", Arity.one) { (expr, ctx) =>
      evalValue(ctx, expr) match
        case Left(_) => Right(false)
        case Right(ExprValue.Cell(CellValue.Number(_))) => Right(true)
        case Right(ExprValue.Cell(CellValue.Formula(_, Some(CellValue.Number(_)), _))) =>
          Right(true)
        case Right(ExprValue.Number(_)) => Right(true)
        case Right(_) => Right(false)
    }

  val istext: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISTEXT", Arity.one) { (expr, ctx) =>
      evalValue(ctx, expr) match
        case Left(_) => Right(false)
        case Right(ExprValue.Cell(CellValue.Text(_))) => Right(true)
        case Right(ExprValue.Cell(CellValue.Formula(_, Some(CellValue.Text(_)), _))) =>
          Right(true)
        case Right(ExprValue.Text(_)) => Right(true)
        case Right(_) => Right(false)
    }

  val isblank: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISBLANK", Arity.one) { (expr, ctx) =>
      evalValue(ctx, expr) match
        case Left(_) => Right(false)
        case Right(ExprValue.Cell(CellValue.Empty)) => Right(true)
        case Right(_) => Right(false)
    }

  /**
   * GH-476: N(value) — Excel's information-family coercion, absent from the roster and therefore an
   * UnknownFunction host error on ordinary banker files.
   *
   * Excel's table: numbers pass through, dates become their serial, TRUE/FALSE become 1/0, text and
   * blanks become 0, and an error value propagates. References and formula caches carry errors on
   * the value channel, so N lifts them onto the absorbing Left channel here.
   */
  val n: FunctionSpec[BigDecimal] { type Args = UnaryCellValue } =
    FunctionSpec.simple[BigDecimal, UnaryCellValue](
      "N",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (expr, ctx) =>
      evalValue(ctx, expr).flatMap(numericCoercionN)
    }

  private def numericCoercionN(value: ExprValue): Either[EvalError, BigDecimal] =
    value match
      case ExprValue.Number(x) => Right(x)
      case ExprValue.Bool(b) => Right(if b then BigDecimal(1) else BigDecimal(0))
      case ExprValue.Date(d) =>
        Right(BigDecimal(CellValue.dateTimeToExcelSerial(d.atStartOfDay())))
      case ExprValue.DateTime(dt) => Right(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))
      case ExprValue.Text(_) => Right(BigDecimal(0))
      case ExprValue.Opaque(_) => Right(BigDecimal(0))
      case ExprValue.Cell(cv) => numericCoercionNCell(cv)

  private def numericCoercionNCell(cv: CellValue): Either[EvalError, BigDecimal] =
    cv match
      case CellValue.Number(x) => Right(x)
      case CellValue.Bool(b) => Right(if b then BigDecimal(1) else BigDecimal(0))
      case CellValue.DateTime(dt) => Right(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))
      case CellValue.Formula(_, Some(cached), _) => numericCoercionNCell(cached)
      case CellValue.Error(err) => Left(EvalError.ErrorValue(err, Some("N")))
      case _ => Right(BigDecimal(0))
