package com.tjclp.xl.formula

import com.tjclp.xl.cells.{CellError, CellValue}

trait FunctionSpecsTypeCheck extends FunctionSpecsBase:
  val iferror: FunctionSpec[CellValue] { type Args = IfErrorArgs } =
    FunctionSpec.simple[CellValue, IfErrorArgs]("IFERROR", Arity.two) { (args, ctx) =>
      val (valueExpr, valueIfErrorExpr) = args
      evalAny(ctx, valueExpr) match
        case Left(_) =>
          evalAny(ctx, valueIfErrorExpr).map(toCellValue)
        case Right(cv: CellValue) =>
          cv match
            case CellValue.Error(_) =>
              evalAny(ctx, valueIfErrorExpr).map(toCellValue)
            case _ => Right(cv)
        case Right(other) =>
          Right(toCellValue(other))
    }

  val iserror: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISERROR", Arity.one) { (expr, ctx) =>
      evalAny(ctx, expr) match
        case Left(_) => Right(true)
        case Right(cv: CellValue) =>
          cv match
            case CellValue.Error(_) => Right(true)
            case _ => Right(false)
        case Right(_) => Right(false)
    }

  val iserr: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISERR", Arity.one) { (expr, ctx) =>
      evalAny(ctx, expr) match
        case Left(_) => Right(true)
        case Right(cv: CellValue) =>
          cv match
            case CellValue.Error(err) => Right(err != CellError.NA)
            case _ => Right(false)
        case Right(_) => Right(false)
    }

  val isnumber: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISNUMBER", Arity.one) { (expr, ctx) =>
      evalAny(ctx, expr) match
        case Left(_) => Right(false)
        case Right(cv: CellValue) =>
          cv match
            case CellValue.Number(_) => Right(true)
            case CellValue.Formula(_, Some(CellValue.Number(_))) => Right(true)
            case _ => Right(false)
        case Right(_: BigDecimal) => Right(true)
        case Right(_: Int) => Right(true)
        case Right(_: Long) => Right(true)
        case Right(_: Double) => Right(true)
        case Right(_) => Right(false)
    }

  val istext: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISTEXT", Arity.one) { (expr, ctx) =>
      evalAny(ctx, expr) match
        case Left(_) => Right(false)
        case Right(cv: CellValue) =>
          cv match
            case CellValue.Text(_) => Right(true)
            case CellValue.Formula(_, Some(CellValue.Text(_))) => Right(true)
            case _ => Right(false)
        case Right(_: String) => Right(true)
        case Right(_) => Right(false)
    }

  val isblank: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISBLANK", Arity.one) { (expr, ctx) =>
      evalAny(ctx, expr) match
        case Left(_) => Right(false)
        case Right(cv: CellValue) =>
          cv match
            case CellValue.Empty => Right(true)
            case _ => Right(false)
        case Right(_) => Right(false)
    }
