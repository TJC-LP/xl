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
        case Left(_) => Right(true)
        case Right(ExprValue.Cell(CellValue.Error(err))) => Right(err != CellError.NA)
        case Right(_) => Right(false)
    }

  val isnumber: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISNUMBER", Arity.one) { (expr, ctx) =>
      evalValue(ctx, expr) match
        case Left(_) => Right(false)
        case Right(ExprValue.Cell(CellValue.Number(_))) => Right(true)
        case Right(ExprValue.Cell(CellValue.Formula(_, Some(CellValue.Number(_))))) =>
          Right(true)
        case Right(ExprValue.Number(_)) => Right(true)
        case Right(_) => Right(false)
    }

  val istext: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue]("ISTEXT", Arity.one) { (expr, ctx) =>
      evalValue(ctx, expr) match
        case Left(_) => Right(false)
        case Right(ExprValue.Cell(CellValue.Text(_))) => Right(true)
        case Right(ExprValue.Cell(CellValue.Formula(_, Some(CellValue.Text(_))))) =>
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
