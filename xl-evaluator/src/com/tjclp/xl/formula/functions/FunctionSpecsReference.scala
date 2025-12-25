package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.addressing.{ARef, CellRange}

trait FunctionSpecsReference extends FunctionSpecsBase:
  private def extractARef(expr: TExpr[?]): Option[ARef] = expr match
    case TExpr.PolyRef(ref, _) => Some(ref)
    case TExpr.Ref(ref, _, _) => Some(ref)
    case TExpr.SheetPolyRef(_, ref, _) => Some(ref)
    case TExpr.SheetRef(_, ref, _, _) => Some(ref)
    case TExpr.RangeRef(range) => Some(range.start)
    case TExpr.SheetRange(_, range) => Some(range.start)
    case _ => None

  private def extractCellRange(expr: TExpr[?]): Option[CellRange] = expr match
    case TExpr.RangeRef(range) => Some(range)
    case TExpr.SheetRange(_, range) => Some(range)
    case _ => None

  @annotation.tailrec
  private def columnToLetter(col: Int, acc: String = ""): String =
    if col < 0 then acc
    else if acc.isEmpty && col <= 25 then ('A' + col).toChar.toString
    else
      val remainder = col % 26
      val quotient = col / 26 - 1
      val letter = ('A' + remainder).toChar
      if quotient < 0 then letter.toString + acc
      else columnToLetter(quotient, letter.toString + acc)

  val row: FunctionSpec[BigDecimal] { type Args = AnyExpr } =
    FunctionSpec.simple[BigDecimal, AnyExpr]("ROW", Arity.one) { (expr, ctx) =>
      extractARef(expr) match
        case Some(aref) => Right(BigDecimal(aref.row.index0 + 1))
        case None =>
          Left(
            EvalError.EvalFailed(
              "ROW requires a cell reference",
              Some(s"ROW($expr)")
            )
          )
    }

  val column: FunctionSpec[BigDecimal] { type Args = AnyExpr } =
    FunctionSpec.simple[BigDecimal, AnyExpr]("COLUMN", Arity.one) { (expr, ctx) =>
      extractARef(expr) match
        case Some(aref) => Right(BigDecimal(aref.col.index0 + 1))
        case None =>
          Left(
            EvalError.EvalFailed(
              "COLUMN requires a cell reference",
              Some(s"COLUMN($expr)")
            )
          )
    }

  val rows: FunctionSpec[BigDecimal] { type Args = AnyExpr } =
    FunctionSpec.simple[BigDecimal, AnyExpr]("ROWS", Arity.one) { (expr, ctx) =>
      extractCellRange(expr) match
        case Some(range) =>
          val rowCount = range.rowEnd.index0 - range.rowStart.index0 + 1
          Right(BigDecimal(rowCount))
        case None =>
          Left(
            EvalError.EvalFailed(
              "ROWS requires a range argument",
              Some(s"ROWS($expr)")
            )
          )
    }

  val columns: FunctionSpec[BigDecimal] { type Args = AnyExpr } =
    FunctionSpec.simple[BigDecimal, AnyExpr]("COLUMNS", Arity.one) { (expr, ctx) =>
      extractCellRange(expr) match
        case Some(range) =>
          val colCount = range.colEnd.index0 - range.colStart.index0 + 1
          Right(BigDecimal(colCount))
        case None =>
          Left(
            EvalError.EvalFailed(
              "COLUMNS requires a range argument",
              Some(s"COLUMNS($expr)")
            )
          )
    }

  val address: FunctionSpec[String] { type Args = AddressArgs } =
    FunctionSpec.simple[String, AddressArgs]("ADDRESS", Arity.Range(2, 5)) { (args, ctx) =>
      val (rowExpr, colExpr, absNumOpt, a1Opt, sheetOpt) = args
      val absNumExpr = absNumOpt.getOrElse(TExpr.Lit(BigDecimal(1)))
      val a1Expr = a1Opt.getOrElse(TExpr.Lit(true))
      for
        row <- ctx.evalExpr(rowExpr)
        col <- ctx.evalExpr(colExpr)
        absNum <- ctx.evalExpr(absNumExpr)
        a1Style <- ctx.evalExpr(a1Expr)
        sheetName <- sheetOpt match
          case Some(expr) => ctx.evalExpr(expr).map(Some(_))
          case None => Right(None)
      yield
        val rowInt = row.toInt
        val colInt = col.toInt
        val absType = absNum.toInt

        if rowInt < 1 || colInt < 1 then "#VALUE!"
        else if a1Style then
          val colLetter = columnToLetter(colInt - 1)
          val (colPrefix, rowPrefix) = absType match
            case 1 => ("$", "$")
            case 2 => ("", "$")
            case 3 => ("$", "")
            case _ => ("", "")
          val refStr = s"$colPrefix$colLetter$rowPrefix$rowInt"
          sheetName match
            case Some(sn) => s"$sn!$refStr"
            case None => refStr
        else
          val rowPart = absType match
            case 1 | 2 => s"R$rowInt"
            case _ => s"R[$rowInt]"
          val colPart = absType match
            case 1 | 3 => s"C$colInt"
            case _ => s"C[$colInt]"
          val refStr = s"$rowPart$colPart"
          sheetName match
            case Some(sn) => s"$sn!$refStr"
            case None => refStr
    }
