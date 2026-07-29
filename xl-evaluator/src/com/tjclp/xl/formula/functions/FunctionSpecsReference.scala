package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}

trait FunctionSpecsReference extends FunctionSpecsBase:
  // extractARef is inherited from FunctionSpecsBase (shared with OFFSET).

  private def extractCellRange(expr: TExpr[?]): Option[CellRange] = expr match
    case TExpr.RangeRef(range) => Some(range)
    case TExpr.SheetRange(_, range) => Some(range)
    case _ => None

  /** The sheet a reference expression is qualified with, if any (CELL reads the qualifier). */
  private def extractSheetName(expr: TExpr[?]): Option[SheetName] = expr match
    case TExpr.SheetPolyRef(sheet, _, _) => Some(sheet)
    case TExpr.SheetRef(sheet, _, _, _) => Some(sheet)
    case TExpr.SheetRange(sheet, _) => Some(sheet)
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

  val row: FunctionSpec[BigDecimal] { type Args = Option[AnyExpr] } =
    FunctionSpec.simple[BigDecimal, Option[AnyExpr]](
      "ROW",
      Arity.Range(0, 1),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (exprOpt, ctx) =>
      exprOpt match
        case Some(expr) =>
          extractARef(expr) match
            case Some(aref) => Right(BigDecimal(aref.row.index0 + 1))
            case None =>
              Left(
                EvalError.EvalFailed(
                  "ROW requires a cell reference",
                  Some(s"ROW($expr)")
                )
              )
        case None =>
          // Zero-argument form: ROW() returns row of current cell
          ctx.currentCell match
            case Some(ref) => Right(BigDecimal(ref.row.index0 + 1))
            case None =>
              Left(
                EvalError.EvalFailed(
                  "ROW() with no arguments requires a cell context",
                  Some("ROW()")
                )
              )
    }

  val column: FunctionSpec[BigDecimal] { type Args = Option[AnyExpr] } =
    FunctionSpec.simple[BigDecimal, Option[AnyExpr]](
      "COLUMN",
      Arity.Range(0, 1),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (exprOpt, ctx) =>
      exprOpt match
        case Some(expr) =>
          extractARef(expr) match
            case Some(aref) => Right(BigDecimal(aref.col.index0 + 1))
            case None =>
              Left(
                EvalError.EvalFailed(
                  "COLUMN requires a cell reference",
                  Some(s"COLUMN($expr)")
                )
              )
        case None =>
          // Zero-argument form: COLUMN() returns column of current cell
          ctx.currentCell match
            case Some(ref) => Right(BigDecimal(ref.col.index0 + 1))
            case None =>
              Left(
                EvalError.EvalFailed(
                  "COLUMN() with no arguments requires a cell context",
                  Some("COLUMN()")
                )
              )
    }

  val rows: FunctionSpec[BigDecimal] { type Args = AnyExpr } =
    FunctionSpec.simple[BigDecimal, AnyExpr](
      "ROWS",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (expr, ctx) =>
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
    FunctionSpec.simple[BigDecimal, AnyExpr](
      "COLUMNS",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (expr, ctx) =>
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

  /**
   * GH-424: CELL(info_type, [reference]) — the info arms the house corpus uses. Every print tab in
   * the exemplar books carries `=CELL("Filename",$B$2)` as its path + tab self-label.
   *
   * Supported arms (case-insensitive like Excel): "filename" (directory + [file] + sheet from the
   * eval context's workbook path — the empty string when no path is known, Excel's pre-save
   * behavior), "address" (absolute A1 form of the reference's top-left cell), "row" and "col"
   * (1-based coordinates). The reference argument is positional — its ADDRESS is the datum, so it
   * is never evaluated; omitted, the arms fall back to the current cell (ROW()/COLUMN() convention;
   * "filename" needs only the ambient sheet). Exotic arms (format, contents, parentheses, prefix,
   * ...) are Excel's #VALUE! as a clean per-cell error value, never a throw.
   *
   * Volatile like NOW(): the value derives from the evaluation context (workbook path, position),
   * never from cell data, so full recalculation always recomputes it fresh.
   */
  val cellFn: FunctionSpec[CellValue] { type Args = CellArgs } =
    FunctionSpec.simple[CellValue, CellArgs]("CELL", Arity.Range(1, 2)) { (args, ctx) =>
      val (infoTypeExpr, refOpt) = args
      ctx.evalExpr(infoTypeExpr).flatMap { infoType =>
        // (top-left ARef if the reference or cell context provides one, subject sheet)
        val target: Either[EvalError, (Option[ARef], SheetName)] = refOpt match
          case Some(expr) =>
            extractARef(expr) match
              case Some(aref) =>
                Right((Some(aref), extractSheetName(expr).getOrElse(ctx.sheet.name)))
              case None =>
                Left(
                  EvalError.EvalFailed(
                    "CELL requires a cell reference",
                    Some(s"CELL($infoType, $expr)")
                  )
                )
          case None => Right((ctx.currentCell, ctx.sheet.name))
        def positional(f: ARef => CellValue): Either[EvalError, CellValue] =
          target.flatMap {
            case (Some(aref), _) => Right(f(aref))
            case (None, _) =>
              Left(
                EvalError.EvalFailed(
                  s"CELL(\"$infoType\") with no reference requires a cell context",
                  Some(s"CELL($infoType)")
                )
              )
          }
        infoType.toLowerCase match
          case "filename" =>
            target.map { case (_, sheetName) =>
              ctx.workbookPath match
                case None => CellValue.Text("") // Excel: empty until the workbook is saved
                case Some(path) =>
                  // Excel's shape: directory + [file] + sheet, e.g. /models/[lbo.xlsx]DCF
                  val cut = math.max(path.lastIndexOf('/'), path.lastIndexOf('\\')) + 1
                  CellValue.Text(s"${path.take(cut)}[${path.drop(cut)}]${sheetName.value}")
            }
          case "address" =>
            positional(aref =>
              CellValue.Text(s"$$${columnToLetter(aref.col.index0)}$$${aref.row.index0 + 1}")
            )
          case "row" => positional(aref => CellValue.Number(BigDecimal(aref.row.index0 + 1)))
          case "col" => positional(aref => CellValue.Number(BigDecimal(aref.col.index0 + 1)))
          case other =>
            Left(
              EvalError.ErrorValue(
                CellError.Value,
                Some(
                  s"CELL: info_type '$other' is not supported " +
                    "(supported: filename, address, row, col)"
                )
              )
            )
      }
    }
