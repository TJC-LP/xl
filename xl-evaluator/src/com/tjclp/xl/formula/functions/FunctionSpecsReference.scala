package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{ArrayArithmetic, ArrayResult, EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet

trait FunctionSpecsReference extends FunctionSpecsBase:
  // extractARef is inherited from FunctionSpecsBase (shared with OFFSET).

  private def extractCellRange(expr: TExpr[?]): Option[CellRange] = expr match
    case TExpr.RangeRef(range, _) => Some(range)
    case TExpr.SheetRange(_, range, _) => Some(range)
    case _ => None

  /**
   * GH-612: the error an error-literal argument names — `ROW(#REF!)`, `ROWS(#REF!)` are that error,
   * not a host failure. Takes `TExpr[?]` so the `TExpr[Nothing]` case is not judged unreachable
   * against a `TExpr[Any]` argument.
   */
  private def errorLiteral(expr: TExpr[?]): Option[CellError] = expr match
    case TExpr.ErrorLit(error) => Some(error)
    case _ => None

  /** The sheet a reference expression is qualified with, if any (CELL reads the qualifier). */
  private def extractSheetName(expr: TExpr[?]): Option[SheetName] = expr match
    case TExpr.SheetPolyRef(sheet, _, _) => Some(sheet)
    case TExpr.SheetRef(sheet, _, _, _) => Some(sheet)
    case TExpr.SheetRange(sheet, _, _) => Some(sheet)
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

  /**
   * GH-476: HYPERLINK(link_location, [friendly_name]) — absent from the roster, so an ordinary
   * "click through to the source" cell aborted recalc with an UnknownFunction host error.
   *
   * The evaluated VALUE of a HYPERLINK cell is its display text: friendly_name when supplied, the
   * link itself otherwise. The jump target is display-layer metadata, not part of the value plane
   * (xl models real hyperlinks on `Sheet.hyperlinks`).
   */
  val hyperlink: FunctionSpec[String] { type Args = HyperlinkArgs } =
    FunctionSpec.simple[String, HyperlinkArgs]("HYPERLINK", Arity.Range(1, 2)) { (args, ctx) =>
      val (linkExpr, friendlyOpt) = args
      for
        link <- ctx.evalExpr(linkExpr)
        friendly <- friendlyOpt.fold[Either[EvalError, Option[String]]](Right(None))(expr =>
          ctx.evalExpr(expr).map(Some(_))
        )
      yield friendly.getOrElse(link)
    }

  val row: FunctionSpec[BigDecimal] { type Args = Option[AnyExpr] } =
    FunctionSpec.simple[BigDecimal, Option[AnyExpr]](
      "ROW",
      Arity.Range(0, 1),
      flags = FunctionFlags(returnsNumeric = true)
    ) { (exprOpt, ctx) =>
      exprOpt match
        case Some(expr) =>
          // GH-612: ROW(#REF!) is #REF!, the error value the argument names
          errorLiteral(expr).map(e => Left(EvalError.ErrorValue(e))).getOrElse {
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
          errorLiteral(expr).map(e => Left(EvalError.ErrorValue(e))).getOrElse {
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
      // GH-612: ROWS(#REF!) is #REF!, the error value the argument names
      errorLiteral(expr).map(e => Left(EvalError.ErrorValue(e))).getOrElse {
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
    }

  val columns: FunctionSpec[BigDecimal] { type Args = AnyExpr } =
    FunctionSpec.simple[BigDecimal, AnyExpr](
      "COLUMNS",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true)
    ) { (expr, ctx) =>
      errorLiteral(expr).map(e => Left(EvalError.ErrorValue(e))).getOrElse {
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

  // ===== GH-604: SINGLE — the implicit-intersection operator `@` =====

  /** A defined-name argument, as the range location it may denote. */
  private def nameLocation(expr: TExpr[?]): Option[TExpr.RangeLocation] = expr match
    case TExpr.NameRef(name) => Some(TExpr.RangeLocation.Name(name, None))
    case TExpr.SheetNameRef(sheet, name) => Some(TExpr.RangeLocation.Name(name, Some(sheet)))
    case _ => None

  /**
   * Excel's implicit intersection of `range` with the formula's own cell: a single cell is itself;
   * a one-row range yields the cell in the formula's column, a one-column range the cell in the
   * formula's row; anything else — a 2-D range, or a vector the formula's row/column does not cross
   * — is `#VALUE!`.
   */
  private def intersectionCell(range: CellRange, current: Option[ARef]): Either[EvalError, ARef] =
    if range.width == 1 && range.height == 1 then Right(range.start)
    else
      current match
        case None =>
          Left(
            EvalError.EvalFailed(
              s"@${range.toA1} (implicit intersection) needs the formula's cell position",
              Some("@range")
            )
          )
        case Some(cell) =>
          val col = cell.col.index0
          val row = cell.row.index0
          if range.height == 1 && col >= range.colStart.index0 && col <= range.colEnd.index0 then
            Right(ARef.from0(col, range.rowStart.index0))
          else if range.width == 1 && row >= range.rowStart.index0 && row <= range.rowEnd.index0
          then Right(ARef.from0(range.colStart.index0, row))
          else
            Left(
              EvalError.ErrorValue(
                CellError.Value,
                Some(s"@${range.toA1}: no cell in the formula's row or column (${cell.toA1})")
              )
            )

  /**
   * Shapes that print as one primary — everything the `@` operand slot re-parses unparenthesized.
   */
  private def isPrimaryShape(expr: TExpr[?]): Boolean = expr match
    case _: TExpr.Add | _: TExpr.Sub | _: TExpr.Mul | _: TExpr.Div | _: TExpr.Pow |
        _: TExpr.Percent | _: TExpr.Concat | _: TExpr.UnaryPlus[?] | _: TExpr.Eq[?] |
        _: TExpr.Neq[?] | _: TExpr.Lt[?] | _: TExpr.Lte[?] | _: TExpr.Gt[?] | _: TExpr.Gte[?] =>
      false
    case TExpr.Coerced(inner, _) => isPrimaryShape(inner)
    case TExpr.ToInt(inner) => isPrimaryShape(inner)
    case TExpr.DateToSerial(inner) => isPrimaryShape(inner)
    case TExpr.DateTimeToSerial(inner) => isPrimaryShape(inner)
    case _ => true

  private def renderIntersectionOperand(arg: ArgSpec.SumProductArg, printer: ArgPrinter): String =
    arg match
      case Left(location) => printer.location(location)
      case Right(expr) =>
        val text = printer.expr(expr)
        if isPrimaryShape(expr) then text else s"($text)"

  /**
   * SINGLE(x) — GH-604: the stored form (`_xlfn.SINGLE`) of Excel 365's implicit-intersection
   * operator, spelled `@x` in the formula bar and in this model (the parser accepts both, the
   * printer emits `@x`, `FormulaStorage` maps the two at the `<f>` boundary).
   *
   * Semantics: a scalar or a single cell is itself; a range intersects with the formula's cell (see
   * [[intersectionCell]]) — the cell in the formula's row for a column vector, in its column for a
   * row vector, `#VALUE!` otherwise; a defined name bound to a range intersects the same way; an
   * array VALUE (a call result, not a reference) collapses to its top-left element.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  val single: FunctionSpec[CellValue] { type Args = ArgSpec.SumProductArg } =
    FunctionSpec.simple[CellValue, ArgSpec.SumProductArg](
      "SINGLE",
      Arity.one,
      renderFn = Some((arg, printer) => s"@${renderIntersectionOperand(arg, printer)}")
    ) { (arg, ctx) =>
      def intersect(targetSheet: Sheet, range: CellRange): Either[EvalError, CellValue] =
        intersectionCell(range, ctx.currentCell).flatMap(rangeCellReader(targetSheet, ctx))
      def resolved(location: TExpr.RangeLocation): Option[(Sheet, CellRange)] =
        Evaluator.resolveRangeLocation(location, ctx.sheet, ctx.workbook).toOption
      arg match
        case Left(location) =>
          Evaluator.resolveRangeLocation(location, ctx.sheet, ctx.workbook).flatMap {
            case (targetSheet, range) => intersect(targetSheet, range)
          }
        case Right(expr) =>
          nameLocation(expr).flatMap(resolved) match
            case Some((targetSheet, range)) => intersect(targetSheet, range)
            case None =>
              val value = expr match
                case _: TExpr.PolyRef | _: TExpr.SheetPolyRef | _: TExpr.UnaryPlus[?] =>
                  TExpr.asResolvedValueExpr(expr)
                case other => other
              ctx.evalArrayExpr(value.asInstanceOf[TExpr[Any]]).map {
                case ar: ArrayResult => if ar.isEmpty then CellValue.Empty else ar(0, 0)
                case scalar => ArrayArithmetic.anyToCellValue(scalar)
              }
    }
