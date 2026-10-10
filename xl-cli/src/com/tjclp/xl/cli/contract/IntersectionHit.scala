package com.tjclp.xl.cli.contract

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.ImplicitIntersection.Divergence

/**
 * GH-714: the `IMPLICIT_INTERSECTION` warning for the plain formula cells a write left whose value
 * as a plain cell (Excel's legacy implicit intersection) differs from their value as an array
 * formula — modelled on [[OffGridHit.warning]]: one warning, the first ten cells listed, then a
 * count, located at the first cell. Informational: never a `--strict` reason.
 */
object IntersectionHit:

  private val MaxListed = 10

  private val remedy = "use SUMPRODUCT, putf --array (batch \"array\": true), or an explicit @"

  private def where(d: Divergence): String =
    s"${SheetName.quoteForFormula(d.sheet.value)}!${d.cell.toA1}"

  private def show(v: CellValue): String = v match
    case CellValue.Number(n) => n.bigDecimal.stripTrailingZeros.toPlainString
    case CellValue.Text(s) => s"\"$s\""
    case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
    case CellValue.Error(e) =>
      import CellError.toExcel
      e.toExcel
    case CellValue.Empty => "0"
    case CellValue.Formula(_, Some(cached), _) => show(cached)
    case other => other.toString

  /** What the array formula gives: its value, or for a spill its shape and top-left value. */
  private def arrayValue(d: Divergence): String =
    if d.isShape then s"a ${d.rows}x${d.cols} spill (top-left ${show(d.arrayTopLeft)})"
    else show(d.arrayTopLeft)

  private def sentence(d: Divergence): String =
    s"${where(d)}: =${d.formula} is ${show(d.plain)} as a plain cell (implicit intersection) " +
      s"but ${arrayValue(d)} as an array formula"

  def warning(hits: Vector[Divergence]): Option[Warning] =
    hits.headOption.map { first =>
      val message =
        if hits.size == 1 then s"${sentence(first)}; $remedy"
        else
          val listed = hits.take(MaxListed).map(sentence).mkString("; ")
          val more = if hits.size > MaxListed then s"; … (+${hits.size - MaxListed} more)" else ""
          s"${hits.size} formulas evaluate differently as plain cells than as array formulas: " +
            s"$listed$more; $remedy"
      Warning(
        WarningCode.IMPLICIT_INTERSECTION,
        message,
        Some(Location(None, Some(first.sheet.value), Some(first.cell.toA1), None))
      )
    }
