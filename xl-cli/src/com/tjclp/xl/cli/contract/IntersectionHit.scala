package com.tjclp.xl.cli.contract

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.ImplicitIntersection.{Divergence, Report}

/**
 * GH-714: the `IMPLICIT_INTERSECTION` warning for the plain formula cells a write left whose value
 * as a plain cell (Excel's legacy implicit intersection) differs from their value as an array
 * formula — modelled on [[OffGridHit.warning]]: one warning, the first ten cells listed, then a
 * count, located at the first cell. A check that stopped at its budget says how many cells it
 * evaluated. Informational: never a `--strict` reason.
 */
object IntersectionHit:

  private val MaxListed = 10

  private val remedy = "use SUMPRODUCT, putf --array (batch \"array\": true), or an explicit @"

  private def where(d: Divergence): String =
    s"${SheetName.quoteForFormula(d.sheet.value)}!${d.cell.toA1}"

  /** A value as Excel writes it in a formula: text quoted, its own quotes doubled (`"a""b"`). */
  private def show(v: CellValue): String = v match
    case CellValue.Number(n) => n.bigDecimal.stripTrailingZeros.toPlainString
    case CellValue.Text(s) => s"\"${s.replace("\"", "\"\"")}\""
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

  /** `1001` or `at least 1001`: how many candidates there were, as far as the check counted. */
  private def candidates(report: Report): String =
    if report.exact then s"${report.candidates}" else s"at least ${report.candidates}"

  /**
   * #739: the `INTERSECTION_SAMPLED` note for a check that stopped at its budget with no divergence
   * among the cells it evaluated — silence would read as "all clear". None when the check found a
   * divergence (the warning says it sampled) or evaluated every candidate.
   */
  def sampledNote(report: Report): Option[Warning] =
    Option.when(report.sampled && report.hits.isEmpty)(
      Warning(
        WarningCode.INTERSECTION_SAMPLED,
        s"implicit-intersection check: the first ${report.checked} of ${candidates(report)} " +
          "candidate cells were checked and none evaluates differently as a plain cell than as " +
          "an array formula; the rest were not checked"
      )
    )

  def warning(report: Report): Option[Warning] =
    val hits = report.hits
    val sampled =
      if report.sampled then s" (first ${report.checked} of ${candidates(report)} cells checked)"
      else ""
    hits.headOption.map { first =>
      val message =
        if hits.size == 1 then s"${sentence(first)}; $remedy$sampled"
        else
          val listed = hits.take(MaxListed).map(sentence).mkString("; ")
          val more = if hits.size > MaxListed then s"; … (+${hits.size - MaxListed} more)" else ""
          s"${hits.size} formulas evaluate differently as plain cells than as array formulas: " +
            s"$listed$more; $remedy$sampled"
      Warning(
        WarningCode.IMPLICIT_INTERSECTION,
        message,
        Some(Location(None, Some(first.sheet.value), Some(first.cell.toA1), None))
      )
    }
