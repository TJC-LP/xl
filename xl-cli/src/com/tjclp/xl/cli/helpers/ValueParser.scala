package com.tjclp.xl.cli.helpers

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formatted.{Formatted, FormattedParsers}
import com.tjclp.xl.styles.numfmt.NumFmt

import com.tjclp.xl.cli.output.RendererCommon

/**
 * Value parsing utilities for CLI commands.
 *
 * Provides helpers for parsing string inputs to CellValue and formatting CellValue for display.
 */
object ValueParser:

  private def textValue(s: String): CellValue =
    val text = if s.startsWith("\"") && s.endsWith("\"") then s.drop(1).dropRight(1) else s
    CellValue.Text(text)

  /**
   * Parse a positional `put` value using the same detection as batch `put`.
   *
   * Detection is total: malformed lookalikes remain text. When disabled, the input is always text.
   */
  def parsePutValue(s: String, detect: Boolean): Formatted =
    if detect then
      FormattedParsers.detect(s) match
        case Formatted(CellValue.Text(_), numFmt) => Formatted(textValue(s), numFmt)
        case formatted => formatted
    else Formatted(textValue(s), NumFmt.General)

  /**
   * Parse a string into a CellValue.
   *
   * Attempts to parse as:
   *   1. Number (BigDecimal)
   *   2. Boolean (true/false, case-insensitive)
   *   3. Text (with optional quote stripping)
   *
   * @param s
   *   String to parse
   * @return
   *   Parsed CellValue
   */
  def parseValue(s: String): CellValue =
    scala.util.Try(BigDecimal(s)).toOption.map(CellValue.Number.apply).getOrElse {
      s.toLowerCase match
        case "true" => CellValue.Bool(true)
        case "false" => CellValue.Bool(false)
        case _ => textValue(s)
    }

  /**
   * GH-430: `TABLE(...)` is the derived display text of a `<f t="dataTable">` record, not a real
   * Excel function — writing it as a plain formula would be a #NAME? error on open. Returns the
   * rejection message for a top-level TABLE( expression (case-insensitive, leading `=` stripped);
   * None when the formula is writable. The steering points at the GH-419 authoring surfaces.
   */
  def dataTableFormulaError(formula: String): Option[String] =
    val stripped = formula.trim.stripPrefix("=").trim
    Option.when(stripped.toUpperCase.startsWith("TABLE("))(
      s"'$formula' is a data-table record display text, not a writable formula " +
        "(Excel would show #NAME?). Author it with the data-table batch op " +
        """({"op":"data-table","ref":"D5:F6","rowInput":"B1","colInput":"B2"}) or """ +
        "sheet.dataTable(interior, rowInput, colInput) in scripts (GH-419)."
    )

  /**
   * Format a CellValue for display.
   *
   * @param value
   *   CellValue to format
   * @return
   *   String representation
   */
  def formatCellValue(value: CellValue): String =
    value match
      case CellValue.Text(s) => s
      case CellValue.Number(n) =>
        if n.isWhole then n.toBigInt.toString
        else n.underlying.stripTrailingZeros.toPlainString
      case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
      case CellValue.DateTime(dt) => dt.toString
      case CellValue.Error(err) => err.toExcel
      case CellValue.RichText(rt) => rt.toPlainText
      case CellValue.Empty => ""
      case CellValue.Formula(expr, cached, kind) =>
        val displayExpr = RendererCommon.formulaDisplay(expr, kind)
        cached.map(formatCellValue).getOrElse(displayExpr)
