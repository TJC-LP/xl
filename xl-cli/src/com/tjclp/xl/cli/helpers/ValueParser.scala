package com.tjclp.xl.cli.helpers

import com.tjclp.xl.cells.CellValue

import com.tjclp.xl.cli.output.RendererCommon

/**
 * Value parsing utilities for CLI commands.
 *
 * Provides helpers for parsing string inputs to CellValue and formatting CellValue for display.
 */
object ValueParser:

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
        case _ =>
          val text = if s.startsWith("\"") && s.endsWith("\"") then s.drop(1).dropRight(1) else s
          CellValue.Text(text)
    }

  /**
   * GH-430: `TABLE(...)` is the derived display text of a `<f t="dataTable">` record, not a real
   * Excel function — writing it as a plain formula would be a #NAME? error on open. Returns the
   * rejection message for a top-level TABLE( expression (case-insensitive, leading `=` stripped);
   * None when the formula is writable. Data-table authoring is tracked in GH-419.
   */
  def dataTableFormulaError(formula: String): Option[String] =
    val stripped = formula.trim.stripPrefix("=").trim
    Option.when(stripped.toUpperCase.startsWith("TABLE("))(
      s"'$formula' is a data-table record display text, not a writable formula " +
        "(Excel would show #NAME?). Data-table authoring is tracked in GH-419."
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
