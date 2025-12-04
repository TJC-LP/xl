package com.tjclp.xl.cli.output

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{CellValue, Comment}
import com.tjclp.xl.error.XLError
import com.tjclp.xl.styles.{CellStyle, StyleId}
import com.tjclp.xl.styles.alignment.Align
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Output formatting utilities for xl CLI.
 */
object Format:

  /**
   * Format a success message for a put operation.
   */
  def putSuccess(ref: ARef, value: CellValue): String =
    val typeStr = valueType(value)
    val valueStr = formatValue(value)
    s"Put: ${ref.toA1} = $valueStr ($typeStr)"

  /**
   * Format a batch put success message.
   */
  def batchSuccess(updates: Vector[(ARef, CellValue)]): String =
    val sb = new StringBuilder
    sb.append(s"Batch update: ${updates.size} cells\n")
    updates.foreach { case (ref, value) =>
      val typeStr = valueType(value)
      val valueStr = formatValue(value)
      sb.append(s"  ${ref.toA1} = $valueStr ($typeStr)\n")
    }
    sb.toString.stripSuffix("\n")

  /**
   * Format a workbook open success message.
   */
  def openSuccess(path: String, sheets: Vector[String]): String =
    val sheetList = sheets.mkString(", ")
    s"""Opened: $path
       |Sheets: $sheetList (${sheets.size} total)
       |Active: ${sheets.headOption.getOrElse("(none)")}""".stripMargin

  /**
   * Format a workbook create success message.
   */
  def createSuccess(sheets: Vector[String]): String =
    val sheetList = sheets.mkString(", ")
    s"""Created new workbook
       |Sheets: $sheetList (${sheets.size} total)
       |Active: ${sheets.headOption.getOrElse("(none)")}""".stripMargin

  /**
   * Format a save success message.
   */
  def saveSuccess(path: String, sheetCount: Int, cellCount: Int): String =
    s"Saved: $path ($sheetCount sheets, $cellCount cells)"

  /**
   * Format an eval success message.
   */
  def evalSuccess(formula: String, result: CellValue, overrides: List[String]): String =
    val resultStr = formatValue(result)
    val typeStr = valueType(result)
    val overridesStr =
      if overrides.isEmpty then ""
      else s"\nWith: ${overrides.mkString(", ")}"
    s"""Formula: $formula
       |Result: $resultStr ($typeStr)$overridesStr""".stripMargin

  /**
   * Format a sheet select success message.
   */
  def selectSuccess(
    name: String,
    usedRange: Option[String],
    cellCount: Int,
    formulaCount: Int
  ): String =
    val rangeStr = usedRange.getOrElse("(empty)")
    s"""Selected: $name
       |Used range: $rangeStr
       |Cells: $cellCount non-empty, $formulaCount formulas""".stripMargin

  /**
   * Format an error message with location and suggestion.
   */
  def error(err: XLError): String =
    val (errType, details, suggestion) = errorDetails(err)
    s"""Error: $errType
       |Details: $details
       |Suggestion: $suggestion""".stripMargin

  /**
   * Format a simple error message.
   */
  def errorSimple(message: String): String =
    s"Error: $message"

  /**
   * Format cell info output with full details.
   *
   * Shows raw value, formatted value (using NumFmt), style info (non-default only), comment,
   * hyperlink, and formula dependencies.
   */
  def cellInfo(
    ref: ARef,
    value: CellValue,
    formatted: String,
    style: Option[CellStyle],
    comment: Option[Comment],
    hyperlink: Option[String],
    dependencies: Vector[ARef],
    dependents: Vector[ARef]
  ): String =
    val sb = new StringBuilder
    sb.append(s"Cell: ${ref.toA1}\n")
    sb.append(s"Type: ${valueType(value)}\n")

    // For formulas, show expression and cached value separately
    value match
      case CellValue.Formula(expr, cached) =>
        val displayExpr = if expr.startsWith("=") then expr else s"=$expr"
        sb.append(s"Formula: $displayExpr\n")
        cached.foreach { v =>
          sb.append(s"Cached: ${formatValue(v)}\n")
          // Only show Formatted for cached value if different from raw
          val cachedRaw = formatValue(v)
          if formatted != cachedRaw && formatted != cachedRaw.stripPrefix("\"").stripSuffix("\"")
          then sb.append(s"Formatted: $formatted\n")
        }
      case CellValue.Empty =>
        sb.append("Value: (empty)\n")
      case _ =>
        sb.append(s"Raw: ${formatValue(value)}\n")
        // Only show Formatted if different from Raw
        val rawStr = formatValue(value)
        if formatted != rawStr && formatted != rawStr.stripPrefix("\"").stripSuffix("\"") then
          sb.append(s"Formatted: $formatted\n")

    // Style (non-default properties only)
    formatStyle(style).foreach(s => sb.append(s).append("\n"))

    // Comment
    comment.foreach { c =>
      val authorStr = c.author.map(a => s" (Author: $a)").getOrElse("")
      sb.append(s"""Comment: "${c.text.toPlainText}"$authorStr\n""")
    }

    // Hyperlink
    hyperlink.foreach(h => sb.append(s"Hyperlink: $h\n"))

    // Dependencies and dependents
    val depsStr = if dependencies.isEmpty then "(none)" else dependencies.map(_.toA1).mkString(", ")
    val deptsStr = if dependents.isEmpty then "(none)" else dependents.map(_.toA1).mkString(", ")
    sb.append(s"Dependencies: $depsStr\n")
    sb.append(s"Dependents: $deptsStr")

    sb.toString

  /**
   * Format style properties (non-default only).
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def formatStyle(style: Option[CellStyle]): Option[String] =
    style.flatMap { s =>
      val parts = Vector.newBuilder[String]

      // Font (if non-default)
      if s.font != Font.default then
        var fontDesc = Vector(s.font.name, s"${s.font.sizePt}pt")
        if s.font.bold then fontDesc = fontDesc :+ "bold"
        if s.font.italic then fontDesc = fontDesc :+ "italic"
        if s.font.underline then fontDesc = fontDesc :+ "underline"
        s.font.color.foreach {
          case Color.Rgb(argb) =>
            // Show RGB color as 6-digit hex (drop alpha)
            fontDesc = fontDesc :+ f"${argb & 0xffffff}%06X"
          case Color.Theme(slot, tint) =>
            // Show theme color as descriptive string
            val tintStr = if tint == 0.0 then "" else f" tint=$tint%.2f"
            fontDesc = fontDesc :+ s"$slot$tintStr"
        }
        parts += s"Font: ${fontDesc.mkString(" ")}"

      // Fill (if non-default)
      s.fill match
        case Fill.Solid(color) =>
          val colorStr = color match
            case Color.Rgb(argb) => f"${argb & 0xffffff}%06X"
            case Color.Theme(slot, tint) =>
              val tintStr = if tint == 0.0 then "" else f" tint=$tint%.2f"
              s"$slot$tintStr"
          parts += s"Fill: $colorStr (solid)"
        case Fill.Pattern(_, _, _) => parts += "Fill: (pattern)"
        case Fill.None => ()

      // NumFmt (if non-General)
      if s.numFmt != NumFmt.General then
        val idStr = s.numFmtId.map(id => s" (id: $id)").getOrElse("")
        val codeStr = s.numFmt match
          case NumFmt.Custom(code) => code
          case other => other.toString
        parts += s"NumFmt: $codeStr$idStr"

      // Alignment (if non-default)
      if s.align != Align.default then
        val wrapStr = if s.align.wrapText then ", wrap" else ""
        parts += s"Align: ${s.align.horizontal}, ${s.align.vertical}$wrapStr"

      val result = parts.result()
      if result.isEmpty then None
      else Some(result.map("  " + _).mkString("Style:\n", "\n", ""))
    }

  private def valueType(value: CellValue): String =
    value match
      case CellValue.Text(_) => "text"
      case CellValue.Number(_) => "number"
      case CellValue.Bool(_) => "boolean"
      case CellValue.DateTime(_) => "datetime"
      case CellValue.Error(_) => "error"
      case CellValue.RichText(_) => "richtext"
      case CellValue.Empty => "empty"
      case CellValue.Formula(_, _) => "formula"

  private def formatValue(value: CellValue): String =
    value match
      case CellValue.Text(s) => s"\"$s\""
      case CellValue.Number(n) =>
        if n.isWhole then n.toBigInt.toString
        else n.underlying.stripTrailingZeros.toPlainString
      case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
      case CellValue.DateTime(dt) => dt.toString
      case CellValue.Error(err) => err.toExcel
      case CellValue.RichText(rt) => s"\"${rt.toPlainText}\""
      case CellValue.Empty => "(empty)"
      case CellValue.Formula(expr, cached) =>
        val displayExpr = if expr.startsWith("=") then expr else s"=$expr"
        cached.map(formatValue).getOrElse(displayExpr)

  private def errorDetails(err: XLError): (String, String, String) =
    err match
      case XLError.InvalidCellRef(ref, reason) =>
        (
          "InvalidCellRef",
          s"'$ref' is not a valid cell reference: $reason",
          "Use A1-style references like A1, B5, or AA100"
        )
      case XLError.InvalidRange(range, reason) =>
        (
          "InvalidRange",
          s"'$range' is not a valid range: $reason",
          "Use ranges like A1:D10 or B5:B20"
        )
      case XLError.InvalidSheetName(name, reason) =>
        (
          "InvalidSheetName",
          s"'$name' is not a valid sheet name: $reason",
          "Sheet names cannot contain []:*?/\\"
        )
      case XLError.SheetNotFound(name) =>
        (
          "SheetNotFound",
          s"No sheet named '$name' exists",
          "Use 'xl sheets' to see available sheets"
        )
      case XLError.DuplicateSheet(name) =>
        ("DuplicateSheet", s"A sheet named '$name' already exists", "Choose a different name")
      case XLError.IOError(reason) =>
        ("IOError", reason, "Check file path and permissions")
      case XLError.ParseError(location, reason) =>
        ("ParseError", s"$reason at $location", "Check the syntax of your input")
      case XLError.TypeMismatch(expected, actual, ctx) =>
        ("TypeMismatch", s"Expected $expected but got $actual in $ctx", "Check the data types")
      case XLError.FormulaError(expr, reason) =>
        ("FormulaError", s"$reason in formula '$expr'", "Check the formula syntax")
      case other =>
        ("Error", other.message, "Check the documentation for more information")
