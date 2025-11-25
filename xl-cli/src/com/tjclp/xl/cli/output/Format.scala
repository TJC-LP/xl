package com.tjclp.xl.cli.output

import com.tjclp.xl.error.XLError
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.addressing.ARef

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
   * Format cell info output.
   */
  def cellInfo(
    ref: ARef,
    value: CellValue,
    formatted: String,
    dependencies: Vector[ARef],
    dependents: Vector[ARef]
  ): String =
    val typeStr = valueType(value)
    val depsStr =
      if dependencies.isEmpty then "(none)"
      else dependencies.map(_.toA1).mkString(", ")
    val deptsStr =
      if dependents.isEmpty then "(none)"
      else dependents.map(_.toA1).mkString(", ")

    val formulaLine = value match
      case CellValue.Formula(expr, _) => s"Formula: =$expr\n"
      case _ => ""

    s"""Cell: ${ref.toA1}
       |Type: $typeStr
       |${formulaLine}Value: $formatted
       |Dependencies: $depsStr
       |Dependents: $deptsStr""".stripMargin

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
        cached.map(formatValue).getOrElse(s"=$expr")

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
