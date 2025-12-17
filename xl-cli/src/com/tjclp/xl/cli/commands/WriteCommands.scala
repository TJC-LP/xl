package com.tjclp.xl.cli.commands

import java.nio.file.Path

import cats.effect.IO
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{CellRange, Column, Row}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.helpers.{BatchParser, SheetResolver, StyleBuilder, ValueParser}
import com.tjclp.xl.cli.output.Format
import com.tjclp.xl.formula.{
  FormulaParser,
  FormulaPrinter,
  FormulaShifter,
  ParseError,
  SheetEvaluator
}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.sheets.styleSyntax
import com.tjclp.xl.styles.CellStyle

/**
 * Write command handlers.
 *
 * Commands that modify workbook data: put, putf, style, row, col, batch.
 */
object WriteCommands:

  /**
   * Write value to cell(s).
   */
  def put(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    refStr: String,
    valueStr: String,
    outputPath: Path
  ): IO[String] =
    for
      resolved <- SheetResolver.resolveRef(wb, sheetOpt, refStr, "put")
      (targetSheet, refOrRange) = resolved
      value = ValueParser.parseValue(valueStr)
      // Support both single cells and ranges
      (updatedSheet, cellCount) = refOrRange match
        case Left(ref) =>
          (targetSheet.put(ref, value), 1)
        case Right(range) =>
          // Fill all cells in range with same value
          val cells = range.cells.toList
          val sheet = cells.foldLeft(targetSheet)((s, ref) => s.put(ref, value))
          (sheet, cells.size)
      updatedWb = wb.put(updatedSheet)
      _ <- ExcelIO.instance[IO].write(updatedWb, outputPath)
    yield refOrRange match
      case Left(ref) => s"${Format.putSuccess(ref, value)}\nSaved: $outputPath"
      case Right(range) =>
        s"Applied value to $cellCount cells in ${range.toA1}\nSaved: $outputPath"

  /**
   * Write formula to cell(s).
   */
  def putFormula(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    refStr: String,
    formulaStr: String,
    outputPath: Path
  ): IO[String] =
    for
      resolved <- SheetResolver.resolveRef(wb, sheetOpt, refStr, "putf")
      (targetSheet, refOrRange) = resolved
      formula = if formulaStr.startsWith("=") then formulaStr.drop(1) else formulaStr
      // Parse formula to AST for dragging support
      fullFormula = s"=$formula"
      parsedExpr <- IO.fromEither(
        FormulaParser.parse(fullFormula).left.map { e =>
          new Exception(ParseError.formatWithContext(e, fullFormula))
        }
      )
      // Apply formula with Excel-style dragging
      (updatedSheet, cellCount) = refOrRange match
        case Left(ref) =>
          // Single cell: apply formula as-is, evaluate and cache result
          val cachedValue = SheetEvaluator.evaluateFormula(targetSheet)(fullFormula).toOption
          (targetSheet.put(ref, CellValue.Formula(formula, cachedValue)), 1)
        case Right(range) =>
          // Range: apply formula with shifting based on anchor modes
          val startRef = range.start
          val startCol = Column.index0(startRef.col)
          val startRow = Row.index0(startRef.row)
          val cells = range.cells.toList
          val sheet = cells.foldLeft(targetSheet) { (s, targetRef) =>
            val colDelta = Column.index0(targetRef.col) - startCol
            val rowDelta = Row.index0(targetRef.row) - startRow
            val shiftedExpr = FormulaShifter.shift(parsedExpr, colDelta, rowDelta)
            val shiftedFormula = FormulaPrinter.print(shiftedExpr, includeEquals = false)
            // Evaluate against current sheet state and cache result
            val fullShiftedFormula = s"=$shiftedFormula"
            val cachedValue = SheetEvaluator.evaluateFormula(s)(fullShiftedFormula).toOption
            s.put(targetRef, CellValue.Formula(shiftedFormula, cachedValue))
          }
          (sheet, cells.size)
      updatedWb = wb.put(updatedSheet)
      _ <- ExcelIO.instance[IO].write(updatedWb, outputPath)
    yield refOrRange match
      case Left(ref) =>
        s"${Format.putSuccess(ref, CellValue.Formula(formula))}\nSaved: $outputPath"
      case Right(range) =>
        s"Applied formula to $cellCount cells in ${range.toA1} (with anchor-aware dragging)\nSaved: $outputPath"

  /**
   * Apply styling to cells.
   *
   * By default, merges specified style properties with existing cell styles. Use replace=true to
   * replace the entire style instead of merging.
   */
  def style(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    rangeStr: String,
    bold: Boolean,
    italic: Boolean,
    underline: Boolean,
    bg: Option[String],
    fg: Option[String],
    fontSize: Option[Double],
    fontName: Option[String],
    align: Option[String],
    valign: Option[String],
    wrap: Boolean,
    numFormat: Option[String],
    border: Option[String],
    borderTop: Option[String],
    borderRight: Option[String],
    borderBottom: Option[String],
    borderLeft: Option[String],
    borderColor: Option[String],
    replace: Boolean,
    outputPath: Path
  ): IO[String] =
    for
      resolved <- SheetResolver.resolveRef(wb, sheetOpt, rangeStr, "style")
      (targetSheet, refOrRange) = resolved
      range = refOrRange match
        case Right(r) => r
        case Left(ref) => CellRange(ref, ref)
      newStyle <- StyleBuilder.buildCellStyle(
        bold,
        italic,
        underline,
        bg,
        fg,
        fontSize,
        fontName,
        align,
        valign,
        wrap,
        numFormat,
        border,
        borderTop,
        borderRight,
        borderBottom,
        borderLeft,
        borderColor
      )
      // Apply style to each cell, either replacing or merging
      updatedSheet =
        if replace then
          // Replace mode: apply style uniformly to all cells
          styleSyntax.withRangeStyle(targetSheet)(range, newStyle)
        else
          // Merge mode: merge with each cell's existing style
          range.cells.foldLeft(targetSheet) { (sheet, ref) =>
            val existingStyle = sheet.cells
              .get(ref)
              .flatMap(_.styleId)
              .flatMap(sheet.styleRegistry.get)
              .getOrElse(CellStyle.default)
            val mergedStyle = StyleBuilder.mergeStyles(existingStyle, newStyle)
            styleSyntax.withRangeStyle(sheet)(CellRange(ref, ref), mergedStyle)
          }
      updatedWb = wb.put(updatedSheet)
      _ <- ExcelIO.instance[IO].write(updatedWb, outputPath)
      appliedList = StyleBuilder.buildStyleDescription(
        bold,
        italic,
        underline,
        bg,
        fg,
        fontSize,
        fontName,
        align,
        valign,
        wrap,
        numFormat,
        border
      )
      modeLabel = if replace then " (replace)" else " (merge)"
    yield s"Styled: ${range.toA1}$modeLabel\nApplied: ${appliedList.mkString(", ")}\nSaved: $outputPath"

  /**
   * Set row properties (height, hide/show).
   */
  def row(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    rowNum: Int,
    height: Option[Double],
    hide: Boolean,
    show: Boolean,
    outputPath: Path
  ): IO[String] =
    SheetResolver.requireSheet(wb, sheetOpt, "row").flatMap { sheet =>
      val rowRef = Row.from1(rowNum)
      val currentProps = sheet.getRowProperties(rowRef)
      val newProps = currentProps.copy(
        height = height.orElse(currentProps.height),
        hidden = if hide then true else if show then false else currentProps.hidden
      )
      val updatedSheet = sheet.setRowProperties(rowRef, newProps)
      val updatedWb = wb.put(updatedSheet)
      ExcelIO.instance[IO].write(updatedWb, outputPath).map { _ =>
        val changes = List(
          height.map(h => s"height=$h"),
          if hide then Some("hidden=true") else None,
          if show then Some("hidden=false") else None
        ).flatten
        s"Row $rowNum: ${changes.mkString(", ")}\nSaved: $outputPath"
      }
    }

  /**
   * Set column properties (width, hide/show).
   */
  def col(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    colStr: String,
    width: Option[Double],
    hide: Boolean,
    show: Boolean,
    outputPath: Path
  ): IO[String] =
    SheetResolver.requireSheet(wb, sheetOpt, "col").flatMap { sheet =>
      IO.fromEither(Column.fromLetter(colStr).left.map(e => new Exception(e))).flatMap { colRef =>
        val currentProps = sheet.getColumnProperties(colRef)
        val newProps = currentProps.copy(
          width = width.orElse(currentProps.width),
          hidden = if hide then true else if show then false else currentProps.hidden
        )
        val updatedSheet = sheet.setColumnProperties(colRef, newProps)
        val updatedWb = wb.put(updatedSheet)
        ExcelIO.instance[IO].write(updatedWb, outputPath).map { _ =>
          val changes = List(
            width.map(w => s"width=$w"),
            if hide then Some("hidden=true") else None,
            if show then Some("hidden=false") else None
          ).flatten
          s"Column $colStr: ${changes.mkString(", ")}\nSaved: $outputPath"
        }
      }
    }

  /**
   * Apply multiple operations atomically (JSON from stdin or file).
   */
  def batch(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    source: String,
    outputPath: Path
  ): IO[String] =
    BatchParser.readBatchInput(source).flatMap { input =>
      BatchParser.parseBatchOperations(input).flatMap { ops =>
        BatchParser.applyBatchOperations(wb, sheetOpt, ops).flatMap { updatedWb =>
          ExcelIO.instance[IO].write(updatedWb, outputPath).map { _ =>
            val summary = ops
              .map {
                case BatchParser.BatchOp.Put(ref, value) => s"  PUT $ref = $value"
                case BatchParser.BatchOp.PutFormula(ref, formula) => s"  PUTF $ref = $formula"
              }
              .mkString("\n")
            s"Applied ${ops.size} operations:\n$summary\nSaved: $outputPath"
          }
        }
      }
    }
