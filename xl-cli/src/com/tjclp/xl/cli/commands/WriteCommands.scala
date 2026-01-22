package com.tjclp.xl.cli.commands

import java.nio.file.Path

import cats.effect.IO
import cats.implicits.*
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
  SheetEvaluator,
  TExpr
}
import com.tjclp.xl.formula.eval.DependentRecalculation.*
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.sheets.styleSyntax
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.cli.{FillDirection, SortDirection, SortKey, SortMode}

/**
 * Write command handlers.
 *
 * Commands that modify workbook data: put, putf, style, row, col, batch.
 */
object WriteCommands:

  /**
   * Validate that the count of values/formulas matches the cell count in a range.
   *
   * Returns None if valid, Some(Exception) if counts mismatch with actionable error message.
   */
  private def validateCountMatch(
    command: String,
    range: CellRange,
    cellCount: Int,
    providedCount: Int,
    itemType: String
  ): Option[Exception] =
    if cellCount == providedCount then None
    else
      val suggestion =
        if providedCount == 1 then
          s"Hint: Use 1 $itemType for fill pattern: `$command ${range.toA1} <$itemType>`"
        else if providedCount < cellCount then
          s"Hint: Provide exactly $cellCount ${itemType}s to match range size."
        else s"Hint: Range has $cellCount cells but you provided $providedCount ${itemType}s."

      Some(
        new Exception(
          s"Range ${range.toA1} has $cellCount cells but $providedCount ${itemType}s provided. " +
            suggestion
        )
      )

  /**
   * Write value(s) to cell(s).
   *
   * Supports three modes:
   *   1. Single cell: put A1 100
   *   2. Fill pattern: put A1:A10 100 (all cells get same value)
   *   3. Batch values: put A1:A3 1 2 3 (values map 1:1, row-major)
   */
  def put(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    refStr: String,
    values: List[String],
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    for
      resolved <- SheetResolver.resolveRef(wb, sheetOpt, refStr, "put")
      (targetSheet, refOrRange) = resolved

      // Determine mode based on ref type and value count
      result <- (refOrRange, values) match
        case (Left(ref), List(singleValue)) =>
          // Mode 1: Single cell
          putSingleCell(wb, targetSheet, ref, singleValue, outputPath, config)

        case (Right(range), List(singleValue)) =>
          // Mode 2: Fill pattern
          putFillPattern(wb, targetSheet, range, singleValue, outputPath, config)

        case (Right(range), multipleValues @ (_ :: _ :: _)) =>
          // Mode 3: Batch values (validate count matches) - 2+ values
          putBatchValues(wb, targetSheet, range, multipleValues, outputPath, config)

        case (Left(ref), multipleValues @ (_ :: _)) =>
          IO.raiseError(
            new Exception(
              s"Cannot put ${multipleValues.length} values to single cell ${ref.toA1}. " +
                "Use a range (e.g., A1:A3) for batch operations."
            )
          )

        case (_, Nil) =>
          IO.raiseError(new Exception("No values provided for put command"))
    yield result

  /** Mode 1: Put single value to single cell */
  private def putSingleCell(
    wb: Workbook,
    sheet: Sheet,
    ref: com.tjclp.xl.addressing.ARef,
    valueStr: String,
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    val value = ValueParser.parseValue(valueStr)
    val updatedSheet = sheet.put(ref, value)
    val updatedWb = wb.put(updatedSheet).recalculateDependents(sheet.name, Set(ref))
    ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config).map { _ =>
      s"${Format.putSuccess(ref, value)}\nSaved: $outputPath"
    }

  /** Mode 2: Fill all cells in range with same value */
  private def putFillPattern(
    wb: Workbook,
    sheet: Sheet,
    range: CellRange,
    valueStr: String,
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    val value = ValueParser.parseValue(valueStr)
    val cellCount = range.cellCount
    val modifiedRefs = range.cells.toSet
    val updatedSheet = range.cells.foldLeft(sheet)((s, ref) => s.put(ref, value))
    val updatedWb = wb.put(updatedSheet).recalculateDependents(sheet.name, modifiedRefs)
    ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config).map { _ =>
      s"Filled $cellCount cells in ${range.toA1} with value ${value}\nSaved: $outputPath"
    }

  /** Mode 3: Put different values to each cell (row-major order) */
  private def putBatchValues(
    wb: Workbook,
    sheet: Sheet,
    range: CellRange,
    values: List[String],
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    val cellCount = range.cellCount.toInt
    validateCountMatch("put", range, cellCount, values.length, "value") match
      case Some(error) => IO.raiseError(error)
      case None =>
        val updates = range.cellsRowMajor
          .zip(values.iterator)
          .map { (ref, valueStr) =>
            (ref, ValueParser.parseValue(valueStr))
          }
          .toVector
        val modifiedRefs = updates.map(_._1).toSet
        val updatedSheet = sheet.put(updates*)
        val updatedWb = wb.put(updatedSheet).recalculateDependents(sheet.name, modifiedRefs)
        ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config).map { _ =>
          s"Put ${values.length} values to ${range.toA1} (row-major)\nSaved: $outputPath"
        }

  /**
   * Write formula to cell(s).
   */
  def putFormula(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    refStr: String,
    formulas: List[String],
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    for
      resolved <- SheetResolver.resolveRef(wb, sheetOpt, refStr, "putf")
      (targetSheet, refOrRange) = resolved

      // Determine mode based on ref type and formula count
      result <- (refOrRange, formulas) match
        case (Left(ref), List(singleFormula)) =>
          // Mode 1: Single cell
          putfSingleCell(wb, targetSheet, ref, singleFormula, outputPath, config)

        case (Right(range), List(singleFormula)) =>
          // Mode 2: Formula dragging (existing behavior with $ anchors)
          putfFormulaDragging(wb, targetSheet, range, singleFormula, outputPath, config)

        case (Right(range), multipleFormulas @ (_ :: _ :: _)) =>
          // Mode 3: Batch formulas (no dragging, apply as-is) - 2+ formulas
          putfBatchFormulas(wb, targetSheet, range, multipleFormulas, outputPath, config)

        case (Left(ref), multipleFormulas @ (_ :: _)) =>
          IO.raiseError(
            new Exception(
              s"Cannot put ${multipleFormulas.length} formulas to single cell ${ref.toA1}. " +
                "Use a range (e.g., B1:B3) for batch formula operations."
            )
          )

        case (_, Nil) =>
          IO.raiseError(new Exception("No formulas provided for putf command"))
    yield result

  /** Mode 1: Put single formula to single cell */
  private def putfSingleCell(
    wb: Workbook,
    sheet: Sheet,
    ref: com.tjclp.xl.addressing.ARef,
    formulaStr: String,
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    val formula = if formulaStr.startsWith("=") then formulaStr.drop(1) else formulaStr
    val fullFormula = s"=$formula"
    for
      parsedExpr <- IO.fromEither(
        FormulaParser.parse(fullFormula).left.map { e =>
          new Exception(ParseError.formatWithContext(e, fullFormula))
        }
      )
      cachedValue = SheetEvaluator.evaluateFormula(sheet)(fullFormula, workbook = Some(wb)).toOption
      sheetWithFormula = sheet.put(ref, CellValue.Formula(formula, cachedValue))
      // Auto-apply date format if formula involves date functions
      finalSheet =
        if TExpr.containsDateFunction(parsedExpr) then
          val numFmt =
            if TExpr.containsTimeFunction(parsedExpr) then NumFmt.DateTime else NumFmt.Date
          val existingStyle = sheetWithFormula.cells
            .get(ref)
            .flatMap(_.styleId)
            .flatMap(sheetWithFormula.styleRegistry.get)
            .getOrElse(CellStyle.default)
          val mergedStyle = existingStyle.withNumFmt(numFmt)
          styleSyntax.withRangeStyle(sheetWithFormula)(CellRange(ref, ref), mergedStyle)
        else sheetWithFormula
      updatedWb = wb.put(finalSheet).recalculateDependents(sheet.name, Set(ref))
      _ <- ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config)
    yield s"${Format.putSuccess(ref, CellValue.Formula(formula))}\nSaved: $outputPath"

  /** Mode 2: Formula dragging with anchor-aware shifting (existing behavior) */
  private def putfFormulaDragging(
    wb: Workbook,
    sheet: Sheet,
    range: CellRange,
    formulaStr: String,
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    val formula = if formulaStr.startsWith("=") then formulaStr.drop(1) else formulaStr
    val fullFormula = s"=$formula"
    for
      parsedExpr <- IO.fromEither(
        FormulaParser.parse(fullFormula).left.map { e =>
          new Exception(ParseError.formatWithContext(e, fullFormula))
        }
      )
      // Apply formula with Excel-style dragging (existing logic)
      updatedSheet = putfDraggingLogic(sheet, wb, range, formula, parsedExpr)
      modifiedRefs = range.cells.toSet
      updatedWb = wb.put(updatedSheet).recalculateDependents(sheet.name, modifiedRefs)
      _ <- ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config)
      cellCount = range.cellCount
    yield s"Applied formula to $cellCount cells in ${range.toA1} (with anchor-aware dragging)\nSaved: $outputPath"

  /** Mode 3: Batch formulas (no dragging, apply as-is) */
  private def putfBatchFormulas(
    wb: Workbook,
    sheet: Sheet,
    range: CellRange,
    formulas: List[String],
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    val cellCount = range.cellCount.toInt
    validateCountMatch("putf", range, cellCount, formulas.length, "formula") match
      case Some(error) => IO.raiseError(error)
      case None =>
        for
          // Parse and apply each formula (using iterator to avoid toList)
          updates <- range.cellsRowMajor.zip(formulas.iterator).toList.traverse {
            (ref, formulaStr) =>
              val formula = if formulaStr.startsWith("=") then formulaStr.drop(1) else formulaStr
              val fullFormula = s"=$formula"
              IO.fromEither(
                FormulaParser.parse(fullFormula).left.map { e =>
                  new Exception(
                    s"Formula for ${ref.toA1}: ${ParseError.formatWithContext(e, fullFormula)}"
                  )
                }
              ).map { _ =>
                val cachedValue =
                  SheetEvaluator.evaluateFormula(sheet)(fullFormula, workbook = Some(wb)).toOption
                (ref, CellValue.Formula(formula, cachedValue))
              }
          }
          modifiedRefs = updates.map(_._1).toSet
          updatedSheet = sheet.put(updates*)
          updatedWb = wb.put(updatedSheet).recalculateDependents(sheet.name, modifiedRefs)
          _ <- ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config)
        yield s"Put ${formulas.length} formulas to ${range.toA1} (explicit, no dragging)\nSaved: $outputPath"

  /** Helper: Apply formula dragging logic (extracted from original putFormula) */
  private def putfDraggingLogic(
    sheet: Sheet,
    wb: Workbook,
    range: CellRange,
    formula: String,
    parsedExpr: TExpr[?]
  ): Sheet =
    val startRef = range.start
    val startCol = Column.index0(startRef.col)
    val startRow = Row.index0(startRef.row)
    val cells = range.cells.toList
    // Apply formula with shifting (existing logic)
    val sheetWithFormulas = cells.foldLeft(sheet) { (s, targetRef) =>
      val colDelta = Column.index0(targetRef.col) - startCol
      val rowDelta = Row.index0(targetRef.row) - startRow
      val shiftedExpr = FormulaShifter.shift(parsedExpr, colDelta, rowDelta)
      val shiftedFormula = FormulaPrinter.print(shiftedExpr, includeEquals = false)
      val fullShiftedFormula = s"=$shiftedFormula"
      val cachedValue =
        SheetEvaluator.evaluateFormula(s)(fullShiftedFormula, workbook = Some(wb)).toOption
      s.put(targetRef, CellValue.Formula(shiftedFormula, cachedValue))
    }
    // Auto-apply date format if needed
    if TExpr.containsDateFunction(parsedExpr) then
      val numFmt = if TExpr.containsTimeFunction(parsedExpr) then NumFmt.DateTime else NumFmt.Date
      cells.foldLeft(sheetWithFormulas) { (s, cellRef) =>
        val existingStyle = s.cells
          .get(cellRef)
          .flatMap(_.styleId)
          .flatMap(s.styleRegistry.get)
          .getOrElse(CellStyle.default)
        val mergedStyle = existingStyle.withNumFmt(numFmt)
        styleSyntax.withRangeStyle(s)(CellRange(cellRef, cellRef), mergedStyle)
      }
    else sheetWithFormulas

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
    outputPath: Path,
    config: WriterConfig
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
      _ <- ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config)
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
    outputPath: Path,
    config: WriterConfig
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
      ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config).map { _ =>
        val changes = List(
          height.map(h => s"height=$h"),
          if hide then Some("hidden=true") else None,
          if show then Some("hidden=false") else None
        ).flatten
        s"Row $rowNum: ${changes.mkString(", ")}\nSaved: $outputPath"
      }
    }

  /**
   * Set column properties (width, hide/show, auto-fit). Supports single columns (A) or ranges
   * (A:F).
   */
  def col(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    colStr: String,
    width: Option[Double],
    hide: Boolean,
    show: Boolean,
    autoFit: Boolean,
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    SheetResolver.requireSheet(wb, sheetOpt, "col").flatMap { sheet =>
      // Try parsing as column range (A:F) first, then single column (A)
      parseColumnSpec(colStr) match
        case Left(err) => IO.raiseError(new Exception(err))
        case Right(columns) =>
          val (updatedSheet, results) = columns.foldLeft((sheet, List.empty[String])) {
            case ((s, msgs), colRef) =>
              val currentProps = s.getColumnProperties(colRef)
              val effectiveWidth: Option[Double] =
                if autoFit then Some(calculateAutoFitWidth(s, colRef))
                else width
              val newProps = currentProps.copy(
                width = effectiveWidth.orElse(currentProps.width),
                hidden = if hide then true else if show then false else currentProps.hidden
              )
              val newSheet = s.setColumnProperties(colRef, newProps)
              val changes = List(
                effectiveWidth.map(w => f"$w%.2f${if autoFit then " (auto)" else ""}"),
                if hide then Some("hide") else None,
                if show then Some("show") else None
              ).flatten
              val msg = s"${colRef.toLetter}: ${changes.mkString(", ")}"
              (newSheet, msgs :+ msg)
          }
          val updatedWb = wb.put(updatedSheet)
          ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config).map { _ =>
            results match
              case single :: Nil => s"Column $single\nSaved: $outputPath"
              case multiple =>
                s"Columns:\n${multiple.map("  " + _).mkString("\n")}\nSaved: $outputPath"
          }
    }

  /**
   * Auto-fit all columns (or specified range) based on content.
   */
  def autoFit(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    columnsOpt: Option[String],
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    SheetResolver.requireSheet(wb, sheetOpt, "autofit").flatMap { sheet =>
      // Determine columns to auto-fit
      val columnsResult: Either[String, List[Column]] = columnsOpt match
        case Some(spec) => parseColumnSpec(spec)
        case None =>
          // Auto-fit all used columns
          sheet.usedRange match
            case Some(range) =>
              val cols = (range.colStart.index0 to range.colEnd.index0).map(Column.from0).toList
              Right(cols)
            case None => Right(List.empty) // Empty sheet

      columnsResult match
        case Left(err) => IO.raiseError(new Exception(err))
        case Right(columns) if columns.isEmpty =>
          IO.pure(s"No columns to auto-fit (empty sheet)\nSaved: $outputPath")
        case Right(columns) =>
          val (updatedSheet, widths) = columns.foldLeft((sheet, List.empty[(Column, Double)])) {
            case ((s, ws), colRef) =>
              val w = calculateAutoFitWidth(s, colRef)
              val currentProps = s.getColumnProperties(colRef)
              val newProps = currentProps.copy(width = Some(w))
              (s.setColumnProperties(colRef, newProps), ws :+ (colRef, w))
          }
          val updatedWb = wb.put(updatedSheet)
          ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config).map { _ =>
            val summary = widths.map { case (c, w) => f"${c.toLetter}: $w%.2f" }.mkString(", ")
            s"Auto-fit ${columns.size} column(s): $summary\nSaved: $outputPath"
          }
    }

  /** Parse column spec: single column (A) or range (A:F) */
  private def parseColumnSpec(spec: String): Either[String, List[Column]] =
    if spec.contains(':') then
      // Column range like A:F
      CellRange.parse(spec).map { range =>
        (range.colStart.index0 to range.colEnd.index0).map(Column.from0).toList
      }
    else
      // Single column
      Column.fromLetter(spec).map(c => List(c))

  /**
   * Calculate optimal column width based on cell content. Returns width in Excel character units.
   * Uses a multiplier approach calibrated against Excel's actual autofit behavior.
   */
  private def calculateAutoFitWidth(sheet: Sheet, col: Column): Double =
    val cellsInColumn = sheet.cells.filter { case (ref, _) => ref.col == col }

    if cellsInColumn.isEmpty then 8.43 // Default Excel column width
    else
      val maxCharWidth = cellsInColumn.values
        .map { cell =>
          estimateCellWidth(cell.value)
        }
        .maxOption
        .getOrElse(0)

      // Use multiplier + small padding (calibrated to match Excel's font-metric-aware autofit)
      // Excel uses ~0.85x char count for proportional fonts (Calibri 11pt default)
      val width =
        BigDecimal(maxCharWidth * 0.85 + 1.5).setScale(2, BigDecimal.RoundingMode.HALF_UP).toDouble
      math.max(width, 5.0) // Allow narrower columns like Excel does

  /**
   * Estimate the display width of a cell value in characters.
   */
  private def estimateCellWidth(value: CellValue): Int =
    import com.tjclp.xl.cells.CellValue.*
    value match
      case Text(s) => s.length
      case Number(n) => formatNumber(n).length
      case Bool(b) => if b then 4 else 5 // "TRUE" or "FALSE"
      case Error(e) => e.toString.length
      case Empty => 0
      case DateTime(dt) => dt.toString.length // ISO format
      case Formula(_, Some(cached)) => estimateCellWidth(cached)
      case Formula(expr, None) => expr.length
      case RichText(rt) => rt.toPlainText.length

  /** Format number for display width estimation */
  private def formatNumber(n: BigDecimal): String =
    if n.isWhole then n.toBigInt.toString
    else n.underlying.stripTrailingZeros.toPlainString

  /**
   * Apply multiple operations atomically (JSON from stdin or file).
   */
  def batch(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    source: String,
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    BatchParser.readBatchInput(source).flatMap { input =>
      BatchParser.parseBatchOperations(input).flatMap { ops =>
        BatchParser.applyBatchOperations(wb, sheetOpt, ops).flatMap { updatedWb =>
          ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config).map { _ =>
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

  /**
   * Fill cells with source value/formula (Excel Ctrl+D/Ctrl+R).
   *
   * Fill direction controls how source cells map to target cells:
   *   - Down: Source row is repeated down (A1 → A1:A10, or A1:E1 → A1:E10)
   *   - Right: Source column is repeated right (A1 → A1:J1, or A1:A5 → A1:J5)
   *
   * Formulas are shifted relative to the source position using Excel's anchor rules.
   */
  def fill(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    sourceStr: String,
    targetStr: String,
    direction: FillDirection,
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    for
      // Resolve source and target
      sourceResolved <- SheetResolver.resolveRef(wb, sheetOpt, sourceStr, "fill")
      targetResolved <- SheetResolver.resolveRef(wb, sheetOpt, targetStr, "fill")
      (targetSheet, _) = targetResolved

      // Convert source to range (single cell becomes 1x1 range)
      sourceRange = sourceResolved._2 match
        case Left(ref) => CellRange(ref, ref)
        case Right(range) => range

      // Target must be a range
      targetRange <- targetResolved._2 match
        case Right(range) => IO.pure(range)
        case Left(ref) =>
          IO.raiseError(
            new Exception(s"Fill target must be a range, not a single cell: ${ref.toA1}")
          )

      // Validate source/target compatibility based on direction
      _ <- validateFillRanges(sourceRange, targetRange, direction)

      // Apply fill operation
      updatedSheet = applyFill(targetSheet, wb, sourceRange, targetRange, direction)
      modifiedRefs = targetRange.cells.toSet
      updatedWb = wb.put(updatedSheet).recalculateDependents(targetSheet.name, modifiedRefs)
      _ <- ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config)
      dirLabel = if direction == FillDirection.Right then "right" else "down"
    yield s"Filled ${targetRange.toA1} from ${sourceRange.toA1} ($dirLabel)\nSaved: $outputPath"

  /** Validate that source and target ranges are compatible for fill direction */
  private def validateFillRanges(
    source: CellRange,
    target: CellRange,
    direction: FillDirection
  ): IO[Unit] =
    direction match
      case FillDirection.Down =>
        // For fill down, source columns must match target columns
        val sourceStartCol = Column.index0(source.start.col)
        val sourceEndCol = Column.index0(source.end.col)
        val targetStartCol = Column.index0(target.start.col)
        val targetEndCol = Column.index0(target.end.col)
        if sourceStartCol == targetStartCol && sourceEndCol == targetEndCol then IO.unit
        else
          IO.raiseError(
            new Exception(
              s"Fill down requires matching columns. Source: ${source.toA1}, Target: ${target.toA1}"
            )
          )

      case FillDirection.Right =>
        // For fill right, source rows must match target rows
        val sourceStartRow = Row.index0(source.start.row)
        val sourceEndRow = Row.index0(source.end.row)
        val targetStartRow = Row.index0(target.start.row)
        val targetEndRow = Row.index0(target.end.row)
        if sourceStartRow == targetStartRow && sourceEndRow == targetEndRow then IO.unit
        else
          IO.raiseError(
            new Exception(
              s"Fill right requires matching rows. Source: ${source.toA1}, Target: ${target.toA1}"
            )
          )

  /** Apply fill operation by copying source cells to target range with formula shifting */
  private def applyFill(
    sheet: Sheet,
    wb: Workbook,
    source: CellRange,
    target: CellRange,
    direction: FillDirection
  ): Sheet =
    direction match
      case FillDirection.Down =>
        applyFillDown(sheet, wb, source, target)
      case FillDirection.Right =>
        applyFillRight(sheet, wb, source, target)

  /** Fill down: repeat source row(s) down through target range */
  private def applyFillDown(
    sheet: Sheet,
    wb: Workbook,
    source: CellRange,
    target: CellRange
  ): Sheet =
    val sourceStartRow = Row.index0(source.start.row)
    val sourceEndRow = Row.index0(source.end.row)
    val sourceRowCount = sourceEndRow - sourceStartRow + 1

    val targetStartRow = Row.index0(target.start.row)
    val targetEndRow = Row.index0(target.end.row)

    val startCol = Column.index0(source.start.col)
    val endCol = Column.index0(source.end.col)

    // For each target row, copy from corresponding source row (cycling if needed)
    (targetStartRow to targetEndRow).foldLeft(sheet) { (s, targetRowIdx) =>
      // Determine which source row to copy from (0-indexed within source)
      val sourceRowOffset = (targetRowIdx - targetStartRow) % sourceRowCount
      val sourceRowIdx = sourceStartRow + sourceRowOffset.toInt
      val rowDelta = targetRowIdx - sourceRowIdx

      // Skip if we're on a source row (no shifting needed for source itself)
      if rowDelta == 0 then s
      else
        // Copy each cell in the row
        (startCol to endCol).foldLeft(s) { (s2, colIdx) =>
          // ARef.from0 takes (colIndex, rowIndex)
          val sourceRef = com.tjclp.xl.addressing.ARef.from0(colIdx, sourceRowIdx)
          val targetRef = com.tjclp.xl.addressing.ARef.from0(colIdx, targetRowIdx)
          copyCell(s2, wb, sourceRef, targetRef, colDelta = 0, rowDelta = rowDelta.toInt)
        }
    }

  /** Fill right: repeat source column(s) right through target range */
  private def applyFillRight(
    sheet: Sheet,
    wb: Workbook,
    source: CellRange,
    target: CellRange
  ): Sheet =
    val sourceStartCol = Column.index0(source.start.col)
    val sourceEndCol = Column.index0(source.end.col)
    val sourceColCount = sourceEndCol - sourceStartCol + 1

    val targetStartCol = Column.index0(target.start.col)
    val targetEndCol = Column.index0(target.end.col)

    val startRow = Row.index0(source.start.row)
    val endRow = Row.index0(source.end.row)

    // For each target column, copy from corresponding source column (cycling if needed)
    (targetStartCol to targetEndCol).foldLeft(sheet) { (s, targetColIdx) =>
      // Determine which source column to copy from (0-indexed within source)
      val sourceColOffset = (targetColIdx - targetStartCol) % sourceColCount
      val sourceColIdx = sourceStartCol + sourceColOffset.toInt
      val colDelta = targetColIdx - sourceColIdx

      // Skip if we're on a source column (no shifting needed for source itself)
      if colDelta == 0 then s
      else
        // Copy each cell in the column
        (startRow to endRow).foldLeft(s) { (s2, rowIdx) =>
          // ARef.from0 takes (colIndex, rowIndex)
          val sourceRef = com.tjclp.xl.addressing.ARef.from0(sourceColIdx, rowIdx)
          val targetRef = com.tjclp.xl.addressing.ARef.from0(targetColIdx, rowIdx)
          copyCell(s2, wb, sourceRef, targetRef, colDelta = colDelta.toInt, rowDelta = 0)
        }
    }

  /** Copy a single cell with formula shifting */
  private def copyCell(
    sheet: Sheet,
    wb: Workbook,
    sourceRef: com.tjclp.xl.addressing.ARef,
    targetRef: com.tjclp.xl.addressing.ARef,
    colDelta: Int,
    rowDelta: Int
  ): Sheet =
    sheet.cells.get(sourceRef) match
      case None => sheet // Empty source cell, nothing to copy
      case Some(sourceCell) =>
        sourceCell.value match
          case CellValue.Formula(formula, _) =>
            // Shift formula references
            val fullFormula = s"=$formula"
            FormulaParser.parse(fullFormula) match
              case Left(_) =>
                // If formula can't be parsed, copy as-is
                val cachedValue =
                  SheetEvaluator.evaluateFormula(sheet)(fullFormula, workbook = Some(wb)).toOption
                sheet.put(targetRef, CellValue.Formula(formula, cachedValue))
              case Right(parsedExpr) =>
                val shiftedExpr = FormulaShifter.shift(parsedExpr, colDelta, rowDelta)
                val shiftedFormula = FormulaPrinter.print(shiftedExpr, includeEquals = false)
                val fullShiftedFormula = s"=$shiftedFormula"
                val cachedValue =
                  SheetEvaluator
                    .evaluateFormula(sheet)(fullShiftedFormula, workbook = Some(wb))
                    .toOption
                sheet.put(targetRef, CellValue.Formula(shiftedFormula, cachedValue))

          case value =>
            // Non-formula: copy value as-is
            sheet.put(targetRef, value)

  /**
   * Sort rows in a range by specified column(s).
   *
   * Sorts rows in-place while preserving:
   *   - Cell styles (styles move with their rows)
   *   - Comments (move with cells)
   *   - Formulas (relative references adjusted to match new row position)
   *
   * Note: Only cells within the specified range columns are moved. Cells outside the range (e.g.,
   * column C when sorting A:B) are not affected. This matches Excel's behavior for range-based
   * sorting.
   *
   * @param sortKeys
   *   Sort criteria (column, direction, mode) - at least one required
   * @param hasHeader
   *   If true, first row is excluded from sort
   */
  def sort(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    rangeStr: String,
    sortKeys: List[SortKey],
    hasHeader: Boolean,
    outputPath: Path,
    config: WriterConfig
  ): IO[String] =
    for
      resolved <- SheetResolver.resolveRef(wb, sheetOpt, rangeStr, "sort")
      (targetSheet, refOrRange) = resolved

      // Sort requires a range, not single cell
      range <- refOrRange match
        case Right(r) => IO.pure(r)
        case Left(ref) =>
          IO.raiseError(
            new Exception(s"sort requires a range, not single cell: ${ref.toA1}")
          )

      // Validate sort columns are within range
      _ <- validateSortColumns(range, sortKeys)

      // Perform sort
      sortedSheet = applySortToRange(targetSheet, range, sortKeys, hasHeader)
      updatedWb = wb.put(sortedSheet)
      _ <- ExcelIO.instance[IO].writeWith(updatedWb, outputPath, config)

      // Build result message
      keyDesc = sortKeys
        .map(k =>
          s"${k.column} ${if k.direction == SortDirection.Descending then "desc" else "asc"}"
        )
        .mkString(", ")
      headerNote = if hasHeader then " (header row preserved)" else ""
    yield s"Sorted ${range.toA1} by $keyDesc$headerNote\nSaved: $outputPath"

  /** Validate that all sort columns are within the range */
  private def validateSortColumns(range: CellRange, keys: List[SortKey]): IO[Unit] =
    val rangeColStart = Column.index0(range.colStart)
    val rangeColEnd = Column.index0(range.colEnd)

    keys.traverse_ { key =>
      Column.fromLetter(key.column) match
        case Left(err) => IO.raiseError(new Exception(s"Invalid column: ${key.column}"))
        case Right(col) =>
          val colIdx = Column.index0(col)
          if colIdx < rangeColStart || colIdx > rangeColEnd then
            IO.raiseError(
              new Exception(
                s"Sort column ${key.column} is outside range ${range.toA1}"
              )
            )
          else IO.unit
    }

  /** Apply sort to the specified range */
  private def applySortToRange(
    sheet: Sheet,
    range: CellRange,
    sortKeys: List[SortKey],
    hasHeader: Boolean
  ): Sheet =
    val rowStart = Row.index0(range.rowStart)
    val rowEnd = Row.index0(range.rowEnd)
    val colStart = Column.index0(range.colStart)
    val colEnd = Column.index0(range.colEnd)

    // Determine which rows to sort (skip header if specified)
    val dataRowStart = if hasHeader then rowStart + 1 else rowStart

    // If only header row or empty, nothing to sort
    if dataRowStart > rowEnd then sheet
    else
      // Extract row data as Map[colIndex -> Cell] for O(1) lookup during comparison
      val rowsToSort: Vector[(Int, Map[Int, Cell])] =
        (dataRowStart to rowEnd).map { rowIdx =>
          val cellMap = (colStart to colEnd).flatMap { colIdx =>
            val ref = com.tjclp.xl.addressing.ARef.from0(colIdx, rowIdx)
            sheet.cells.get(ref).map(colIdx -> _)
          }.toMap
          (rowIdx, cellMap)
        }.toVector

      // Sort rows using comparison function
      val comparator = buildRowComparator(sortKeys)
      val sortedRows = rowsToSort.sortWith { (a, b) =>
        comparator(a._2, b._2) < 0
      }

      // Reconstruct sheet with sorted rows
      reconstructWithSortedRows(sheet, range, dataRowStart, sortedRows)

  /** Build a comparator for rows based on sort keys. Uses Map for O(1) cell lookup. */
  private def buildRowComparator(
    sortKeys: List[SortKey]
  ): (Map[Int, Cell], Map[Int, Cell]) => Int =
    (rowA, rowB) =>
      // Find first non-equal comparison
      sortKeys.iterator
        .flatMap { key =>
          // Column was already validated, so this should always succeed
          Column.fromLetter(key.column).toOption.map { col =>
            val colIdx = col.index0
            val cellA = rowA.get(colIdx) // O(1) lookup
            val cellB = rowB.get(colIdx) // O(1) lookup

            val valueA = cellA.map(c => getSortableValue(c.value, key.mode))
            val valueB = cellB.map(c => getSortableValue(c.value, key.mode))

            val cmp = compareSortValues(valueA, valueB, key.mode)

            // Apply direction
            if key.direction == SortDirection.Descending then -cmp else cmp
          }
        }
        .find(_ != 0)
        .getOrElse(0)

  /** ADT for sortable values with natural ordering */
  private enum SortValue:
    case Empty
    case Error
    case Num(value: Double)
    case Str(value: String)

  /** Extract a sortable value from a cell value */
  private def getSortableValue(value: CellValue, mode: SortMode): SortValue =
    import com.tjclp.xl.cells.CellValue.*
    value match
      case CellValue.Empty => SortValue.Empty
      case Text(s) =>
        if mode == SortMode.Numeric then
          s.toDoubleOption.map(SortValue.Num(_)).getOrElse(SortValue.Str(s.toLowerCase))
        else SortValue.Str(s.toLowerCase)
      case Number(n) => SortValue.Num(n.toDouble)
      case Bool(b) => SortValue.Num(if b then 1.0 else 0.0)
      case DateTime(dt) =>
        // Convert to Excel serial number for comparison
        SortValue.Num(CellValue.dateTimeToExcelSerial(dt))
      case Formula(_, Some(cached)) =>
        getSortableValue(cached, mode)
      case Formula(_, None) => SortValue.Str("")
      case RichText(rt) => SortValue.Str(rt.toPlainText.toLowerCase)
      case CellValue.Error(_) => SortValue.Error

  /** Compare two sort values (Empty/Error sort last) */
  private def compareSortValues(
    a: Option[SortValue],
    b: Option[SortValue],
    mode: SortMode
  ): Int =
    (a, b) match
      case (None, None) => 0
      case (None, _) => 1 // Empty cells sort last
      case (_, None) => -1
      case (Some(va), Some(vb)) =>
        (va, vb) match
          case (SortValue.Empty, SortValue.Empty) => 0
          case (SortValue.Empty, _) => 1
          case (_, SortValue.Empty) => -1
          case (SortValue.Error, SortValue.Error) => 0
          case (SortValue.Error, _) => 1
          case (_, SortValue.Error) => -1
          case (SortValue.Num(na), SortValue.Num(nb)) => na.compare(nb)
          case (SortValue.Str(sa), SortValue.Str(sb)) => sa.compare(sb)
          case (SortValue.Num(n), SortValue.Str(s)) =>
            // In numeric mode, numbers come before strings
            if mode == SortMode.Numeric then -1 else n.toString.compare(s)
          case (SortValue.Str(s), SortValue.Num(n)) =>
            if mode == SortMode.Numeric then 1 else s.compare(n.toString)

  /** Reconstruct sheet with rows in sorted order */
  private def reconstructWithSortedRows(
    sheet: Sheet,
    range: CellRange,
    dataRowStart: Int,
    sortedRows: Vector[(Int, Map[Int, Cell])]
  ): Sheet =
    val rowEnd = Row.index0(range.rowEnd)
    val colStart = Column.index0(range.colStart)
    val colEnd = Column.index0(range.colEnd)

    // Remove all cells in the data portion of the range
    val sheetWithoutDataCells = (dataRowStart to rowEnd).foldLeft(sheet) { (s, rowIdx) =>
      (colStart to colEnd).foldLeft(s) { (s2, colIdx) =>
        val ref = com.tjclp.xl.addressing.ARef.from0(colIdx, rowIdx)
        s2.remove(ref)
      }
    }

    // Insert sorted cells at new positions
    sortedRows.zipWithIndex.foldLeft(sheetWithoutDataCells) {
      case (s, ((originalRowIdx, cellMap), newRowOffset)) =>
        val newRowIdx = dataRowStart + newRowOffset
        val rowDelta = newRowIdx - originalRowIdx
        cellMap.values.foldLeft(s) { (s2, cell) =>
          // Create new cell with updated row position, preserving styleId and other properties
          val newRef = com.tjclp.xl.addressing.ARef.from0(cell.col.index0, newRowIdx)
          // Shift formula references if the row moved
          val shiftedValue = shiftCellValueForSort(cell.value, rowDelta)
          val newCell = Cell(newRef, shiftedValue, cell.styleId, cell.comment, cell.hyperlink)
          s2.put(newCell)
        }
    }

  /**
   * Shift formula references when a cell moves to a different row during sorting.
   *
   * When a row moves from position X to Y, relative row references in formulas should be adjusted
   * by (Y - X) so they continue pointing to the same relative positions. This matches Excel's
   * behavior for sorting.
   *
   * @param value
   *   The cell value to potentially shift
   * @param rowDelta
   *   The number of rows the cell moved (positive = down, negative = up)
   * @return
   *   Shifted cell value (or original if not a formula or rowDelta is 0)
   */
  private def shiftCellValueForSort(value: CellValue, rowDelta: Int): CellValue =
    if rowDelta == 0 then value
    else
      value match
        case CellValue.Formula(formula, _) =>
          val fullFormula = s"=$formula"
          FormulaParser.parse(fullFormula) match
            case Left(_) =>
              // If formula can't be parsed, keep as-is
              value
            case Right(parsedExpr) =>
              // Shift only row references, not columns (colDelta = 0)
              val shiftedExpr = FormulaShifter.shift(parsedExpr, 0, rowDelta)
              val shiftedFormula = FormulaPrinter.print(shiftedExpr, includeEquals = false)
              // Clear cached value since it will need re-evaluation
              CellValue.Formula(shiftedFormula, None)
        case other => other
