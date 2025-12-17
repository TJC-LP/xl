package com.tjclp.xl.cli.commands

import java.nio.file.Path

import cats.effect.IO
import cats.implicits.*
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, RefType, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.ViewFormat
import com.tjclp.xl.cli.helpers.{SheetResolver, ValueParser}
import com.tjclp.xl.cli.output.{CsvRenderer, Format, JsonRenderer, Markdown}
import com.tjclp.xl.cli.raster.{BatikRasterizer, ImageMagick}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.formula.{DependencyGraph, SheetEvaluator}
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Read command handlers.
 *
 * Commands that read data without modification: bounds, view, cell, search, stats, eval.
 */
object ReadCommands:

  /**
   * Show used range of current sheet.
   */
  def bounds(wb: Workbook, sheetOpt: Option[Sheet]): IO[String] =
    SheetResolver.requireSheet(wb, sheetOpt, "bounds").map { sheet =>
      val name = sheet.name.value
      val usedRange = sheet.usedRange
      val cellCount = sheet.cells.size
      usedRange match
        case Some(range) =>
          val rowCount = range.end.row.index0 - range.start.row.index0 + 1
          val colCount = range.end.col.index0 - range.start.col.index0 + 1
          s"""Sheet: $name
             |Used range: ${range.toA1}
             |Rows: ${range.start.row.index1}-${range.end.row.index1} ($rowCount total)
             |Columns: ${range.start.col.toLetter}-${range.end.col.toLetter} ($colCount total)
             |Non-empty: $cellCount cells""".stripMargin
        case None =>
          s"""Sheet: $name
             |Used range: (empty)
             |Non-empty: 0 cells""".stripMargin
    }

  /**
   * View range in various formats.
   */
  def view(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    rangeStr: String,
    showFormulas: Boolean,
    evalFormulas: Boolean,
    limit: Int,
    format: ViewFormat,
    printScale: Boolean,
    showGridlines: Boolean,
    showLabels: Boolean,
    dpi: Int,
    quality: Int,
    rasterOutput: Option[Path],
    skipEmpty: Boolean,
    headerRow: Option[Int],
    useImageMagick: Boolean = false
  ): IO[String] =
    for
      resolved <- SheetResolver.resolveRef(wb, sheetOpt, rangeStr, "view")
      (targetSheet, refOrRange) = resolved
      range = refOrRange match
        case Right(r) => r
        case Left(ref) => CellRange(ref, ref) // Single cell as range
      limitedRange = limitRange(range, limit)
      theme = wb.metadata.theme // Use workbook's parsed theme
      result <- format match
        case ViewFormat.Markdown =>
          IO.pure(
            Markdown.renderRange(targetSheet, limitedRange, showFormulas, skipEmpty, evalFormulas)
          )
        case ViewFormat.Html =>
          // Pre-evaluate formulas if --eval flag is set
          val sheetToRender =
            if evalFormulas then evaluateSheetFormulas(targetSheet) else targetSheet
          IO.pure(
            sheetToRender.toHtml(
              limitedRange,
              theme = theme,
              applyPrintScale = printScale,
              showLabels = showLabels
            )
          )
        case ViewFormat.Svg =>
          // Pre-evaluate formulas if --eval flag is set
          val sheetToRender =
            if evalFormulas then evaluateSheetFormulas(targetSheet) else targetSheet
          IO.pure(
            sheetToRender.toSvg(
              limitedRange,
              theme = theme,
              showGridlines = showGridlines,
              showLabels = showLabels
            )
          )
        case ViewFormat.Json =>
          IO.pure(
            JsonRenderer.renderRange(
              targetSheet,
              limitedRange,
              showFormulas,
              skipEmpty,
              headerRow,
              evalFormulas
            )
          )
        case ViewFormat.Csv =>
          IO.pure(
            CsvRenderer.renderRange(
              targetSheet,
              limitedRange,
              showFormulas,
              showLabels,
              skipEmpty,
              evalFormulas
            )
          )
        case ViewFormat.Png | ViewFormat.Jpeg | ViewFormat.WebP | ViewFormat.Pdf =>
          rasterOutput match
            case None =>
              IO.raiseError(
                new Exception(
                  s"--raster-output required for ${format.toString.toLowerCase} format (binary output cannot go to stdout)"
                )
              )
            case Some(outputPath) =>
              // Pre-evaluate formulas if --eval flag is set
              val sheetToRender =
                if evalFormulas then evaluateSheetFormulas(targetSheet) else targetSheet
              val svg = sheetToRender.toSvg(
                limitedRange,
                theme = theme,
                showGridlines = showGridlines,
                showLabels = showLabels
              )

              // Use Batik (pure JVM) by default, ImageMagick as fallback
              // Batik requires AWT which may not be available in GraalVM native images
              if useImageMagick then
                val rasterFormat = format match
                  case ViewFormat.Png => ImageMagick.Format.Png
                  case ViewFormat.Jpeg => ImageMagick.Format.Jpeg(quality)
                  case ViewFormat.WebP => ImageMagick.Format.WebP
                  case ViewFormat.Pdf => ImageMagick.Format.Pdf
                  case _ => ImageMagick.Format.Png // unreachable
                ImageMagick.convertSvgToRaster(svg, outputPath, rasterFormat, dpi).map { _ =>
                  s"Exported: $outputPath (${format.toString.toLowerCase}, ${dpi} DPI, ImageMagick)"
                }
              else
                val batikFormat = format match
                  case ViewFormat.Png => BatikRasterizer.Format.Png
                  case ViewFormat.Jpeg => BatikRasterizer.Format.Jpeg(quality)
                  case ViewFormat.WebP => BatikRasterizer.Format.WebP
                  case ViewFormat.Pdf => BatikRasterizer.Format.Pdf
                  case _ => BatikRasterizer.Format.Png // unreachable
                // Try Batik first, fall back to ImageMagick if AWT not available
                BatikRasterizer
                  .convertSvgToRaster(svg, outputPath, batikFormat, dpi)
                  .map(_ => s"Exported: $outputPath (${format.toString.toLowerCase}, ${dpi} DPI)")
                  .handleErrorWith {
                    case e: BatikRasterizer.RasterizationError
                        if e.getMessage.contains("AWT") || e.getMessage.contains("native image") =>
                      // Batik failed due to AWT, try ImageMagick as fallback
                      val imgFormat = format match
                        case ViewFormat.Png => ImageMagick.Format.Png
                        case ViewFormat.Jpeg => ImageMagick.Format.Jpeg(quality)
                        case ViewFormat.WebP => ImageMagick.Format.WebP
                        case ViewFormat.Pdf => ImageMagick.Format.Pdf
                        case _ => ImageMagick.Format.Png
                      ImageMagick
                        .convertSvgToRaster(svg, outputPath, imgFormat, dpi)
                        .map { _ =>
                          s"Warning: Batik unavailable (native image), using ImageMagick fallback\n" +
                            s"Exported: $outputPath (${format.toString.toLowerCase}, ${dpi} DPI, ImageMagick)"
                        }
                        .handleErrorWith { imgErr =>
                          IO.raiseError(
                            new Exception(
                              s"Rasterization failed: Batik requires AWT (unavailable in native image), " +
                                s"and ImageMagick fallback failed: ${imgErr.getMessage}"
                            )
                          )
                        }
                    case e => IO.raiseError(e)
                  }
    yield result

  /**
   * Get cell details.
   */
  def cell(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    refStr: String,
    noStyle: Boolean
  ): IO[String] =
    for
      resolved <- SheetResolver.resolveRef(wb, sheetOpt, refStr, "cell")
      (targetSheet, refOrRange) = resolved
      ref <- refOrRange match
        case Left(r) => IO.pure(r)
        case Right(_) =>
          IO.raiseError(new Exception("cell command requires single cell, not range"))
      cellOpt = targetSheet.cells.get(ref)
      value = cellOpt.map(_.value).getOrElse(CellValue.Empty)
      // Get style from registry for NumFmt formatting (unless --no-style)
      style =
        if noStyle then None else cellOpt.flatMap(_.styleId).flatMap(targetSheet.styleRegistry.get)
      numFmt = style.map(_.numFmt).getOrElse(NumFmt.General)
      // For formulas with cached values, format the cached value
      valueToFormat = value match
        case CellValue.Formula(_, Some(cached)) => cached
        case other => other
      formatted = NumFmtFormatter.formatValue(valueToFormat, numFmt)
      // Get comment from sheet (sheet.getComment, not cell.comment)
      comment = targetSheet.getComment(ref)
      // Get hyperlink from cell
      hyperlink = cellOpt.flatMap(_.hyperlink)
      // Build dependency graph for dependencies/dependents
      graph = DependencyGraph.fromSheet(targetSheet)
      deps = graph.dependencies.getOrElse(ref, Set.empty).toVector.sortBy(_.toA1)
      dependents = graph.dependents.getOrElse(ref, Set.empty).toVector.sortBy(_.toA1)
    yield Format.cellInfo(ref, value, formatted, style, comment, hyperlink, deps, dependents)

  /**
   * Search for cells matching pattern.
   */
  def search(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    pattern: String,
    limit: Int,
    sheetsFilter: Option[String]
  ): IO[String] =
    IO.fromEither(
      scala.util
        .Try(pattern.r)
        .toEither
        .left
        .map(e => new Exception(s"Invalid regex pattern: ${e.getMessage}"))
    ).flatMap { regex =>
      // Determine which sheets to search:
      // 1. If --sheets provided, use those
      // 2. Else if --sheet provided, use that
      // 3. Else search all sheets
      val targetSheets: IO[Vector[Sheet]] = (sheetsFilter, sheetOpt) match
        case (Some(filterStr), _) =>
          // --sheets=Sheet1,Sheet2
          val names = filterStr.split(",").map(_.trim).toVector
          names.traverse { name =>
            IO.fromEither(SheetName(name).left.map(e => new Exception(e))).flatMap { sn =>
              IO.fromOption(wb.sheets.find(_.name == sn))(
                new Exception(
                  s"Sheet not found: $name. Available: ${wb.sheets.map(_.name.value).mkString(", ")}"
                )
              )
            }
          }
        case (None, Some(sheet)) =>
          // --sheet provided, use that single sheet
          IO.pure(Vector(sheet))
        case (None, None) =>
          // No filter, search all sheets
          IO.pure(wb.sheets)

      targetSheets.map { sheets =>
        val results = sheets.iterator
          .flatMap { s =>
            s.cells.iterator
              .filter { case (_, cell) =>
                val text = ValueParser.formatCellValue(cell.value)
                regex.findFirstIn(text).isDefined
              }
              .map { case (ref, cell) =>
                val value = ValueParser.formatCellValue(cell.value)
                // Always use qualified refs for clarity
                (s"${s.name.value}!${ref.toA1}", value)
              }
          }
          .take(limit)
          .toVector
        val sheetDesc =
          if sheets.size == 1 then sheets.headOption.map(_.name.value).getOrElse("sheet")
          else s"${sheets.size} sheets"
        s"Found ${results.size} matches in $sheetDesc:\n\n${Markdown.renderSearchResultsWithRef(results)}"
      }
    }

  /**
   * Calculate statistics for numeric values in range.
   */
  def stats(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    refStr: String
  ): IO[String] =
    for
      resolved <- SheetResolver.resolveRef(wb, sheetOpt, refStr, "stats")
      (sheet, refOrRange) = resolved
      range = refOrRange match
        case Right(r) => r
        case Left(ref) => CellRange(ref, ref) // Single cell as 1x1 range
      cells = sheet.getRange(range)
      numbers = cells.flatMap { cell =>
        cell.value match
          case CellValue.Number(n) => Some(n)
          case CellValue.Formula(_, Some(CellValue.Number(n))) => Some(n)
          case _ => None
      }.toVector
      _ <- IO
        .raiseError(new Exception(s"No numeric values in range ${range.toA1}"))
        .whenA(numbers.isEmpty)
    yield
      val count = numbers.size
      val sum = numbers.sum
      val min = numbers.foldLeft(BigDecimal(Double.MaxValue))(_ min _)
      val max = numbers.foldLeft(BigDecimal(Double.MinValue))(_ max _)
      val mean = sum / count
      f"count: $count, sum: $sum%.2f, min: $min%.2f, max: $max%.2f, mean: $mean%.2f"

  /**
   * Evaluate formula without modifying sheet.
   */
  def eval(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    formulaStr: String,
    overrides: List[String]
  ): IO[String] =
    for
      sheet <- SheetResolver.requireSheet(wb, sheetOpt, "eval")
      tempSheet <- applyOverrides(sheet, overrides)
      formula = if formulaStr.startsWith("=") then formulaStr else s"=$formulaStr"
      result <- IO.fromEither(
        SheetEvaluator.evaluateFormula(tempSheet)(formula).left.map(e => new Exception(e.message))
      )
    yield Format.evalSuccess(formula, result, overrides)

  // ==========================================================================
  // Private helpers
  // ==========================================================================

  private def limitRange(range: CellRange, maxRows: Int): CellRange =
    val rowCount = range.end.row.index0 - range.start.row.index0 + 1
    if rowCount <= maxRows then range
    else
      val newEndRow = range.start.row.index0 + maxRows - 1
      CellRange(range.start, ARef.from0(range.end.col.index0, newEndRow))

  private def applyOverrides(sheet: Sheet, overrides: List[String]): IO[Sheet] =
    overrides.foldLeft(IO.pure(sheet)) { (sheetIO, override_) =>
      sheetIO.flatMap { s =>
        override_.split("=", 2) match
          case Array(refStr, valueStr) if valueStr.trim.nonEmpty =>
            IO.fromEither(RefType.parse(refStr.trim).left.map(e => new Exception(e))).flatMap {
              case RefType.Cell(ref) =>
                val value = ValueParser.parseValue(valueStr.trim)
                IO.pure(s.put(ref, value))
              case RefType.QualifiedCell(sheetName, ref) =>
                if sheetName == sheet.name then
                  val value = ValueParser.parseValue(valueStr.trim)
                  IO.pure(s.put(ref, value))
                else
                  IO.raiseError(
                    new Exception(
                      s"Cross-sheet override not supported: ${refStr.trim}. " +
                        s"Eval operates on ${sheet.name.value}, not ${sheetName.value}"
                    )
                  )
              case RefType.Range(_) | RefType.QualifiedRange(_, _) =>
                IO.raiseError(
                  new Exception(s"Override requires single cell, not range: ${refStr.trim}")
                )
            }
          case Array(refStr, _) =>
            IO.raiseError(
              new Exception(
                s"Empty value for override: ${refStr.trim}. Use ref=value (e.g., B5=1000)"
              )
            )
          case _ =>
            IO.raiseError(
              new Exception(s"Invalid override format: $override_. Use ref=value (e.g., B5=1000)")
            )
      }
    }

  /**
   * Evaluate all formula cells in a sheet, replacing them with their computed values.
   *
   * This is used to pre-process sheets for rendering when --eval is specified. Formula cells are
   * replaced with their evaluated results (preserving formatting), while non-formula cells remain
   * unchanged.
   */
  private def evaluateSheetFormulas(sheet: Sheet): Sheet =
    sheet.cells.foldLeft(sheet) { case (acc, (ref, cell)) =>
      cell.value match
        case CellValue.Formula(expr, _) =>
          val displayExpr = if expr.startsWith("=") then expr else s"=$expr"
          SheetEvaluator.evaluateFormula(sheet)(displayExpr) match
            case Right(result) =>
              // Replace formula with evaluated result, keeping style
              acc.put(ref, result)
            case Left(error) =>
              // Warn on stderr but keep formula and continue processing
              Console.err.println(
                s"Warning: Failed to evaluate formula at ${ref.toA1}: ${error.message}"
              )
              acc
        case _ =>
          acc
    }
