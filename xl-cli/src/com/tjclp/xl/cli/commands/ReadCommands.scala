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
import com.tjclp.xl.cli.raster.{RasterFormat, RasterizerChain}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.formula.{DependencyGraph, FormulaParser, SheetEvaluator}
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
    strict: Boolean,
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
    rasterizer: Option[String] = None
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
          // Pre-evaluate formulas if --eval flag is set (for cross-sheet reference support)
          val sheetToRender =
            if evalFormulas then
              evaluateSheetFormulas(targetSheet, Some(wb), Some(limitedRange), strict)
            else targetSheet
          IO.pure(
            Markdown.renderRange(
              sheetToRender,
              limitedRange,
              showFormulas,
              skipEmpty,
              evalFormulas = false
            )
          )
        case ViewFormat.Html =>
          // Pre-evaluate formulas if --eval flag is set
          val sheetToRender =
            if evalFormulas then
              evaluateSheetFormulas(targetSheet, Some(wb), Some(limitedRange), strict)
            else targetSheet
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
            if evalFormulas then
              evaluateSheetFormulas(targetSheet, Some(wb), Some(limitedRange), strict)
            else targetSheet
          IO.pure(
            sheetToRender.toSvg(
              limitedRange,
              theme = theme,
              showGridlines = showGridlines,
              showLabels = showLabels
            )
          )
        case ViewFormat.Json =>
          // Pre-evaluate formulas if --eval flag is set (for cross-sheet reference support)
          val sheetToRender =
            if evalFormulas then
              evaluateSheetFormulas(targetSheet, Some(wb), Some(limitedRange), strict)
            else targetSheet
          IO.pure(
            JsonRenderer.renderRange(
              sheetToRender,
              limitedRange,
              showFormulas,
              skipEmpty,
              headerRow,
              evalFormulas = false
            )
          )
        case ViewFormat.Csv =>
          // Pre-evaluate formulas if --eval flag is set (for cross-sheet reference support)
          val sheetToRender =
            if evalFormulas then
              evaluateSheetFormulas(targetSheet, Some(wb), Some(limitedRange), strict)
            else targetSheet
          IO.pure(
            CsvRenderer.renderRange(
              sheetToRender,
              limitedRange,
              showFormulas,
              showLabels,
              skipEmpty,
              evalFormulas = false
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
                if evalFormulas then
                  evaluateSheetFormulas(targetSheet, Some(wb), Some(limitedRange))
                else targetSheet
              val svg = sheetToRender.toSvg(
                limitedRange,
                theme = theme,
                showGridlines = showGridlines,
                showLabels = showLabels
              )

              // Convert ViewFormat to RasterFormat
              val rasterFormat = format match
                case ViewFormat.Png => RasterFormat.Png
                case ViewFormat.Jpeg => RasterFormat.Jpeg(quality)
                case ViewFormat.WebP => RasterFormat.WebP
                case ViewFormat.Pdf => RasterFormat.Pdf
                case _ => RasterFormat.Png // unreachable

              // Use RasterizerChain for automatic fallback
              RasterizerChain
                .convert(svg, outputPath, rasterFormat, dpi, rasterizer)
                .map { usedRasterizer =>
                  s"Exported: $outputPath (${format.toString.toLowerCase}, ${dpi} DPI, $usedRasterizer)"
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
      // Build dependency graph for dependencies/dependents (workbook-level for cross-sheet support)
      currentRef = DependencyGraph.QualifiedRef(targetSheet.name, ref)
      wbGraph = DependencyGraph.fromWorkbook(wb)
      // Get dependencies for this cell
      rawDeps = wbGraph.getOrElse(currentRef, Set.empty)
      // Build reverse graph for dependents
      allDependents = wbGraph.foldLeft(
        Map.empty[DependencyGraph.QualifiedRef, Set[DependencyGraph.QualifiedRef]]
      ) { case (acc, (source, targets)) =>
        targets.foldLeft(acc) { (m, target) =>
          m.updated(target, m.getOrElse(target, Set.empty) + source)
        }
      }
      rawDependents = allDependents.getOrElse(currentRef, Set.empty)
      // Format qualified refs - omit sheet name if same sheet as current cell
      formatQRef = (qref: DependencyGraph.QualifiedRef) =>
        if qref.sheet == targetSheet.name then qref.ref.toA1
        else s"${qref.sheet.value}!${qref.ref.toA1}"
      deps = rawDeps.toVector.sortBy(formatQRef).map(formatQRef)
      dependents = rawDependents.toVector.sortBy(formatQRef).map(formatQRef)
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
   *
   * For constant formulas (e.g., =1+1, =PI()*2) that don't reference any cells, both --file and
   * --sheet are optional. For formulas that reference cells, a file and sheet must be provided.
   */
  def eval(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    formulaStr: String,
    overrides: List[String]
  ): IO[String] =
    val formula = if formulaStr.startsWith("=") then formulaStr else s"=$formulaStr"

    // Check if formula needs a sheet by parsing and checking for cell references
    // GH-197: Use containsCellReferences instead of extractDependencies to avoid
    // enumerating 1M+ cells for full-column ranges like A:A
    val needsSheet = FormulaParser.parse(formula) match
      case Right(expr) =>
        // If formula has cell references or overrides are specified, require a sheet
        DependencyGraph.containsCellReferences(expr) || overrides.nonEmpty
      case Left(_) =>
        // Parse error - let SheetEvaluator handle it (it will give a better error message)
        false

    if needsSheet then
      // Formula references cells - require sheet
      // Check if workbook is empty (no file provided)
      if wb.sheets.isEmpty then
        IO.raiseError(
          new Exception(
            "Formula references cells but no file provided. Use --file to specify an Excel file."
          )
        )
      else
        for
          sheet <- SheetResolver.requireSheet(wb, sheetOpt, "eval")
          tempSheet <- applyOverrides(sheet, overrides)
          // Optimization: Only pre-evaluate formulas in the dependency closure of the target.
          // This ensures:
          // 1. Formula chains work correctly (e.g., =C1 where C1=B1+50, B1=A1*2)
          // 2. Overrides propagate through chains (TJC-698): A1=200 → B1=400 → C1=450
          // 3. Performance: Only evaluate O(k) formulas in closure, not O(n) total formulas

          // 1. Extract target formula's direct dependencies
          // GH-197: Use bounded extraction to avoid enumerating 1M+ cells for full-column ranges
          targetDeps <- IO.fromEither(
            FormulaParser
              .parse(formula)
              .map(expr => DependencyGraph.extractDependenciesBounded(expr, tempSheet.usedRange))
              .left
              .map(e => new Exception(s"Parse error: $e"))
          )

          // 2. Build dependency graph and compute transitive closure
          graph = DependencyGraph.fromSheet(tempSheet)
          allDeps = DependencyGraph.transitiveDependencies(graph, targetDeps)

          // 3. Filter to only formula cells in the closure
          formulaDeps = allDeps.filter(ref =>
            tempSheet(ref).value match
              case _: CellValue.Formula => true
              case _ => false
          )

          // 4. Get evaluation order (topological sort filtered to closure)
          evalOrder <- IO.fromEither(
            if formulaDeps.isEmpty then scala.util.Right(List.empty[ARef])
            else
              DependencyGraph
                .topologicalSort(graph)
                .map(_.filter(formulaDeps.contains))
                .left
                .map(e => new Exception(e.toString))
          )

          // 5. Evaluate only formulas in the closure
          evalSheet <- evalOrder.foldLeft(IO.pure(tempSheet)) { (sheetIO, ref) =>
            sheetIO.flatMap { s =>
              IO.fromEither(
                SheetEvaluator
                  .evaluateCell(s)(ref, workbook = Some(wb))
                  .map(value => s.put(ref, value))
                  .left
                  .map(e => new Exception(e.message))
              )
            }
          }

          result <- IO.fromEither(
            SheetEvaluator
              .evaluateFormula(evalSheet)(formula, workbook = Some(wb))
              .left
              .map(e => new Exception(e.message))
          )
        yield Format.evalSuccess(formula, result, overrides)
    else
      // Constant formula - use empty sheet or provided sheet
      val sheet = sheetOpt.getOrElse(Sheet("_eval"))
      // For constant formulas, workbook context is not needed
      val wbOpt = if wb.sheets.nonEmpty then Some(wb) else None
      for result <- IO.fromEither(
          SheetEvaluator
            .evaluateFormula(sheet)(formula, workbook = wbOpt)
            .left
            .map(e => new Exception(e.message))
        )
      yield Format.evalSuccess(formula, result, overrides)

  /**
   * Evaluate array formula and display result as table.
   *
   * Array formulas like TRANSPOSE return multiple values that are displayed as a grid. Unlike
   * regular eval which returns a single value, evala shows the full spilled array.
   */
  def evalArray(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    formulaStr: String,
    targetRefOpt: Option[String],
    overrides: List[String]
  ): IO[String] =
    val formula = if formulaStr.startsWith("=") then formulaStr else s"=$formulaStr"

    // Array formulas always need a sheet context
    if wb.sheets.isEmpty then
      IO.raiseError(
        new Exception(
          "Array formula evaluation requires a file. Use --file to specify an Excel file."
        )
      )
    else
      for
        sheet <- SheetResolver.requireSheet(wb, sheetOpt, "evala")
        tempSheet <- applyOverrides(sheet, overrides)
        // Pre-evaluate formulas in the dependency closure (same as eval)
        // GH-197: Use bounded extraction to avoid enumerating 1M+ cells for full-column ranges
        targetDeps <- IO.fromEither(
          FormulaParser
            .parse(formula)
            .map(expr => DependencyGraph.extractDependenciesBounded(expr, tempSheet.usedRange))
            .left
            .map(e => new Exception(s"Parse error: $e"))
        )
        graph = DependencyGraph.fromSheet(tempSheet)
        allDeps = DependencyGraph.transitiveDependencies(graph, targetDeps)
        formulaDeps = allDeps.filter(ref =>
          tempSheet(ref).value match
            case _: CellValue.Formula => true
            case _ => false
        )
        evalOrder <- IO.fromEither(
          if formulaDeps.isEmpty then scala.util.Right(List.empty[ARef])
          else
            DependencyGraph
              .topologicalSort(graph)
              .map(_.filter(formulaDeps.contains))
              .left
              .map(e => new Exception(e.toString))
        )
        evalSheet <- evalOrder.foldLeft(IO.pure(tempSheet)) { (sheetIO, ref) =>
          sheetIO.flatMap { s =>
            IO.fromEither(
              SheetEvaluator
                .evaluateCell(s)(ref, workbook = Some(wb))
                .map(value => s.put(ref, value))
                .left
                .map(e => new Exception(e.message))
            )
          }
        }
        // Parse target ref or use a virtual cell far from data
        originRef = targetRefOpt
          .flatMap(ARef.parse(_).toOption)
          .getOrElse(ARef.from0(25, 999)) // Z1000
        result <- IO.fromEither(
          SheetEvaluator
            .evaluateArrayFormula(evalSheet)(formula, originRef, workbook = Some(wb))
            .left
            .map(e => new Exception(e.message))
        )
        (updatedSheet, spillRange) = result
      yield Format.evalArraySuccess(formula, updatedSheet, spillRange, overrides)

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
   * Evaluate formula cells in a sheet, replacing them with their computed values.
   *
   * This is used to pre-process sheets for rendering when --eval is specified. Formula cells are
   * replaced with their evaluated results (preserving formatting), while non-formula cells remain
   * unchanged.
   *
   * Uses dependency-aware evaluation to correctly handle nested formulas (e.g., F5=SUM(B5:E5) where
   * B5=SUM(B2:B4)). Formulas are evaluated in topological order so dependencies are computed first.
   *
   * When a range is specified, only evaluates formulas within that range (plus their transitive
   * dependencies). This is much more efficient than evaluating the entire sheet when viewing a
   * small subset.
   *
   * @param sheet
   *   The sheet to evaluate formulas in
   * @param workbook
   *   Optional workbook context for cross-sheet formula references
   * @param range
   *   Optional range to limit evaluation to (formulas outside this range are not evaluated)
   */
  private def evaluateSheetFormulas(
    sheet: Sheet,
    workbook: Option[Workbook] = None,
    range: Option[CellRange] = None,
    strict: Boolean = false
  ): Sheet =
    val evalResult = range match
      case Some(r) =>
        // Targeted evaluation: only evaluate formulas in range + their dependencies
        SheetEvaluator.evaluateForRange(sheet)(r, workbook = workbook)
      case None =>
        // Full evaluation: evaluate all formulas
        SheetEvaluator.evaluateWithDependencyCheck(sheet)(workbook = workbook)

    evalResult match
      case Right(results) =>
        // Apply evaluated results. Sheet.put preserves existing cell styleId automatically.
        results.foldLeft(sheet) { case (acc, (ref, value)) =>
          acc.put(ref, value)
        }
      case Left(error) =>
        if strict then throw new Exception(s"Formula evaluation failed: ${error.message}")
        else
          // Warn on stderr but return original sheet
          Console.err.println(s"Warning: Formula evaluation failed: ${error.message}")
          sheet
