package com.tjclp.xl.cli

import java.nio.file.Path

import scala.util.chaining.*

import cats.effect.{ExitCode, IO, IOApp}
import cats.implicits.*
import com.monovore.decline.*
import com.monovore.decline.effect.*

import com.tjclp.xl.{*, given}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.addressing.{ARef, CellRange, Column, RefType, Row, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.output.{CsvRenderer, Format, JsonRenderer, Markdown}
import com.tjclp.xl.cli.raster.ImageMagick
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.formula.{DependencyGraph, SheetEvaluator}
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties, styleSyntax}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

/** Read version from generated resource, fallback to dev */
private object BuildInfo:
  val version: String =
    val props = new java.util.Properties()
    val stream = Option(getClass.getResourceAsStream("/version.properties"))
    stream.foreach(props.load)
    Option(props.getProperty("version")).getOrElse("dev")

/**
 * XL CLI - LLM-friendly Excel operations.
 *
 * Stateless by design: each command is self-contained. Use global flags:
 *   - `-f, --file` — Input file (required)
 *   - `-s, --sheet` — Sheet name (optional, defaults to first)
 *   - `-o, --output` — Output file for mutations (required for put/putf)
 */
object Main
    extends CommandIOApp(
      name = "xl",
      header = "LLM-friendly Excel operations (stateless)",
      version = BuildInfo.version
    ):

  override def main: Opts[IO[ExitCode]] =
    // Workbook-level: only --file (no --sheet)
    val workbookOpts = (fileOpt, sheetsCmd).mapN { (file, cmd) =>
      run(file, None, None, cmd)
    }

    // Sheet-level read-only: --file and --sheet (no --output)
    val sheetReadOnlySubcmds =
      boundsCmd orElse viewCmd orElse cellCmd orElse searchCmd orElse evalCmd

    val sheetReadOnlyOpts = (fileOpt, sheetOpt, sheetReadOnlySubcmds).mapN { (file, sheet, cmd) =>
      run(file, sheet, None, cmd)
    }

    // Sheet-level write: --file, --sheet, and --output (required)
    val sheetWriteSubcmds =
      putCmd orElse putfCmd orElse styleCmd orElse rowCmd orElse colCmd orElse batchCmd

    val sheetWriteOpts =
      (fileOpt, sheetOpt, outputOpt, sheetWriteSubcmds).mapN { (file, sheet, out, cmd) =>
        run(file, sheet, Some(out), cmd)
      }

    // Standalone: no --file required (creates new files)
    val standaloneOpts = newCmd.map { case (outPath, sheetName) =>
      runStandalone(outPath, sheetName)
    }

    standaloneOpts orElse workbookOpts orElse sheetReadOnlyOpts orElse sheetWriteOpts

  // ==========================================================================
  // Global options
  // ==========================================================================

  private val fileOpt =
    Opts.option[Path]("file", "Excel file to operate on (required)", "f")

  private val sheetOpt =
    Opts.option[String]("sheet", "Sheet to select (optional, defaults to first)", "s").orNone

  private val outputOpt =
    Opts.option[Path]("output", "Output file (required)", "o")

  // ==========================================================================
  // Command definitions
  // ==========================================================================

  private val rangeArg = Opts.argument[String]("range")
  private val refArg = Opts.argument[String]("ref")
  private val valueArg = Opts.argument[String]("value")
  private val patternArg = Opts.argument[String]("pattern")

  private val formulasOpt = Opts.flag("formulas", "Show formulas instead of values").orFalse
  private val limitOpt = Opts.option[Int]("limit", "Maximum rows to display").withDefault(50)
  private val formatOpt = Opts
    .option[String]("format", "Output format: markdown, html, svg, json, csv, png, jpeg, webp, pdf")
    .withDefault("markdown")
    .mapValidated { s =>
      s.toLowerCase match
        case "markdown" | "md" => cats.data.Validated.valid(ViewFormat.Markdown)
        case "html" => cats.data.Validated.valid(ViewFormat.Html)
        case "svg" => cats.data.Validated.valid(ViewFormat.Svg)
        case "json" => cats.data.Validated.valid(ViewFormat.Json)
        case "csv" => cats.data.Validated.valid(ViewFormat.Csv)
        case "png" => cats.data.Validated.valid(ViewFormat.Png)
        case "jpeg" | "jpg" => cats.data.Validated.valid(ViewFormat.Jpeg)
        case "webp" => cats.data.Validated.valid(ViewFormat.WebP)
        case "pdf" => cats.data.Validated.valid(ViewFormat.Pdf)
        case other =>
          cats.data.Validated.invalidNel(
            s"Unknown format: $other. Use markdown, html, svg, json, csv, png, jpeg, webp, or pdf"
          )
    }
  private val printScaleOpt =
    Opts.flag("print-scale", "Apply print scaling (for PDF-like output)").orFalse
  private val gridlinesOpt =
    Opts.flag("gridlines", "Show cell gridlines in SVG output").orFalse
  private val showLabelsOpt =
    Opts.flag("show-labels", "Include column letters (A, B, C) and row numbers (1, 2, 3)").orFalse
  private val dpiOpt =
    Opts.option[Int]("dpi", "DPI for raster output (default: 144 for retina)").withDefault(144)
  private val qualityOpt =
    Opts.option[Int]("quality", "JPEG quality 1-100 (default: 90)").withDefault(90)
  private val rasterOutputOpt =
    Opts
      .option[Path](
        "raster-output",
        "Output file for raster formats (required for png/jpeg/webp/pdf)"
      )
      .orNone
  private val skipEmptyOpt =
    Opts.flag("skip-empty", "Skip empty cells (JSON) or empty rows/columns (tabular)").orFalse
  private val headerRowOpt =
    Opts
      .option[Int](
        "header-row",
        "Use values from this row as keys in JSON output (1-based row number)"
      )
      .orNone
  private val allSheetsOpt =
    Opts.flag("all-sheets", "Search across all sheets in the workbook").orFalse

  // --- Standalone commands (no --file required) ---

  private val outputArg = Opts.argument[Path]("output")
  private val sheetNameOpt =
    Opts.option[String]("sheet-name", "Sheet name (defaults to 'Sheet1')").withDefault("Sheet1")

  val newCmd: Opts[(Path, String)] = Opts.subcommand("new", "Create a blank xlsx file") {
    (outputArg, sheetNameOpt).tupled
  }

  // --- Read-only commands ---

  val sheetsCmd: Opts[Command] = Opts.subcommand("sheets", "List all sheets") {
    Opts(Command.Sheets)
  }

  val boundsCmd: Opts[Command] = Opts.subcommand("bounds", "Show used range of current sheet") {
    Opts(Command.Bounds)
  }

  val viewCmd: Opts[Command] =
    Opts.subcommand("view", "View range (markdown, html, svg, json, csv, png, jpeg, webp, pdf)") {
      (
        rangeArg,
        formulasOpt,
        limitOpt,
        formatOpt,
        printScaleOpt,
        gridlinesOpt,
        showLabelsOpt,
        dpiOpt,
        qualityOpt,
        rasterOutputOpt,
        skipEmptyOpt,
        headerRowOpt
      )
        .mapN(Command.View.apply)
    }

  val cellCmd: Opts[Command] = Opts.subcommand("cell", "Get cell details") {
    refArg.map(Command.Cell.apply)
  }

  val searchCmd: Opts[Command] = Opts.subcommand("search", "Search for cells") {
    (patternArg, limitOpt, allSheetsOpt).mapN(Command.Search.apply)
  }

  // --- Analyze ---

  private val formulaArg = Opts.argument[String]("formula")
  private val withOpt =
    Opts.option[String]("with", "Comma-separated overrides (e.g., A1=100,B2=200)", "w").orNone

  val evalCmd: Opts[Command] = Opts.subcommand("eval", "Evaluate formula without modifying sheet") {
    (formulaArg, withOpt).mapN { (formula, withStr) =>
      val overrides = withStr.toList.flatMap(_.split(",").map(_.trim).filter(_.nonEmpty))
      Command.Eval(formula, overrides)
    }
  }

  // --- Mutate (require -o) ---

  val putCmd: Opts[Command] = Opts.subcommand("put", "Write value to cell") {
    (refArg, valueArg).mapN(Command.Put.apply)
  }

  val putfCmd: Opts[Command] = Opts.subcommand("putf", "Write formula to cell") {
    (refArg, valueArg).mapN(Command.PutFormula.apply)
  }

  // --- Style command options ---
  private val boldOpt = Opts.flag("bold", "Bold text").orFalse
  private val italicOpt = Opts.flag("italic", "Italic text").orFalse
  private val underlineOpt = Opts.flag("underline", "Underline text").orFalse
  private val bgOpt =
    Opts.option[String]("bg", "Background color (name, #hex, or rgb(r,g,b))").orNone
  private val fgOpt = Opts.option[String]("fg", "Text color (name, #hex, or rgb(r,g,b))").orNone
  private val fontSizeOpt = Opts.option[Double]("font-size", "Font size in points").orNone
  private val fontNameOpt = Opts.option[String]("font-name", "Font family name").orNone
  private val alignOpt = Opts.option[String]("align", "Horizontal: left, center, right").orNone
  private val valignOpt = Opts.option[String]("valign", "Vertical: top, middle, bottom").orNone
  private val wrapOpt = Opts.flag("wrap", "Enable text wrapping").orFalse
  private val numFormatOpt =
    Opts
      .option[String]("format", "Number format: general, number, currency, percent, date, text")
      .orNone
  private val borderOpt =
    Opts.option[String]("border", "Border style: none, thin, medium, thick").orNone
  private val borderColorOpt = Opts.option[String]("border-color", "Border color").orNone

  val styleCmd: Opts[Command] = Opts.subcommand("style", "Apply styling to cells") {
    (
      rangeArg,
      boldOpt,
      italicOpt,
      underlineOpt,
      bgOpt,
      fgOpt,
      fontSizeOpt,
      fontNameOpt,
      alignOpt,
      valignOpt,
      wrapOpt,
      numFormatOpt,
      borderOpt,
      borderColorOpt
    ).mapN(Command.Style.apply)
  }

  // --- Row/Column command options ---
  private val rowArg = Opts.argument[Int]("row")
  private val colArg = Opts.argument[String]("col")
  private val heightOpt = Opts.option[Double]("height", "Row height in points").orNone
  private val widthOpt = Opts.option[Double]("width", "Column width in character units").orNone
  private val hideOpt = Opts.flag("hide", "Hide row/column").orFalse
  private val showOpt = Opts.flag("show", "Show (unhide) row/column").orFalse

  val rowCmd: Opts[Command] = Opts.subcommand("row", "Set row properties (height, hide/show)") {
    (rowArg, heightOpt, hideOpt, showOpt).mapN(Command.RowOp.apply)
  }

  val colCmd: Opts[Command] = Opts.subcommand("col", "Set column properties (width, hide/show)") {
    (colArg, widthOpt, hideOpt, showOpt).mapN(Command.ColOp.apply)
  }

  // --- Batch command ---
  private val batchArg = Opts.argument[String]("operations").withDefault("-")
  val batchCmd: Opts[Command] =
    Opts.subcommand("batch", "Apply multiple operations atomically (JSON from stdin or file)") {
      batchArg.map(Command.Batch.apply)
    }

  // ==========================================================================
  // Command execution
  // ==========================================================================

  private def run(
    filePath: Path,
    sheetNameOpt: Option[String],
    outputOpt: Option[Path],
    cmd: Command
  ): IO[ExitCode] =
    execute(filePath, sheetNameOpt, outputOpt, cmd).attempt.flatMap {
      case Right(output) =>
        IO.println(output).as(ExitCode.Success)
      case Left(err) =>
        IO.println(Format.errorSimple(err.getMessage)).as(ExitCode.Error)
    }

  private def runStandalone(outPath: Path, sheetName: String): IO[ExitCode] =
    (for
      name <- IO.fromEither(SheetName(sheetName).left.map(e => new Exception(e)))
      wb = Workbook(Vector(Sheet(name)))
      _ <- ExcelIO.instance[IO].write(wb, outPath)
    yield s"Created ${outPath.toAbsolutePath} with sheet '$sheetName'").attempt.flatMap {
      case Right(output) =>
        IO.println(output).as(ExitCode.Success)
      case Left(err) =>
        IO.println(Format.errorSimple(err.getMessage)).as(ExitCode.Error)
    }

  private def execute(
    filePath: Path,
    sheetNameOpt: Option[String],
    outputOpt: Option[Path],
    cmd: Command
  ): IO[String] =
    for
      wb <- ExcelIO.instance[IO].read(filePath)
      sheet <- resolveSheet(wb, sheetNameOpt)
      result <- executeCommand(wb, sheet, outputOpt, cmd)
    yield result

  private def resolveSheet(wb: Workbook, sheetNameOpt: Option[String]): IO[Sheet] =
    sheetNameOpt match
      case Some(name) =>
        IO.fromEither(SheetName.apply(name).left.map(e => new Exception(e))).flatMap { sheetName =>
          IO.fromOption(wb.sheets.find(_.name == sheetName))(
            new Exception(
              s"Sheet not found: $name. Available: ${wb.sheets.map(_.name.value).mkString(", ")}"
            )
          )
        }
      case None =>
        IO.fromOption(wb.sheets.headOption)(new Exception("Workbook has no sheets"))

  /**
   * Resolve a reference string to a (Sheet, Either[ARef, CellRange]).
   *
   * Supports qualified refs like `Sheet1!A1` which override the default sheet context. This mirrors
   * Excel's behavior: you can be "on" Sheet1 while referencing Sheet2!A1.
   */
  private def resolveRef(
    wb: Workbook,
    defaultSheet: Sheet,
    refStr: String
  ): IO[(Sheet, Either[ARef, CellRange])] =
    IO.fromEither(RefType.parse(refStr).left.map(e => new Exception(e))).flatMap {
      case RefType.Cell(ref) =>
        IO.pure((defaultSheet, Left(ref)))
      case RefType.Range(range) =>
        IO.pure((defaultSheet, Right(range)))
      case RefType.QualifiedCell(sheetName, ref) =>
        findSheet(wb, sheetName).map(s => (s, Left(ref)))
      case RefType.QualifiedRange(sheetName, range) =>
        findSheet(wb, sheetName).map(s => (s, Right(range)))
    }

  private def findSheet(wb: Workbook, name: SheetName): IO[Sheet] =
    IO.fromOption(wb.sheets.find(_.name == name))(
      new Exception(
        s"Sheet not found: ${name.value}. Available: ${wb.sheets.map(_.name.value).mkString(", ")}"
      )
    )

  private def executeCommand(
    wb: Workbook,
    sheet: Sheet,
    outputOpt: Option[Path],
    cmd: Command
  ): IO[String] = cmd match

    case Command.Sheets =>
      val sheetStats = wb.sheets.map { s =>
        val usedRange = s.usedRange
        val cellCount = s.cells.size
        val formulaCount = s.cells.values.count(_.isFormula)
        (s.name.value, usedRange, cellCount, formulaCount)
      }
      IO.pure(Markdown.renderSheetList(sheetStats))

    case Command.Bounds =>
      val name = sheet.name.value
      val usedRange = sheet.usedRange
      val cellCount = sheet.cells.size
      IO.pure(usedRange match
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
             |Non-empty: 0 cells""".stripMargin)

    case Command.View(
          rangeStr,
          showFormulas,
          limit,
          format,
          printScale,
          showGridlines,
          showLabels,
          dpi,
          quality,
          rasterOutput,
          skipEmpty,
          headerRow
        ) =>
      for
        resolved <- resolveRef(wb, sheet, rangeStr)
        (targetSheet, refOrRange) = resolved
        range = refOrRange match
          case Right(r) => r
          case Left(ref) => CellRange(ref, ref) // Single cell as range
        limitedRange = limitRange(range, limit)
        theme = wb.metadata.theme // Use workbook's parsed theme
        result <- format match
          case ViewFormat.Markdown =>
            IO.pure(Markdown.renderRange(targetSheet, limitedRange, showFormulas, skipEmpty))
          case ViewFormat.Html =>
            IO.pure(
              targetSheet.toHtml(
                limitedRange,
                theme = theme,
                applyPrintScale = printScale,
                showLabels = showLabels
              )
            )
          case ViewFormat.Svg =>
            IO.pure(
              targetSheet.toSvg(
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
                headerRow
              )
            )
          case ViewFormat.Csv =>
            IO.pure(
              CsvRenderer.renderRange(
                targetSheet,
                limitedRange,
                showFormulas,
                showLabels,
                skipEmpty
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
                val svg = targetSheet.toSvg(
                  limitedRange,
                  theme = theme,
                  showGridlines = showGridlines,
                  showLabels = showLabels
                )
                val rasterFormat = format match
                  case ViewFormat.Png => ImageMagick.Format.Png
                  case ViewFormat.Jpeg => ImageMagick.Format.Jpeg(quality)
                  case ViewFormat.WebP => ImageMagick.Format.WebP
                  case ViewFormat.Pdf => ImageMagick.Format.Pdf
                  case _ => ImageMagick.Format.Png // unreachable
                ImageMagick.convertSvgToRaster(svg, outputPath, rasterFormat, dpi).map { _ =>
                  s"Exported: $outputPath (${format.toString.toLowerCase}, ${dpi} DPI)"
                }
      yield result

    case Command.Cell(refStr) =>
      for
        resolved <- resolveRef(wb, sheet, refStr)
        (targetSheet, refOrRange) = resolved
        ref <- refOrRange match
          case Left(r) => IO.pure(r)
          case Right(_) =>
            IO.raiseError(new Exception("cell command requires single cell, not range"))
        cell = targetSheet.cells.get(ref)
        value = cell.map(_.value).getOrElse(CellValue.Empty)
        // Get style from registry for NumFmt formatting
        style = cell.flatMap(_.styleId).flatMap(targetSheet.styleRegistry.get)
        numFmt = style.map(_.numFmt).getOrElse(NumFmt.General)
        // For formulas with cached values, format the cached value
        valueToFormat = value match
          case CellValue.Formula(_, Some(cached)) => cached
          case other => other
        formatted = NumFmtFormatter.formatValue(valueToFormat, numFmt)
        // Get comment from sheet (sheet.getComment, not cell.comment)
        comment = targetSheet.getComment(ref)
        // Get hyperlink from cell
        hyperlink = cell.flatMap(_.hyperlink)
        // Build dependency graph for dependencies/dependents
        graph = DependencyGraph.fromSheet(targetSheet)
        deps = graph.dependencies.getOrElse(ref, Set.empty).toVector.sortBy(_.toA1)
        dependents = graph.dependents.getOrElse(ref, Set.empty).toVector.sortBy(_.toA1)
      yield Format.cellInfo(ref, value, formatted, style, comment, hyperlink, deps, dependents)

    case Command.Search(pattern, limit, allSheets) =>
      IO.fromEither(
        scala.util
          .Try(pattern.r)
          .toEither
          .left
          .map(e => new Exception(s"Invalid regex pattern: ${e.getMessage}"))
      ).map { regex =>
        if allSheets then
          // Search across all sheets, return qualified refs
          val results = wb.sheets.iterator
            .flatMap { s =>
              s.cells.iterator
                .filter { case (_, cell) =>
                  val text = formatCellValue(cell.value)
                  regex.findFirstIn(text).isDefined
                }
                .map { case (ref, cell) =>
                  val value = formatCellValue(cell.value)
                  (s"${s.name.value}!${ref.toA1}", value)
                }
            }
            .take(limit)
            .toVector
          s"Found ${results.size} matches across ${wb.sheets.size} sheets:\n\n${Markdown.renderSearchResultsWithRef(results)}"
        else
          // Search single sheet, return regular refs
          val results = sheet.cells.iterator
            .filter { case (_, cell) =>
              val text = formatCellValue(cell.value)
              regex.findFirstIn(text).isDefined
            }
            .take(limit)
            .map { case (ref, cell) =>
              val value = formatCellValue(cell.value)
              (ref, value, value)
            }
            .toVector
          Markdown.renderSearchResults(results)
      }

    case Command.Eval(formulaStr, overrides) =>
      for
        tempSheet <- applyOverrides(sheet, overrides)
        formula = if formulaStr.startsWith("=") then formulaStr else s"=$formulaStr"
        result <- IO.fromEither(
          SheetEvaluator.evaluateFormula(tempSheet)(formula).left.map(e => new Exception(e.message))
        )
      yield Format.evalSuccess(formula, result, overrides)

    case Command.Put(refStr, valueStr) =>
      // outputOpt guaranteed Some by CLI parsing for write commands
      outputOpt.fold(IO.raiseError[String](new Exception("Internal: output required"))) {
        outputPath =>
          for
            resolved <- resolveRef(wb, sheet, refStr)
            (targetSheet, refOrRange) = resolved
            ref <- refOrRange match
              case Left(r) => IO.pure(r)
              case Right(_) =>
                IO.raiseError(new Exception("put command requires single cell, not range"))
            value = parseValue(valueStr)
            updatedSheet = targetSheet.put(ref, value)
            updatedWb = wb.put(updatedSheet)
            _ <- ExcelIO.instance[IO].write(updatedWb, outputPath)
          yield s"${Format.putSuccess(ref, value)}\nSaved: $outputPath"
      }

    case Command.PutFormula(refStr, formulaStr) =>
      // outputOpt guaranteed Some by CLI parsing for write commands
      outputOpt.fold(IO.raiseError[String](new Exception("Internal: output required"))) {
        outputPath =>
          for
            resolved <- resolveRef(wb, sheet, refStr)
            (targetSheet, refOrRange) = resolved
            ref <- refOrRange match
              case Left(r) => IO.pure(r)
              case Right(_) =>
                IO.raiseError(new Exception("putf command requires single cell, not range"))
            formula = if formulaStr.startsWith("=") then formulaStr.drop(1) else formulaStr
            value = CellValue.Formula(formula)
            updatedSheet = targetSheet.put(ref, value)
            updatedWb = wb.put(updatedSheet)
            _ <- ExcelIO.instance[IO].write(updatedWb, outputPath)
          yield s"${Format.putSuccess(ref, value)}\nSaved: $outputPath"
      }

    case Command.Style(
          rangeStr,
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
          borderColor
        ) =>
      outputOpt.fold(IO.raiseError[String](new Exception("Internal: output required"))) {
        outputPath =>
          for
            resolved <- resolveRef(wb, sheet, rangeStr)
            (targetSheet, refOrRange) = resolved
            range = refOrRange match
              case Right(r) => r
              case Left(ref) => CellRange(ref, ref)
            style <- buildCellStyle(
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
              borderColor
            )
            updatedSheet = styleSyntax.withRangeStyle(targetSheet)(range, style)
            updatedWb = wb.put(updatedSheet)
            _ <- ExcelIO.instance[IO].write(updatedWb, outputPath)
            appliedList = buildStyleDescription(
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
          yield s"Styled: ${range.toA1}\nApplied: ${appliedList.mkString(", ")}\nSaved: $outputPath"
      }

    case Command.RowOp(rowNum, height, hide, show) =>
      outputOpt.fold(IO.raiseError[String](new Exception("Internal: output required"))) {
        outputPath =>
          val row = Row.from1(rowNum)
          val currentProps = sheet.getRowProperties(row)
          val newProps = currentProps.copy(
            height = height.orElse(currentProps.height),
            hidden = if hide then true else if show then false else currentProps.hidden
          )
          val updatedSheet = sheet.setRowProperties(row, newProps)
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

    case Command.ColOp(colStr, width, hide, show) =>
      outputOpt.fold(IO.raiseError[String](new Exception("Internal: output required"))) {
        outputPath =>
          IO.fromEither(Column.fromLetter(colStr).left.map(e => new Exception(e))).flatMap { col =>
            val currentProps = sheet.getColumnProperties(col)
            val newProps = currentProps.copy(
              width = width.orElse(currentProps.width),
              hidden = if hide then true else if show then false else currentProps.hidden
            )
            val updatedSheet = sheet.setColumnProperties(col, newProps)
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

    case Command.Batch(source) =>
      outputOpt.fold(IO.raiseError[String](new Exception("Internal: output required"))) {
        outputPath =>
          readBatchInput(source).flatMap { input =>
            parseBatchOperations(input).flatMap { ops =>
              applyBatchOperations(wb, sheet, ops).flatMap { updatedWb =>
                ExcelIO.instance[IO].write(updatedWb, outputPath).map { _ =>
                  val summary = ops
                    .map {
                      case BatchOp.Put(ref, value) => s"  PUT $ref = $value"
                      case BatchOp.PutFormula(ref, formula) => s"  PUTF $ref = $formula"
                    }
                    .mkString("\n")
                  s"Applied ${ops.size} operations:\n$summary\nSaved: $outputPath"
                }
              }
            }
          }
      }

  // ==========================================================================
  // Batch operations
  // ==========================================================================

  private enum BatchOp:
    case Put(ref: String, value: String)
    case PutFormula(ref: String, formula: String)

  private def readBatchInput(source: String): IO[String] =
    if source == "-" then IO.blocking(scala.io.Source.stdin.mkString)
    else IO.blocking(scala.io.Source.fromFile(source).mkString)

  /**
   * Parse batch JSON input. Expects format:
   * {{{
   * [
   *   {"op": "put", "ref": "A1", "value": "Hello"},
   *   {"op": "putf", "ref": "B1", "value": "=A1*2"}
   * ]
   * }}}
   */
  private def parseBatchOperations(input: String): IO[Vector[BatchOp]] =
    IO.fromEither {
      val trimmed = input.trim
      if !trimmed.startsWith("[") then Left(new Exception("Batch input must be a JSON array"))
      else parseBatchJson(trimmed)
    }

  /**
   * Simple JSON parser for batch operations. Handles: [{"op":"put"|"putf", "ref":"A1",
   * "value":"..."}]
   */
  private def parseBatchJson(json: String): Either[Exception, Vector[BatchOp]] =
    // Very simple JSON parsing - handles the specific format we need
    val objPattern = """\{[^}]+\}""".r
    val opPattern = """"op"\s*:\s*"(\w+)"""".r
    val refPattern = """"ref"\s*:\s*"([^"]+)"""".r
    val valuePattern = """"value"\s*:\s*"((?:[^"\\]|\\.)*)"""".r

    val ops = objPattern.findAllIn(json).toVector.map { obj =>
      val op = opPattern.findFirstMatchIn(obj).map(_.group(1))
      val ref = refPattern.findFirstMatchIn(obj).map(_.group(1))
      val value = valuePattern
        .findFirstMatchIn(obj)
        .map(_.group(1))
        .map(_.replace("\\\"", "\"").replace("\\n", "\n").replace("\\\\", "\\"))

      (op, ref, value) match
        case (Some("put"), Some(r), Some(v)) => Right(BatchOp.Put(r, v))
        case (Some("putf"), Some(r), Some(v)) => Right(BatchOp.PutFormula(r, v))
        case (Some(unknown), _, _) => Left(new Exception(s"Unknown operation: $unknown"))
        case _ => Left(new Exception(s"Invalid batch operation: $obj"))
    }

    val errors = ops.collect { case Left(e) => e }
    errors.headOption match
      case Some(e) => Left(e)
      case None => Right(ops.collect { case Right(op) => op })

  private def applyBatchOperations(
    wb: Workbook,
    sheet: Sheet,
    ops: Vector[BatchOp]
  ): IO[Workbook] =
    ops
      .foldLeft(IO.pure(sheet)) { (sheetIO, op) =>
        sheetIO.flatMap { s =>
          op match
            case BatchOp.Put(refStr, value) =>
              IO.fromEither(ARef.parse(refStr).left.map(e => new Exception(e))).map { ref =>
                val cellValue = parseValue(value)
                s.put(ref, cellValue)
              }
            case BatchOp.PutFormula(refStr, formula) =>
              IO.fromEither(ARef.parse(refStr).left.map(e => new Exception(e))).map { ref =>
                val normalizedFormula = if formula.startsWith("=") then formula.drop(1) else formula
                s.put(ref, CellValue.Formula(normalizedFormula, None))
              }
        }
      }
      .map(updatedSheet => wb.put(updatedSheet))

  // ==========================================================================
  // Helpers
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
                val value = parseValue(valueStr.trim)
                IO.pure(s.put(ref, value))
              case RefType.QualifiedCell(sheetName, ref) =>
                if sheetName == sheet.name then
                  val value = parseValue(valueStr.trim)
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

  private def parseValue(s: String): CellValue =
    scala.util.Try(BigDecimal(s)).toOption.map(CellValue.Number.apply).getOrElse {
      s.toLowerCase match
        case "true" => CellValue.Bool(true)
        case "false" => CellValue.Bool(false)
        case _ =>
          val text = if s.startsWith("\"") && s.endsWith("\"") then s.drop(1).dropRight(1) else s
          CellValue.Text(text)
    }

  private def formatCellValue(value: CellValue): String =
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
      case CellValue.Formula(expr, cached) =>
        val displayExpr = if expr.startsWith("=") then expr else s"=$expr"
        cached.map(formatCellValue).getOrElse(displayExpr)

  private def buildCellStyle(
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
    borderColor: Option[String]
  ): IO[CellStyle] =
    for
      bgColor <- bg.traverse(s => IO.fromEither(ColorParser.parse(s).left.map(new Exception(_))))
      fgColor <- fg.traverse(s => IO.fromEither(ColorParser.parse(s).left.map(new Exception(_))))
      bdrColor <- borderColor.traverse(s =>
        IO.fromEither(ColorParser.parse(s).left.map(new Exception(_)))
      )
      hAlign <- align.traverse(s => IO.fromEither(parseHAlign(s).left.map(new Exception(_))))
      vAlign <- valign.traverse(s => IO.fromEither(parseVAlign(s).left.map(new Exception(_))))
      bdrStyle <- border.traverse(s =>
        IO.fromEither(parseBorderStyle(s).left.map(new Exception(_)))
      )
      nFmt <- numFormat.traverse(s => IO.fromEither(parseNumFmt(s).left.map(new Exception(_))))
    yield
      val font = Font.default
        .withBold(bold)
        .withItalic(italic)
        .withUnderline(underline)
        .pipe(f => fgColor.fold(f)(c => f.withColor(c)))
        .pipe(f => fontSize.fold(f)(s => f.withSize(s)))
        .pipe(f => fontName.fold(f)(n => f.withName(n)))

      val fill = bgColor.map(Fill.Solid.apply).getOrElse(Fill.None)

      val cellBorder = bdrStyle
        .map(style => Border.all(style, bdrColor))
        .getOrElse(Border.none)

      val alignment = Align.default
        .pipe(a => hAlign.fold(a)(h => a.withHAlign(h)))
        .pipe(a => vAlign.fold(a)(v => a.withVAlign(v)))
        .pipe(a => if wrap then a.withWrap() else a)

      CellStyle(
        font = font,
        fill = fill,
        border = cellBorder,
        numFmt = nFmt.getOrElse(NumFmt.General),
        align = alignment
      )

  private def parseHAlign(s: String): Either[String, HAlign] =
    s.toLowerCase match
      case "left" => Right(HAlign.Left)
      case "center" => Right(HAlign.Center)
      case "right" => Right(HAlign.Right)
      case "justify" => Right(HAlign.Justify)
      case "general" => Right(HAlign.General)
      case other => Left(s"Unknown horizontal alignment: $other. Use left, center, right, justify")

  private def parseVAlign(s: String): Either[String, VAlign] =
    s.toLowerCase match
      case "top" => Right(VAlign.Top)
      case "middle" | "center" => Right(VAlign.Middle)
      case "bottom" => Right(VAlign.Bottom)
      case other => Left(s"Unknown vertical alignment: $other. Use top, middle, bottom")

  private def parseBorderStyle(s: String): Either[String, BorderStyle] =
    s.toLowerCase match
      case "none" => Right(BorderStyle.None)
      case "thin" => Right(BorderStyle.Thin)
      case "medium" => Right(BorderStyle.Medium)
      case "thick" => Right(BorderStyle.Thick)
      case "dashed" => Right(BorderStyle.Dashed)
      case "dotted" => Right(BorderStyle.Dotted)
      case "double" => Right(BorderStyle.Double)
      case other => Left(s"Unknown border style: $other. Use none, thin, medium, thick")

  private def parseNumFmt(s: String): Either[String, NumFmt] =
    s.toLowerCase match
      case "general" => Right(NumFmt.General)
      case "number" => Right(NumFmt.Decimal)
      case "currency" => Right(NumFmt.Currency)
      case "percent" => Right(NumFmt.Percent)
      case "date" => Right(NumFmt.Date)
      case "text" => Right(NumFmt.Text)
      case other =>
        Left(s"Unknown number format: $other. Use general, number, currency, percent, date, text")

  private def buildStyleDescription(
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
    border: Option[String]
  ): List[String] =
    List(
      if bold then Some("bold") else None,
      if italic then Some("italic") else None,
      if underline then Some("underline") else None,
      bg.map(c => s"bg=$c"),
      fg.map(c => s"fg=$c"),
      fontSize.map(s => s"font-size=$s"),
      fontName.map(n => s"font-name=$n"),
      align.map(a => s"align=$a"),
      valign.map(v => s"valign=$v"),
      if wrap then Some("wrap") else None,
      numFormat.map(f => s"format=$f"),
      border.map(b => s"border=$b")
    ).flatten

/**
 * Command ADT representing all CLI operations.
 */
enum Command:
  // Read-only
  case Sheets
  case Bounds
  case View(
    range: String,
    showFormulas: Boolean,
    limit: Int,
    format: ViewFormat,
    printScale: Boolean,
    showGridlines: Boolean,
    showLabels: Boolean,
    dpi: Int,
    quality: Int,
    rasterOutput: Option[Path],
    skipEmpty: Boolean,
    headerRow: Option[Int]
  )
  case Cell(ref: String)
  case Search(pattern: String, limit: Int, allSheets: Boolean)
  // Analyze
  case Eval(formula: String, overrides: List[String])
  // Mutate (require -o)
  case Put(ref: String, value: String)
  case PutFormula(ref: String, formula: String)
  case Style(
    range: String,
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
    borderColor: Option[String]
  )
  case RowOp(row: Int, height: Option[Double], hide: Boolean, show: Boolean)
  case ColOp(col: String, width: Option[Double], hide: Boolean, show: Boolean)
  case Batch(source: String) // "-" for stdin or file path

/**
 * Output format for view command.
 */
enum ViewFormat:
  case Markdown, Html, Svg, Json, Csv
  case Png, Jpeg, WebP, Pdf
