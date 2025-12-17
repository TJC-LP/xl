package com.tjclp.xl.cli

import java.nio.file.Path

import cats.effect.{ExitCode, IO}
import cats.implicits.*
import com.monovore.decline.*
import com.monovore.decline.effect.*

import com.tjclp.xl.{*, given}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cli.commands.{
  CellCommands,
  ReadCommands,
  SheetCommands,
  WorkbookCommands,
  WriteCommands
}
import com.tjclp.xl.cli.helpers.SheetResolver
import com.tjclp.xl.cli.output.Format

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
    val workbookSubcmds = sheetsCmd orElse namesCmd
    val workbookOpts = (fileOpt, workbookSubcmds).mapN { (file, cmd) =>
      run(file, None, None, cmd)
    }

    // Sheet-level read-only: --file and --sheet (no --output)
    val sheetReadOnlySubcmds =
      boundsCmd orElse viewCmd orElse cellCmd orElse searchCmd orElse statsCmd orElse evalCmd

    val sheetReadOnlyOpts = (fileOpt, sheetOpt, sheetReadOnlySubcmds).mapN { (file, sheet, cmd) =>
      run(file, sheet, None, cmd)
    }

    // Sheet-level write: --file, --sheet, and --output (required)
    val sheetWriteSubcmds =
      putCmd orElse putfCmd orElse styleCmd orElse rowCmd orElse colCmd orElse batchCmd orElse addSheetCmd orElse removeSheetCmd orElse renameSheetCmd orElse moveSheetCmd orElse copySheetCmd orElse mergeCmd orElse unmergeCmd

    val sheetWriteOpts =
      (fileOpt, sheetOpt, outputOpt, sheetWriteSubcmds).mapN { (file, sheet, out, cmd) =>
        run(file, sheet, Some(out), cmd)
      }

    // Standalone: no --file required (creates new files)
    val standaloneOpts = newCmd.map { case (outPath, sheetName, sheets) =>
      runStandalone(outPath, sheetName, sheets)
    }

    standaloneOpts orElse workbookOpts orElse sheetReadOnlyOpts orElse sheetWriteOpts

  // ==========================================================================
  // Global options
  // ==========================================================================

  private val fileOpt =
    Opts.option[Path]("file", "Excel file to operate on (required)", "f")

  private val sheetOpt =
    Opts
      .option[String]("sheet", "Sheet to select (required for sheet-level operations)", "s")
      .orNone

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
  private val evalOpt = Opts.flag("eval", "Evaluate formulas (compute live values)").orFalse
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
  private val useImageMagickOpt =
    Opts
      .flag(
        "use-imagemagick",
        "Use ImageMagick for rasterization instead of Batik (requires magick/convert in PATH)"
      )
      .orFalse
  private val sheetsFilterOpt =
    Opts
      .option[String]("sheets", "Comma-separated list of sheets to search (default: all)")
      .orNone

  // --- Standalone commands (no --file required) ---

  private val outputArg = Opts.argument[Path]("output")
  private val sheetNameOpt =
    Opts.option[String]("sheet-name", "Sheet name (defaults to 'Sheet1')").withDefault("Sheet1")
  private val sheetsOpt: Opts[List[String]] =
    Opts.options[String]("sheet", "Sheet name (repeatable for multiple sheets)").orEmpty

  val newCmd: Opts[(Path, String, List[String])] =
    Opts.subcommand("new", "Create a blank xlsx file") {
      (outputArg, sheetNameOpt, sheetsOpt).tupled
    }

  // --- Read-only commands ---

  val sheetsCmd: Opts[CliCommand] = Opts.subcommand("sheets", "List all sheets") {
    Opts(CliCommand.Sheets)
  }

  val namesCmd: Opts[CliCommand] = Opts.subcommand("names", "List defined names (named ranges)") {
    Opts(CliCommand.Names)
  }

  val boundsCmd: Opts[CliCommand] = Opts.subcommand("bounds", "Show used range of current sheet") {
    Opts(CliCommand.Bounds)
  }

  val viewCmd: Opts[CliCommand] =
    Opts.subcommand("view", "View range (markdown, html, svg, json, csv, png, jpeg, webp, pdf)") {
      (
        rangeArg,
        formulasOpt,
        evalOpt,
        limitOpt,
        formatOpt,
        printScaleOpt,
        gridlinesOpt,
        showLabelsOpt,
        dpiOpt,
        qualityOpt,
        rasterOutputOpt,
        skipEmptyOpt,
        headerRowOpt,
        useImageMagickOpt
      )
        .mapN(CliCommand.View.apply)
    }

  private val noStyleOpt =
    Opts.flag("no-style", "Omit style information from output").orFalse

  val cellCmd: Opts[CliCommand] = Opts.subcommand("cell", "Get cell details") {
    (refArg, noStyleOpt).mapN(CliCommand.Cell.apply)
  }

  val searchCmd: Opts[CliCommand] =
    Opts.subcommand("search", "Search for cells (all sheets by default)") {
      (patternArg, limitOpt, sheetsFilterOpt).mapN(CliCommand.Search.apply)
    }

  val statsCmd: Opts[CliCommand] =
    Opts.subcommand("stats", "Calculate statistics for numeric values in range") {
      rangeArg.map(CliCommand.Stats.apply)
    }

  // --- Analyze ---

  private val formulaArg = Opts.argument[String]("formula")
  private val withOpt =
    Opts.option[String]("with", "Comma-separated overrides (e.g., A1=100,B2=200)", "w").orNone

  val evalCmd: Opts[CliCommand] =
    Opts.subcommand("eval", "Evaluate formula without modifying sheet") {
      (formulaArg, withOpt).mapN { (formula, withStr) =>
        val overrides = withStr.toList.flatMap(_.split(",").map(_.trim).filter(_.nonEmpty))
        CliCommand.Eval(formula, overrides)
      }
    }

  // --- Mutate (require -o) ---

  val putCmd: Opts[CliCommand] = Opts.subcommand("put", "Write value to cell") {
    (refArg, valueArg).mapN(CliCommand.Put.apply)
  }

  val putfCmd: Opts[CliCommand] = Opts.subcommand("putf", "Write formula to cell") {
    (refArg, valueArg).mapN(CliCommand.PutFormula.apply)
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
    Opts.option[String]("border", "Border style for all sides: none, thin, medium, thick").orNone
  private val borderTopOpt =
    Opts.option[String]("border-top", "Top border style: none, thin, medium, thick").orNone
  private val borderRightOpt =
    Opts.option[String]("border-right", "Right border style: none, thin, medium, thick").orNone
  private val borderBottomOpt =
    Opts.option[String]("border-bottom", "Bottom border style: none, thin, medium, thick").orNone
  private val borderLeftOpt =
    Opts.option[String]("border-left", "Left border style: none, thin, medium, thick").orNone
  private val borderColorOpt = Opts.option[String]("border-color", "Border color").orNone
  private val replaceOpt = Opts
    .flag("replace", "Replace entire style instead of merging with existing")
    .orFalse

  val styleCmd: Opts[CliCommand] = Opts.subcommand("style", "Apply styling to cells") {
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
      borderTopOpt,
      borderRightOpt,
      borderBottomOpt,
      borderLeftOpt,
      borderColorOpt,
      replaceOpt
    ).mapN(CliCommand.Style.apply)
  }

  // --- Row/Column command options ---
  private val rowArg = Opts.argument[Int]("row")
  private val colArg = Opts.argument[String]("col")
  private val heightOpt = Opts.option[Double]("height", "Row height in points").orNone
  private val widthOpt = Opts.option[Double]("width", "Column width in character units").orNone
  private val hideOpt = Opts.flag("hide", "Hide row/column").orFalse
  private val showOpt = Opts.flag("show", "Show (unhide) row/column").orFalse

  val rowCmd: Opts[CliCommand] = Opts.subcommand("row", "Set row properties (height, hide/show)") {
    (rowArg, heightOpt, hideOpt, showOpt).mapN(CliCommand.RowOp.apply)
  }

  val colCmd: Opts[CliCommand] =
    Opts.subcommand("col", "Set column properties (width, hide/show)") {
      (colArg, widthOpt, hideOpt, showOpt).mapN(CliCommand.ColOp.apply)
    }

  // --- Batch command ---
  private val batchArg = Opts.argument[String]("operations").withDefault("-")
  val batchCmd: Opts[CliCommand] =
    Opts.subcommand("batch", "Apply multiple operations atomically (JSON from stdin or file)") {
      batchArg.map(CliCommand.Batch.apply)
    }

  // --- Sheet management commands ---
  private val sheetNameArg = Opts.argument[String]("name")
  private val afterOpt =
    Opts.option[String]("after", "Insert new sheet after this sheet").orNone
  private val beforeOpt =
    Opts.option[String]("before", "Insert new sheet before this sheet").orNone

  val addSheetCmd: Opts[CliCommand] =
    Opts.subcommand("add-sheet", "Add new empty sheet to workbook") {
      (sheetNameArg, afterOpt, beforeOpt).mapN(CliCommand.AddSheet.apply)
    }

  val removeSheetCmd: Opts[CliCommand] =
    Opts.subcommand("remove-sheet", "Remove sheet from workbook") {
      sheetNameArg.map(CliCommand.RemoveSheet.apply)
    }

  private val newNameArg = Opts.argument[String]("new-name")

  val renameSheetCmd: Opts[CliCommand] =
    Opts.subcommand("rename-sheet", "Rename a sheet") {
      (sheetNameArg, newNameArg).mapN(CliCommand.RenameSheet.apply)
    }

  private val toIndexOpt =
    Opts.option[Int]("to", "Move to index (0-based)").orNone

  val moveSheetCmd: Opts[CliCommand] =
    Opts.subcommand("move-sheet", "Move sheet to new position") {
      (sheetNameArg, toIndexOpt, afterOpt, beforeOpt).mapN(CliCommand.MoveSheet.apply)
    }

  val copySheetCmd: Opts[CliCommand] =
    Opts.subcommand("copy-sheet", "Copy sheet to new name") {
      (sheetNameArg, newNameArg).mapN(CliCommand.CopySheet.apply)
    }

  val mergeCmd: Opts[CliCommand] =
    Opts.subcommand("merge", "Merge cells in range") {
      rangeArg.map(CliCommand.Merge.apply)
    }

  val unmergeCmd: Opts[CliCommand] =
    Opts.subcommand("unmerge", "Unmerge cells in range") {
      rangeArg.map(CliCommand.Unmerge.apply)
    }

  // ==========================================================================
  // Command execution
  // ==========================================================================

  private def run(
    filePath: Path,
    sheetNameOpt: Option[String],
    outputOpt: Option[Path],
    cmd: CliCommand
  ): IO[ExitCode] =
    execute(filePath, sheetNameOpt, outputOpt, cmd).attempt.flatMap {
      case Right(output) =>
        IO.println(output).as(ExitCode.Success)
      case Left(err) =>
        IO.println(Format.errorSimple(err.getMessage)).as(ExitCode.Error)
    }

  private def runStandalone(outPath: Path, sheetName: String, sheets: List[String]): IO[ExitCode] =
    (for
      // --sheet takes precedence over --sheet-name; if neither, default to "Sheet1"
      names <- sheets match
        case Nil =>
          IO.fromEither(SheetName(sheetName).left.map(e => new Exception(e))).map(List(_))
        case list =>
          list.traverse(n => IO.fromEither(SheetName(n).left.map(e => new Exception(e))))
      wb = Workbook(names.map(Sheet(_)).toVector)
      _ <- ExcelIO.instance[IO].write(wb, outPath)
    yield
      val sheetList = names.map(_.value).mkString(", ")
      s"Created ${outPath.toAbsolutePath} with ${names.size} sheet(s): $sheetList"
    ).attempt.flatMap {
      case Right(output) =>
        IO.println(output).as(ExitCode.Success)
      case Left(err) =>
        IO.println(Format.errorSimple(err.getMessage)).as(ExitCode.Error)
    }

  private def execute(
    filePath: Path,
    sheetNameOpt: Option[String],
    outputOpt: Option[Path],
    cmd: CliCommand
  ): IO[String] =
    for
      wb <- ExcelIO.instance[IO].read(filePath)
      sheet <- SheetResolver.resolveSheet(wb, sheetNameOpt)
      result <- executeCommand(wb, sheet, outputOpt, cmd)
    yield result

  private def executeCommand(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    outputOpt: Option[Path],
    cmd: CliCommand
  ): IO[String] = cmd match
    // Workbook commands
    case CliCommand.Sheets =>
      WorkbookCommands.sheets(wb)

    case CliCommand.Names =>
      WorkbookCommands.names(wb)

    // Read commands
    case CliCommand.Bounds =>
      ReadCommands.bounds(wb, sheetOpt)

    case CliCommand.View(
          rangeStr,
          showFormulas,
          evalFormulas,
          limit,
          format,
          printScale,
          showGridlines,
          showLabels,
          dpi,
          quality,
          rasterOutput,
          skipEmpty,
          headerRow,
          useImageMagick
        ) =>
      ReadCommands.view(
        wb,
        sheetOpt,
        rangeStr,
        showFormulas,
        evalFormulas,
        limit,
        format,
        printScale,
        showGridlines,
        showLabels,
        dpi,
        quality,
        rasterOutput,
        skipEmpty,
        headerRow,
        useImageMagick
      )

    case CliCommand.Cell(refStr, noStyle) =>
      ReadCommands.cell(wb, sheetOpt, refStr, noStyle)

    case CliCommand.Search(pattern, limit, sheetsFilter) =>
      ReadCommands.search(wb, sheetOpt, pattern, limit, sheetsFilter)

    case CliCommand.Stats(refStr) =>
      ReadCommands.stats(wb, sheetOpt, refStr)

    case CliCommand.Eval(formulaStr, overrides) =>
      ReadCommands.eval(wb, sheetOpt, formulaStr, overrides)

    // Write commands (require output)
    case CliCommand.Put(refStr, valueStr) =>
      requireOutput(outputOpt)(WriteCommands.put(wb, sheetOpt, refStr, valueStr, _))

    case CliCommand.PutFormula(refStr, formulaStr) =>
      requireOutput(outputOpt)(WriteCommands.putFormula(wb, sheetOpt, refStr, formulaStr, _))

    case CliCommand.Style(
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
          borderTop,
          borderRight,
          borderBottom,
          borderLeft,
          borderColor,
          replace
        ) =>
      requireOutput(outputOpt) { outputPath =>
        WriteCommands.style(
          wb,
          sheetOpt,
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
          borderTop,
          borderRight,
          borderBottom,
          borderLeft,
          borderColor,
          replace,
          outputPath
        )
      }

    case CliCommand.RowOp(rowNum, height, hide, show) =>
      requireOutput(outputOpt)(WriteCommands.row(wb, sheetOpt, rowNum, height, hide, show, _))

    case CliCommand.ColOp(colStr, width, hide, show) =>
      requireOutput(outputOpt)(WriteCommands.col(wb, sheetOpt, colStr, width, hide, show, _))

    case CliCommand.Batch(source) =>
      requireOutput(outputOpt)(WriteCommands.batch(wb, sheetOpt, source, _))

    // Sheet management commands
    case CliCommand.AddSheet(name, afterOpt, beforeOpt) =>
      requireOutput(outputOpt)(SheetCommands.addSheet(wb, name, afterOpt, beforeOpt, _))

    case CliCommand.RemoveSheet(name) =>
      requireOutput(outputOpt)(SheetCommands.removeSheet(wb, name, _))

    case CliCommand.RenameSheet(oldName, newName) =>
      requireOutput(outputOpt)(SheetCommands.renameSheet(wb, oldName, newName, _))

    case CliCommand.MoveSheet(name, toIndexOpt, afterOpt, beforeOpt) =>
      requireOutput(outputOpt)(
        SheetCommands.moveSheet(wb, name, toIndexOpt, afterOpt, beforeOpt, _)
      )

    case CliCommand.CopySheet(sourceName, targetName) =>
      requireOutput(outputOpt)(SheetCommands.copySheet(wb, sourceName, targetName, _))

    // Cell commands
    case CliCommand.Merge(rangeStr) =>
      requireOutput(outputOpt)(CellCommands.merge(wb, sheetOpt, rangeStr, _))

    case CliCommand.Unmerge(rangeStr) =>
      requireOutput(outputOpt)(CellCommands.unmerge(wb, sheetOpt, rangeStr, _))

  // ==========================================================================
  // Helpers
  // ==========================================================================

  private def requireOutput(outputOpt: Option[Path])(f: Path => IO[String]): IO[String] =
    outputOpt.fold(IO.raiseError[String](new Exception("Internal: output required")))(f)
