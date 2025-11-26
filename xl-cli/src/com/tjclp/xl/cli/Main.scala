package com.tjclp.xl.cli

import java.nio.file.Path

import cats.effect.{ExitCode, IO, IOApp}
import cats.implicits.*
import com.monovore.decline.*
import com.monovore.decline.effect.*

import com.tjclp.xl.{*, given}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.addressing.{ARef, CellRange, RefType, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.output.{Format, Markdown}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.formula.{DependencyGraph, SheetEvaluator}
import com.tjclp.xl.styles.numfmt.NumFmt

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
      version = "0.1.0"
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
    val sheetWriteSubcmds = putCmd orElse putfCmd

    val sheetWriteOpts =
      (fileOpt, sheetOpt, outputOpt, sheetWriteSubcmds).mapN { (file, sheet, out, cmd) =>
        run(file, sheet, Some(out), cmd)
      }

    workbookOpts orElse sheetReadOnlyOpts orElse sheetWriteOpts

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
    .option[String]("format", "Output format: markdown, html, or svg")
    .withDefault("markdown")
    .mapValidated { s =>
      s.toLowerCase match
        case "markdown" | "md" => cats.data.Validated.valid(ViewFormat.Markdown)
        case "html" => cats.data.Validated.valid(ViewFormat.Html)
        case "svg" => cats.data.Validated.valid(ViewFormat.Svg)
        case other =>
          cats.data.Validated.invalidNel(s"Unknown format: $other. Use markdown, html, or svg")
    }

  // --- Read-only commands ---

  val sheetsCmd: Opts[Command] = Opts.subcommand("sheets", "List all sheets") {
    Opts(Command.Sheets)
  }

  val boundsCmd: Opts[Command] = Opts.subcommand("bounds", "Show used range of current sheet") {
    Opts(Command.Bounds)
  }

  val viewCmd: Opts[Command] = Opts.subcommand("view", "View range (markdown, html, or svg)") {
    (rangeArg, formulasOpt, limitOpt, formatOpt).mapN(Command.View.apply)
  }

  val cellCmd: Opts[Command] = Opts.subcommand("cell", "Get cell details") {
    refArg.map(Command.Cell.apply)
  }

  val searchCmd: Opts[Command] = Opts.subcommand("search", "Search for cells") {
    (patternArg, limitOpt).mapN(Command.Search.apply)
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

    case Command.View(rangeStr, showFormulas, limit, format) =>
      import com.tjclp.xl.sheets.styleSyntax.*
      for
        resolved <- resolveRef(wb, sheet, rangeStr)
        (targetSheet, refOrRange) = resolved
        range = refOrRange match
          case Right(r) => r
          case Left(ref) => CellRange(ref, ref) // Single cell as range
        limitedRange = limitRange(range, limit)
      yield format match
        case ViewFormat.Markdown => Markdown.renderRange(targetSheet, limitedRange, showFormulas)
        case ViewFormat.Html => targetSheet.toHtml(limitedRange)
        case ViewFormat.Svg => targetSheet.toSvg(limitedRange)

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

    case Command.Search(pattern, limit) =>
      IO.fromEither(
        scala.util
          .Try(pattern.r)
          .toEither
          .left
          .map(e => new Exception(s"Invalid regex pattern: ${e.getMessage}"))
      ).map { regex =>
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
        cached.map(formatCellValue).getOrElse(s"=$expr")

/**
 * Command ADT representing all CLI operations.
 */
enum Command:
  // Read-only
  case Sheets
  case Bounds
  case View(range: String, showFormulas: Boolean, limit: Int, format: ViewFormat)
  case Cell(ref: String)
  case Search(pattern: String, limit: Int)
  // Analyze
  case Eval(formula: String, overrides: List[String])
  // Mutate (require -o)
  case Put(ref: String, value: String)
  case PutFormula(ref: String, formula: String)

/**
 * Output format for view command.
 */
enum ViewFormat:
  case Markdown, Html, Svg
