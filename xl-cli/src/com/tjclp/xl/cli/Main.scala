package com.tjclp.xl.cli

import java.nio.file.Path

import cats.effect.{ExitCode, IO, IOApp}
import cats.implicits.*
import com.monovore.decline.*
import com.monovore.decline.effect.*

import com.tjclp.xl.{*, given}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.output.{Format, Markdown}
import com.tjclp.xl.formula.SheetEvaluator

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
    val subcommands =
      sheetsCmd orElse boundsCmd orElse
        viewCmd orElse cellCmd orElse searchCmd orElse evalCmd orElse
        putCmd orElse putfCmd

    (fileOpt, sheetOpt, outputOpt, subcommands).mapN { (filePath, sheetName, outputPath, cmd) =>
      run(filePath, sheetName, outputPath, cmd)
    }

  // ==========================================================================
  // Global options
  // ==========================================================================

  private val fileOpt =
    Opts.option[Path]("file", "Excel file to operate on (required)", "f")

  private val sheetOpt =
    Opts.option[String]("sheet", "Sheet to select (optional, defaults to first)", "s").orNone

  private val outputOpt =
    Opts.option[Path]("output", "Output file for mutations (required for put/putf)", "o").orNone

  // ==========================================================================
  // Command definitions
  // ==========================================================================

  private val rangeArg = Opts.argument[String]("range")
  private val refArg = Opts.argument[String]("ref")
  private val valueArg = Opts.argument[String]("value")
  private val patternArg = Opts.argument[String]("pattern")

  private val formulasOpt = Opts.flag("formulas", "Show formulas instead of values").orFalse
  private val limitOpt = Opts.option[Int]("limit", "Maximum rows to display").withDefault(50)

  // --- Read-only commands ---

  val sheetsCmd: Opts[Command] = Opts.subcommand("sheets", "List all sheets") {
    Opts(Command.Sheets)
  }

  val boundsCmd: Opts[Command] = Opts.subcommand("bounds", "Show used range of current sheet") {
    Opts(Command.Bounds)
  }

  val viewCmd: Opts[Command] = Opts.subcommand("view", "View range as markdown table") {
    (rangeArg, formulasOpt, limitOpt).mapN(Command.View.apply)
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

  val putCmd: Opts[Command] = Opts.subcommand("put", "Write value to cell (requires -o)") {
    (refArg, valueArg).mapN(Command.Put.apply)
  }

  val putfCmd: Opts[Command] = Opts.subcommand("putf", "Write formula to cell (requires -o)") {
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

    case Command.View(rangeStr, showFormulas, limit) =>
      for
        range <- IO.fromEither(CellRange.parse(rangeStr).left.map(e => new Exception(e.toString)))
        limitedRange = limitRange(range, limit)
      yield Markdown.renderRange(sheet, limitedRange, showFormulas)

    case Command.Cell(refStr) =>
      for
        ref <- IO.fromEither(ARef.parse(refStr).left.map(e => new Exception(e.toString)))
        cell = sheet.cells.get(ref)
        value = cell.map(_.value).getOrElse(CellValue.Empty)
        formatted = formatCellValue(value)
        deps = Vector.empty[ARef]
        dependents = Vector.empty[ARef]
      yield Format.cellInfo(ref, value, formatted, deps, dependents)

    case Command.Search(pattern, limit) =>
      val regex = pattern.r
      val results = sheet.cells.toVector
        .filter { case (_, cell) =>
          val text = formatCellValue(cell.value)
          regex.findFirstIn(text).isDefined
        }
        .take(limit)
        .map { case (ref, cell) =>
          val value = formatCellValue(cell.value)
          (ref, value, value)
        }
      IO.pure(Markdown.renderSearchResults(results))

    case Command.Eval(formulaStr, overrides) =>
      for
        tempSheet <- applyOverrides(sheet, overrides)
        formula = if formulaStr.startsWith("=") then formulaStr else s"=$formulaStr"
        result <- IO.fromEither(
          SheetEvaluator.evaluateFormula(tempSheet)(formula).left.map(e => new Exception(e.message))
        )
      yield Format.evalSuccess(formula, result, overrides)

    case Command.Put(refStr, valueStr) =>
      for
        outputPath <- IO.fromOption(outputOpt)(
          new Exception("Missing required flag: --output (-o). Mutations require an output file.")
        )
        ref <- IO.fromEither(ARef.parse(refStr).left.map(e => new Exception(e.toString)))
        value = parseValue(valueStr)
        updatedSheet = sheet.put(ref, value)
        updatedWb <- IO.fromEither(wb.put(updatedSheet).left.map(e => new Exception(e.message)))
        _ <- ExcelIO.instance[IO].write(updatedWb, outputPath)
      yield s"${Format.putSuccess(ref, value)}\nSaved: $outputPath"

    case Command.PutFormula(refStr, formulaStr) =>
      for
        outputPath <- IO.fromOption(outputOpt)(
          new Exception("Missing required flag: --output (-o). Mutations require an output file.")
        )
        ref <- IO.fromEither(ARef.parse(refStr).left.map(e => new Exception(e.toString)))
        formula = if formulaStr.startsWith("=") then formulaStr.drop(1) else formulaStr
        value = CellValue.Formula(formula)
        updatedSheet = sheet.put(ref, value)
        updatedWb <- IO.fromEither(wb.put(updatedSheet).left.map(e => new Exception(e.message)))
        _ <- ExcelIO.instance[IO].write(updatedWb, outputPath)
      yield s"${Format.putSuccess(ref, value)}\nSaved: $outputPath"

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
          case Array(refStr, valueStr) =>
            IO.fromEither(ARef.parse(refStr.trim).left.map(e => new Exception(e.toString)))
              .map { ref =>
                val value = parseValue(valueStr.trim)
                s.put(ref, value)
              }
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
  case View(range: String, showFormulas: Boolean, limit: Int)
  case Cell(ref: String)
  case Search(pattern: String, limit: Int)
  // Analyze
  case Eval(formula: String, overrides: List[String])
  // Mutate (require -o)
  case Put(ref: String, value: String)
  case PutFormula(ref: String, formula: String)
