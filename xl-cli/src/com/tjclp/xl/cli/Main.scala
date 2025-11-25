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
 * Designed for incremental exploration following Claude Code patterns: don't dump everything at
 * once, explore on demand.
 */
object Main
    extends CommandIOApp(
      name = "xl",
      header = "LLM-friendly Excel operations",
      version = "0.1.0"
    ):

  // Session state - held across REPL interactions
  // For single-command invocations, this starts fresh each time
  private var session: Session = Session.empty

  override def main: Opts[IO[ExitCode]] =
    (openCmd orElse createCmd orElse closeCmd orElse
      sheetsCmd orElse selectCmd orElse boundsCmd orElse
      viewCmd orElse cellCmd orElse searchCmd orElse evalCmd orElse
      putCmd orElse putfCmd orElse
      saveCmd orElse saveasCmd).map(run)

  // ==========================================================================
  // Command definitions
  // ==========================================================================

  private val pathArg = Opts.argument[Path]("path")
  private val rangeArg = Opts.argument[String]("range")
  private val refArg = Opts.argument[String]("ref")
  private val valueArg = Opts.argument[String]("value")
  private val sheetArg = Opts.argument[String]("sheet")
  private val patternArg = Opts.argument[String]("pattern")

  private val readonlyOpt = Opts.flag("readonly", "Open in read-only mode").orFalse
  private val formulasOpt = Opts.flag("formulas", "Show formulas instead of values").orFalse
  private val limitOpt = Opts.option[Int]("limit", "Maximum rows to display").withDefault(50)

  // --- Open/Initialize ---

  val openCmd: Opts[Command] = Opts.subcommand("open", "Open an Excel file") {
    (pathArg, readonlyOpt).mapN(Command.Open.apply)
  }

  val createCmd: Opts[Command] = Opts.subcommand("create", "Create new workbook") {
    Opts
      .option[String]("sheets", "Comma-separated sheet names")
      .withDefault("Sheet1")
      .map(s => Command.Create(s.split(",").toVector.map(_.trim)))
  }

  val closeCmd: Opts[Command] = Opts.subcommand("close", "Close current workbook") {
    Opts.flag("discard", "Discard unsaved changes").orFalse.map(Command.Close.apply)
  }

  // --- Navigate ---

  val sheetsCmd: Opts[Command] = Opts.subcommand("sheets", "List all sheets") {
    Opts(Command.Sheets)
  }

  val selectCmd: Opts[Command] = Opts.subcommand("select", "Set active sheet") {
    sheetArg.map(Command.Select.apply)
  }

  val boundsCmd: Opts[Command] = Opts.subcommand("bounds", "Show used range") {
    sheetArg.orNone.map(Command.Bounds.apply)
  }

  // --- Explore ---

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
  private val withOpt = Opts.options[String]("with", "Temporary cell override (ref=value)").orEmpty

  val evalCmd: Opts[Command] = Opts.subcommand("eval", "Evaluate formula without modifying sheet") {
    (formulaArg, withOpt).mapN(Command.Eval.apply)
  }

  // --- Mutate ---

  val putCmd: Opts[Command] = Opts.subcommand("put", "Write value to cell") {
    (refArg, valueArg).mapN(Command.Put.apply)
  }

  val putfCmd: Opts[Command] = Opts.subcommand("putf", "Write formula to cell") {
    (refArg, valueArg).mapN(Command.PutFormula.apply)
  }

  // --- Persist ---

  val saveCmd: Opts[Command] = Opts.subcommand("save", "Save workbook") {
    Opts(Command.Save)
  }

  val saveasCmd: Opts[Command] = Opts.subcommand("saveas", "Save workbook to new path") {
    pathArg.map(Command.SaveAs.apply)
  }

  // ==========================================================================
  // Command execution
  // ==========================================================================

  private def run(cmd: Command): IO[ExitCode] =
    execute(cmd).attempt.flatMap {
      case Right(output) =>
        IO.println(output).as(ExitCode.Success)
      case Left(err) =>
        IO.println(Format.errorSimple(err.getMessage)).as(ExitCode.Error)
    }

  private def execute(cmd: Command): IO[String] = cmd match
    case Command.Open(path, readonly) =>
      for
        _ <- requireNoWorkbook
        wb <- ExcelIO.instance[IO].read(path)
        _ = session = Session.withWorkbook(wb, path, readonly)
      yield Format.openSuccess(path.toString, wb.sheets.map(_.name.value))

    case Command.Create(sheetNames) =>
      for
        _ <- requireNoWorkbook
        sheets = sheetNames.map(name => Sheet(SheetName.apply(name).toOption.get))
        wb = Workbook(sheets)
        _ = session = Session.newWorkbook(wb)
      yield Format.createSuccess(sheetNames)

    case Command.Close(discard) =>
      for
        _ <- requireWorkbook
        _ <-
          if session.isDirty && !discard
          then IO.raiseError(new Exception("Unsaved changes. Use --discard to force close."))
          else IO.unit
        path = session.path.map(_.toString).getOrElse("(unsaved)")
        _ = session = Session.empty
      yield s"Closed: $path"

    case Command.Sheets =>
      for
        wb <- requireWorkbook
        sheetStats = wb.sheets.map { sheet =>
          val usedRange = sheet.usedRange
          val cellCount = sheet.cells.size
          val formulaCount = sheet.cells.values.count(_.isFormula)
          (sheet.name.value, usedRange, cellCount, formulaCount)
        }
      yield Markdown.renderSheetList(sheetStats)

    case Command.Select(name) =>
      for
        wb <- requireWorkbook
        sheetName <- IO.fromEither(SheetName.apply(name).left.map(e => new Exception(e)))
        sheet <- IO.fromOption(wb.sheets.find(_.name == sheetName))(
          new Exception(s"Sheet not found: $name")
        )
        _ = session = session.selectSheet(sheetName)
        usedRange = sheet.usedRange.map(_.toA1)
        cellCount = sheet.cells.size
        formulaCount = sheet.cells.values.count(_.isFormula)
      yield Format.selectSuccess(name, usedRange, cellCount, formulaCount)

    case Command.Bounds(sheetOpt) =>
      for
        wb <- requireWorkbook
        sheet <- getSheet(sheetOpt)
        name = sheet.name.value
        usedRange = sheet.usedRange
        cellCount = sheet.cells.size
      yield usedRange match
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

    case Command.View(rangeStr, showFormulas, limit) =>
      for
        sheet <- getCurrentSheet
        range <- IO.fromEither(CellRange.parse(rangeStr).left.map(e => new Exception(e.toString)))
        limitedRange = limitRange(range, limit)
      yield Markdown.renderRange(sheet, limitedRange, showFormulas)

    case Command.Cell(refStr) =>
      for
        sheet <- getCurrentSheet
        ref <- IO.fromEither(ARef.parse(refStr).left.map(e => new Exception(e.toString)))
        cell = sheet.cells.get(ref)
        value = cell.map(_.value).getOrElse(CellValue.Empty)
        formatted = formatCellValue(value)
        // TODO: Add dependency analysis when integrated with evaluator
        deps = Vector.empty[ARef]
        dependents = Vector.empty[ARef]
      yield Format.cellInfo(ref, value, formatted, deps, dependents)

    case Command.Search(pattern, limit) =>
      for
        sheet <- getCurrentSheet
        regex = pattern.r
        results = sheet.cells.toVector
          .filter { case (_, cell) =>
            val text = formatCellValue(cell.value)
            regex.findFirstIn(text).isDefined
          }
          .take(limit)
          .map { case (ref, cell) =>
            val value = formatCellValue(cell.value)
            // Simple context: just the value for now
            (ref, value, value)
          }
      yield Markdown.renderSearchResults(results)

    case Command.Eval(formulaStr, overrides) =>
      for
        sheet <- getCurrentSheet
        // Apply temporary overrides to create a hypothetical sheet
        tempSheet <- applyOverrides(sheet, overrides)
        // Normalize formula (add = if missing)
        formula = if formulaStr.startsWith("=") then formulaStr else s"=$formulaStr"
        // Evaluate against the (possibly modified) sheet
        result <- IO.fromEither(
          SheetEvaluator.evaluateFormula(tempSheet)(formula).left.map(e => new Exception(e.message))
        )
      yield Format.evalSuccess(formula, result, overrides)

    case Command.Put(refStr, valueStr) =>
      for
        sheet <- getCurrentSheet
        ref <- IO.fromEither(ARef.parse(refStr).left.map(e => new Exception(e.toString)))
        value = parseValue(valueStr)
        updatedSheet = sheet.put(ref, value)
        _ = session = session.updateSheet(updatedSheet)
      yield Format.putSuccess(ref, value)

    case Command.PutFormula(refStr, formulaStr) =>
      for
        sheet <- getCurrentSheet
        ref <- IO.fromEither(ARef.parse(refStr).left.map(e => new Exception(e.toString)))
        formula = if formulaStr.startsWith("=") then formulaStr.drop(1) else formulaStr
        value = CellValue.Formula(formula)
        updatedSheet = sheet.put(ref, value)
        _ = session = session.updateSheet(updatedSheet)
      yield Format.putSuccess(ref, value)

    case Command.Save =>
      for
        wb <- requireWorkbook
        path <- IO.fromOption(session.path)(new Exception("No file path. Use 'saveas' instead."))
        _ <- ExcelIO.instance[IO].write(wb, path)
        _ = session = session.markClean
        cellCount = wb.sheets.map(_.cells.size).sum
      yield Format.saveSuccess(path.toString, wb.sheets.size, cellCount)

    case Command.SaveAs(path) =>
      for
        wb <- requireWorkbook
        _ <- ExcelIO.instance[IO].write(wb, path)
        _ = session = session.copy(path = Some(path)).markClean
        cellCount = wb.sheets.map(_.cells.size).sum
      yield Format.saveSuccess(path.toString, wb.sheets.size, cellCount)

  // ==========================================================================
  // Helpers
  // ==========================================================================

  private def requireWorkbook: IO[Workbook] =
    IO.fromOption(session.workbook)(new Exception("No workbook open. Use 'xl open <path>' first."))

  private def requireNoWorkbook: IO[Unit] =
    if session.isOpen then
      IO.raiseError(new Exception("A workbook is already open. Use 'xl close' first."))
    else IO.unit

  private def getCurrentSheet: IO[Sheet] =
    IO.fromOption(session.currentSheet)(new Exception("No sheet available."))

  private def getSheet(nameOpt: Option[String]): IO[Sheet] =
    nameOpt match
      case Some(name) =>
        for
          wb <- requireWorkbook
          sheetName <- IO.fromEither(SheetName.apply(name).left.map(e => new Exception(e)))
          sheet <- IO.fromOption(wb.sheets.find(_.name == sheetName))(
            new Exception(s"Sheet not found: $name")
          )
        yield sheet
      case None =>
        getCurrentSheet

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
    // Try parsing as number
    scala.util.Try(BigDecimal(s)).toOption.map(CellValue.Number.apply).getOrElse {
      // Try parsing as boolean
      s.toLowerCase match
        case "true" => CellValue.Bool(true)
        case "false" => CellValue.Bool(false)
        case _ =>
          // Default to text (strip quotes if present)
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
  // Open/Initialize
  case Open(path: Path, readonly: Boolean)
  case Create(sheets: Vector[String])
  case Close(discard: Boolean)
  // Navigate
  case Sheets
  case Select(name: String)
  case Bounds(sheet: Option[String])
  // Explore
  case View(range: String, showFormulas: Boolean, limit: Int)
  case Cell(ref: String)
  case Search(pattern: String, limit: Int)
  // Analyze
  case Eval(formula: String, overrides: List[String])
  // Mutate
  case Put(ref: String, value: String)
  case PutFormula(ref: String, formula: String)
  // Persist
  case Save
  case SaveAs(path: Path)
