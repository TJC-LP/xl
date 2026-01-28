package spreadsheetbench

import zio.*
import zio.json.*
import java.nio.file.Path
import scala.sys.process.*
import scala.math.BigDecimal.RoundingMode

object Evaluator:

  /** Compare model output against expected answer */
  def compare(
      outputPath: Path,
      answerPath: Path,
      answerPosition: String,
      xlPath: String = "xl"
  ): Task[List[RangeResult]] =
    for
      positions <- ZIO.attempt(parsePositions(answerPosition))
      results <- ZIO.foreach(positions) { case (sheetOpt, range) =>
        compareRange(outputPath, answerPath, sheetOpt, range, xlPath)
      }
    yield results

  /** Compare a single range between output and answer workbooks */
  private def compareRange(
      outputPath: Path,
      answerPath: Path,
      sheetOpt: Option[String],
      range: String,
      xlPath: String
  ): Task[RangeResult] =
    val position = sheetOpt.map(s => s"'$s'!$range").getOrElse(range)
    for
      outputCells <- getCellValues(outputPath, sheetOpt, range, xlPath)
      answerCells <- getCellValues(answerPath, sheetOpt, range, xlPath)
      mismatches   = findMismatches(outputCells, answerCells)
    yield RangeResult(
      position = position,
      passed = mismatches.isEmpty,
      mismatches = mismatches
    )

  /** Get cell values from a workbook using xl CLI */
  private def getCellValues(
      path: Path,
      sheetOpt: Option[String],
      range: String,
      xlPath: String
  ): Task[Map[String, ComparableValue]] =
    for
      // If no sheet specified, detect the first sheet in the workbook
      sheet <- sheetOpt match
        case Some(s) => ZIO.succeed(s)
        case None => detectFirstSheet(path, xlPath)

      sheetArgs = List("--sheet", sheet)
      cmd = List(xlPath, "-f", path.toString) ++ sheetArgs ++ List("view", range, "--format", "json", "--eval")

      result <- ZIO.attemptBlocking {
        val output = cmd.!!
        output
      }.mapError(e => new Exception(s"xl CLI failed: ${e.getMessage}"))
      json   <- ZIO.fromEither(result.fromJson[ViewOutput])
                  .mapError(e => new Exception(s"JSON parse error: $e"))
    yield json.rows.flatMap(_.cells).map(c => c.ref -> normalizeCell(c)).toMap

  /** Detect the first sheet name in a workbook */
  private def detectFirstSheet(path: Path, xlPath: String): Task[String] =
    ZIO.attemptBlocking {
      val cmd = List(xlPath, "-f", path.toString, "sheets")
      val output = cmd.!!
      // Parse markdown table output:
      // | #   | Name   | Dimension | State |
      // |-----|--------|-----------|-------|
      // | 1   | Sheet1 | D2:H5     |       |
      //
      // Skip header rows (first 2 lines), then extract Name column from first data row
      val lines = output.linesIterator.toList
      val dataLines = lines.drop(2) // Skip header and separator
      dataLines.headOption.flatMap { line =>
        val cols = line.split("\\|").map(_.trim).filter(_.nonEmpty)
        cols.lift(1) // Name is the second column (after #)
      }.getOrElse("Sheet1")
    }.mapError(e => new Exception(s"Failed to detect sheet: ${e.getMessage}"))

  /** Normalize a cell value for comparison */
  private def normalizeCell(cell: CellJson): ComparableValue =
    cell.`type` match
      case "empty" => ComparableValue.Empty

      case "number" =>
        cell.value match
          case Some(n: Double) =>
            ComparableValue.Number(BigDecimal(n).setScale(2, RoundingMode.HALF_UP))
          case _ => ComparableValue.Empty

      case "text" =>
        cell.value match
          case Some(s: String) =>
            tryParseNumber(s).getOrElse(ComparableValue.Text(s.trim))
          case _ => ComparableValue.Empty

      case "formula" =>
        // Use the cached/formatted value for formulas
        cell.formatted match
          case Some(f) if f.startsWith("=") =>
            // Formula wasn't evaluated - this is an error condition
            // Return the formula text as Text so it can be compared/logged
            ComparableValue.Text(f)
          case Some(f) =>
            // Have a cached value - try to parse as number or text
            tryParseFormatted(f).getOrElse(ComparableValue.Text(f.trim))
          case None =>
            ComparableValue.Empty

      case "boolean" =>
        cell.value match
          case Some(b: Boolean) => ComparableValue.Bool(b)
          case _                => ComparableValue.Empty

      case "error" =>
        cell.value match
          case Some(e: String) => ComparableValue.Error(e)
          case _               => ComparableValue.Error("UNKNOWN")

      case _ => ComparableValue.Empty

  /** Try to parse a formatted string (handles currency, percentages, etc.) */
  private def tryParseFormatted(s: String): Option[ComparableValue] =
    tryParseNumber(s)

  /** Try to parse a string as a number, stripping currency/percent symbols */
  private def tryParseNumber(s: String): Option[ComparableValue.Number] =
    val cleaned = s.trim
      .replaceAll("[\\$€£¥,]", "")
      .replaceAll("%$", "")
      .replaceAll("^\\((.+)\\)$", "-$1") // Handle accounting negative format
      .trim

    if cleaned.isEmpty then None
    else
      scala.util.Try(BigDecimal(cleaned).setScale(2, RoundingMode.HALF_UP))
        .toOption
        .map(ComparableValue.Number(_))

  /** Compare two maps of cell values and find mismatches */
  private def findMismatches(
      output: Map[String, ComparableValue],
      answer: Map[String, ComparableValue]
  ): List[CellMismatch] =
    // Compare all cells in the answer
    answer.toList.flatMap { case (ref, expected) =>
      val actual = output.getOrElse(ref, ComparableValue.Empty)
      if !valuesEqual(expected, actual) then
        Some(CellMismatch(ref, expected, actual))
      else
        None
    }

  /** Compare two values for equality using SpreadsheetBench rules */
  private def valuesEqual(a: ComparableValue, b: ComparableValue): Boolean =
    (a, b) match
      case (ComparableValue.Empty, ComparableValue.Empty) => true
      case (ComparableValue.Number(x), ComparableValue.Number(y)) => x == y
      case (ComparableValue.Text(x), ComparableValue.Text(y)) => x.equalsIgnoreCase(y)
      case (ComparableValue.Bool(x), ComparableValue.Bool(y)) => x == y
      case (ComparableValue.Error(x), ComparableValue.Error(y)) => x == y
      // Empty cells can match empty strings
      case (ComparableValue.Empty, ComparableValue.Text(s)) if s.trim.isEmpty => true
      case (ComparableValue.Text(s), ComparableValue.Empty) if s.trim.isEmpty => true
      case _ => false

  /** Parse answer_position into list of (sheet, range) pairs */
  private def parsePositions(answerPosition: String): List[(Option[String], String)] =
    answerPosition.split(",").toList.map(_.trim).map { pos =>
      // Handle formats like: 'Sheet1'!A1:B10, Sheet1!A1:B10, A1:B10
      val sheetRangePattern = """'?([^'!]+)'?!(.+)""".r
      pos match
        case sheetRangePattern(sheet, range) => (Some(sheet), range)
        case range                           => (None, range)
    }

// ============================================================================
// JSON Models for xl CLI output
// ============================================================================

case class CellJson(
    ref: String,
    `type`: String,
    value: Option[Any] = None,
    formatted: Option[String] = None
)

object CellJson:
  given JsonDecoder[CellJson] = JsonDecoder[zio.json.ast.Json].map { json =>
    val obj = json.asInstanceOf[zio.json.ast.Json.Obj]
    val fields = obj.fields.toMap

    val ref = fields.get("ref").flatMap(_.as[String].toOption).getOrElse("")
    val typ = fields.get("type").flatMap(_.as[String].toOption).getOrElse("empty")
    val formatted = fields.get("formatted").flatMap(_.as[String].toOption)

    val value: Option[Any] = typ match
      case "number"  => fields.get("value").flatMap(_.as[Double].toOption)
      case "text"    => fields.get("value").flatMap(_.as[String].toOption)
      case "boolean" => fields.get("value").flatMap(_.as[Boolean].toOption)
      case "error"   => fields.get("value").flatMap(_.as[String].toOption)
      case _         => None

    CellJson(ref, typ, value, formatted)
  }

case class RowJson(cells: List[CellJson])

object RowJson:
  given JsonDecoder[RowJson] = DeriveJsonDecoder.gen[RowJson]

case class ViewOutput(rows: List[RowJson])

object ViewOutput:
  given JsonDecoder[ViewOutput] = DeriveJsonDecoder.gen[ViewOutput]
