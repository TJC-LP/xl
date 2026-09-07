package com.tjclp.xl.agent.benchmark

import cats.effect.IO
import cats.syntax.all.*
import io.circe.*
import io.circe.generic.semiauto.*
import io.circe.parser.*
import com.tjclp.xl.agent.error.AgentError

import java.nio.file.Path
import scala.math.BigDecimal.RoundingMode
import scala.sys.process.*

object Evaluator:

  /** Compare model output against expected answer */
  def compare(
    outputPath: Path,
    answerPath: Path,
    answerPosition: String,
    xlPath: String = "xl"
  ): IO[List[RangeResult]] =
    for
      positions <- IO(parsePositions(answerPosition))
      results <- positions.traverse { case (sheetOpt, range) =>
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
  ): IO[RangeResult] =
    val position = sheetOpt.map(s => s"'$s'!$range").getOrElse(range)
    for
      outputCells <- getCellValues(outputPath, sheetOpt, range, xlPath)
      answerCells <- getCellValues(answerPath, sheetOpt, range, xlPath)
      mismatches = findMismatches(outputCells, answerCells)
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
  ): IO[Map[String, ComparableValue]] =
    for
      sheet <- sheetOpt match
        case Some(s) => IO.pure(s)
        case None => detectFirstSheet(path, xlPath)

      sheetArgs = List("--sheet", sheet)
      cmd = List(xlPath, "-f", path.toString) ++ sheetArgs ++ List(
        "view",
        range,
        "--format",
        "json",
        "--eval"
      )

      result <- IO
        .blocking(cmd.!!)
        .adaptError(e => AgentError.EvaluationFailed(s"xl CLI failed: ${e.getMessage}"))

      json <- IO.fromEither(
        decode[ViewOutput](result).leftMap(e =>
          AgentError.ParseError(result.take(200), e.getMessage)
        )
      )
    yield json.rows.flatMap(_.cells).map(c => c.ref -> normalizeCell(c)).toMap

  /**
   * The command the grader runs to learn a workbook's sheets: the `--json` envelope, never the text
   * table (GH-592).
   */
  def sheetsCommand(xlPath: String, path: Path): List[String] =
    List(xlPath, "-f", path.toString, "--json", "sheets")

  /**
   * The first sheet of a `sheets --json` envelope: `data[0].name`. A failed envelope (`ok: false`)
   * is an evaluation failure naming `error.code`; anything that is not the envelope — the text
   * table included — is a parse error. There is no default sheet: a guess would grade the wrong
   * range silently.
   */
  def firstSheet(envelope: String): Either[AgentError, String] =
    def parseError(cause: String): AgentError = AgentError.ParseError(envelope.take(200), cause)
    for
      json <- parse(envelope).leftMap(e => parseError(e.getMessage))
      cursor = json.hcursor
      ok <- cursor.get[Boolean]("ok").leftMap(e => parseError(e.getMessage))
      name <-
        if !ok then
          val error = cursor.downField("error")
          val code = error.get[String]("code").getOrElse("UNKNOWN")
          val message = error.get[String]("message").getOrElse("")
          Left(AgentError.EvaluationFailed(s"xl sheets failed [$code]: $message"))
        else
          cursor.get[Vector[Json]]("data").leftMap(e => parseError(e.getMessage)).flatMap {
            case first +: _ =>
              first.hcursor.get[String]("name").leftMap(e => parseError(e.getMessage))
            case _ => Left(AgentError.EvaluationFailed("xl sheets: the workbook has no sheets"))
          }
    yield name

  /** Detect the first sheet name in a workbook from `xl --json sheets` */
  private def detectFirstSheet(path: Path, xlPath: String): IO[String] =
    IO.blocking {
      val out = new StringBuilder
      val err = new StringBuilder
      val logger = ProcessLogger(
        line => { out.append(line).append('\n'); () },
        line => { err.append(line).append('\n'); () }
      )
      val exit = Process(sheetsCommand(xlPath, path)).!(logger)
      (exit, out.toString, err.toString)
    }.adaptError(e => AgentError.EvaluationFailed(s"xl CLI failed to start: ${e.getMessage}"))
      .flatMap { (exit, stdout, stderr) =>
        // With --json the envelope is on stdout whatever the exit code; an empty stdout means the
        // binary died before printing one, and stderr is all there is to report.
        if stdout.trim.isEmpty then
          IO.raiseError(
            AgentError.EvaluationFailed(
              s"xl sheets exited $exit without an envelope: ${stderr.trim}"
            )
          )
        else IO.fromEither(firstSheet(stdout))
      }

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
        cell.formatted match
          case Some(f) if f.startsWith("=") =>
            ComparableValue.Text(f)
          case Some(f) =>
            tryParseFormatted(f).getOrElse(ComparableValue.Text(f.trim))
          case None =>
            ComparableValue.Empty

      case "boolean" =>
        cell.value match
          case Some(b: Boolean) => ComparableValue.Bool(b)
          case _ => ComparableValue.Empty

      case "error" =>
        cell.value match
          case Some(e: String) => ComparableValue.Error(e)
          case _ => ComparableValue.Error("UNKNOWN")

      case _ => ComparableValue.Empty

  /** Try to parse a formatted string */
  private def tryParseFormatted(s: String): Option[ComparableValue] =
    tryParseNumber(s)

  /** Try to parse a string as a number */
  private def tryParseNumber(s: String): Option[ComparableValue.Number] =
    val cleaned = s.trim
      .replaceAll("[\\$\u20ac\u00a3\u00a5,]", "")
      .replaceAll("%$", "")
      .replaceAll("^\\((.+)\\)$", "-$1")
      .trim

    if cleaned.isEmpty then None
    else
      scala.util
        .Try(BigDecimal(cleaned).setScale(2, RoundingMode.HALF_UP))
        .toOption
        .map(ComparableValue.Number(_))

  /** Compare two maps of cell values and find mismatches */
  private def findMismatches(
    output: Map[String, ComparableValue],
    answer: Map[String, ComparableValue]
  ): List[CellMismatch] =
    answer.toList.flatMap { case (ref, expected) =>
      val actual = output.getOrElse(ref, ComparableValue.Empty)
      if !valuesEqual(expected, actual) then Some(CellMismatch(ref, expected, actual))
      else None
    }

  /** Compare two values for equality using SpreadsheetBench rules */
  private def valuesEqual(a: ComparableValue, b: ComparableValue): Boolean =
    (a, b) match
      case (ComparableValue.Empty, ComparableValue.Empty) => true
      case (ComparableValue.Number(x), ComparableValue.Number(y)) => x == y
      case (ComparableValue.Text(x), ComparableValue.Text(y)) => x.equalsIgnoreCase(y)
      case (ComparableValue.Bool(x), ComparableValue.Bool(y)) => x == y
      case (ComparableValue.Error(x), ComparableValue.Error(y)) => x == y
      case (ComparableValue.Empty, ComparableValue.Text(s)) if s.trim.isEmpty => true
      case (ComparableValue.Text(s), ComparableValue.Empty) if s.trim.isEmpty => true
      case _ => false

  /** Parse answer_position into list of (sheet, range) pairs */
  private def parsePositions(answerPosition: String): List[(Option[String], String)] =
    answerPosition.split(",").toList.map(_.trim).map { pos =>
      val sheetRangePattern = """'?([^'!]+)'?!(.+)""".r
      pos match
        case sheetRangePattern(sheet, range) => (Some(sheet), range)
        case range => (None, range)
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
  given Decoder[CellJson] = Decoder.instance { c =>
    for
      ref <- c.get[String]("ref")
      typ <- c.get[String]("type")
      formatted <- c.get[Option[String]]("formatted")
      value <- typ match
        case "number" => c.get[Option[Double]]("value").map(_.map(identity[Any]))
        case "text" => c.get[Option[String]]("value").map(_.map(identity[Any]))
        case "boolean" => c.get[Option[Boolean]]("value").map(_.map(identity[Any]))
        case "error" => c.get[Option[String]]("value").map(_.map(identity[Any]))
        case _ => Right(None)
    yield CellJson(ref, typ, value, formatted)
  }

case class RowJson(cells: List[CellJson])

object RowJson:
  given Decoder[RowJson] = deriveDecoder

case class ViewOutput(rows: List[RowJson])

object ViewOutput:
  given Decoder[ViewOutput] = deriveDecoder
