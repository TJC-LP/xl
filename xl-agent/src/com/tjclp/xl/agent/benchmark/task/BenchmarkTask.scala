package com.tjclp.xl.agent.benchmark.task

import com.tjclp.xl.agent.benchmark.grading.GraderType
import io.circe.*
import io.circe.generic.semiauto.*

import java.nio.file.Path

// ============================================================================
// Unified Benchmark Task Model
// ============================================================================

/** A unified benchmark task that can be used by any grader */
case class BenchmarkTask(
  id: TaskId,
  instruction: String,
  category: TaskCategory,
  inputSource: InputSource,
  evaluation: EvaluationSpec,
  metadata: TaskMetadata = TaskMetadata.empty
):
  def taskIdValue: String = id.value
  def isVba: Boolean =
    instruction.toLowerCase.contains("vba") ||
      instruction.toLowerCase.contains("macro")

object BenchmarkTask:
  given Encoder[BenchmarkTask] = deriveEncoder
  given Decoder[BenchmarkTask] = deriveDecoder

// ============================================================================
// Task ID Opaque Type
// ============================================================================

object TaskId:
  opaque type TaskId = String

  inline def apply(s: String): TaskId = s

  def fromInt(i: Int): TaskId = i.toString
  def fromLong(l: Long): TaskId = l.toString
  def fromJson(json: Json): TaskId = json.asString
    .orElse(json.asNumber.map(_.toInt.map(_.toString).getOrElse(json.noSpaces)))
    .getOrElse(json.noSpaces)

  extension (id: TaskId)
    inline def value: String = id
    inline def toInt: Option[Int] = id.toIntOption

  given Encoder[TaskId] = Encoder.encodeString.contramap(_.value)
  given Decoder[TaskId] = Decoder.decodeString.map(apply)
  given CanEqual[TaskId, TaskId] = CanEqual.derived

export TaskId.TaskId

// ============================================================================
// Task Category
// ============================================================================

/** Categorization of benchmark tasks */
enum TaskCategory derives CanEqual:
  case CellLevel
  case SheetLevel
  case WorkbookLevel
  case Analysis
  case VBA
  case Custom(name: String)

object TaskCategory:
  def fromString(s: String): TaskCategory =
    s.toLowerCase.replaceAll("[_-]", "") match
      case "celllevel" | "cell" => CellLevel
      case "sheetlevel" | "sheet" => SheetLevel
      case "workbooklevel" | "workbook" => WorkbookLevel
      case "analysis" => Analysis
      case "vba" | "macro" => VBA
      case other => Custom(other)

  def fromInstructionType(instructionType: String): TaskCategory =
    fromString(instructionType)

  given Encoder[TaskCategory] = Encoder.encodeString.contramap {
    case CellLevel => "cell_level"
    case SheetLevel => "sheet_level"
    case WorkbookLevel => "workbook_level"
    case Analysis => "analysis"
    case VBA => "vba"
    case Custom(name) => name
  }

  given Decoder[TaskCategory] = Decoder.decodeString.map(fromString)

// ============================================================================
// Input Source
// ============================================================================

/** Where task input files come from */
enum InputSource:
  /** Multiple test cases with input/answer pairs (SpreadsheetBench) */
  case TestCases(cases: Vector[TestCaseFile])

  /** Single input file (TokenBenchmark) */
  case SingleFile(path: Path)

  /** No input file required */
  case NoInput

  /** Path to data directory (for dynamic file resolution) */
  case DataDirectory(dir: Path, pattern: String)

object InputSource:
  given Encoder[InputSource] = Encoder.instance {
    case TestCases(cases) =>
      Json.obj(
        "type" -> Json.fromString("test_cases"),
        "cases" -> Encoder[Vector[TestCaseFile]].apply(cases)
      )
    case SingleFile(path) =>
      Json.obj("type" -> Json.fromString("single_file"), "path" -> Json.fromString(path.toString))
    case NoInput =>
      Json.obj("type" -> Json.fromString("no_input"))
    case DataDirectory(dir, pattern) =>
      Json.obj(
        "type" -> Json.fromString("data_directory"),
        "dir" -> Json.fromString(dir.toString),
        "pattern" -> Json.fromString(pattern)
      )
  }

  given Decoder[InputSource] = Decoder.instance { c =>
    c.get[String]("type").flatMap {
      case "test_cases" => c.get[Vector[TestCaseFile]]("cases").map(TestCases.apply)
      case "single_file" => c.get[String]("path").map(p => SingleFile(Path.of(p)))
      case "no_input" => Right(NoInput)
      case "data_directory" =>
        for
          dir <- c.get[String]("dir")
          pattern <- c.get[String]("pattern")
        yield DataDirectory(Path.of(dir), pattern)
      case other => Left(DecodingFailure(s"Unknown input source type: $other", c.history))
    }
  }

/** A single test case file pair */
case class TestCaseFile(
  caseNum: Int,
  inputPath: Path,
  answerPath: Path
)

object TestCaseFile:
  given Encoder[TestCaseFile] = Encoder.instance { tcf =>
    Json.obj(
      "caseNum" -> Json.fromInt(tcf.caseNum),
      "inputPath" -> Json.fromString(tcf.inputPath.toString),
      "answerPath" -> Json.fromString(tcf.answerPath.toString)
    )
  }

  given Decoder[TestCaseFile] = Decoder.instance { c =>
    for
      caseNum <- c.get[Int]("caseNum")
      inputPath <- c.get[String]("inputPath")
      answerPath <- c.get[String]("answerPath")
    yield TestCaseFile(caseNum, Path.of(inputPath), Path.of(answerPath))
  }

// ============================================================================
// Evaluation Specification
// ============================================================================

/** How a task should be evaluated */
case class EvaluationSpec(
  graderType: GraderType,
  answerPosition: Option[String] = None,
  expectedAnswer: Option[String] = None,
  timeout: Option[Long] = None // Timeout in milliseconds
):
  def requiresFile: Boolean = graderType match
    case GraderType.File | GraderType.FileAndLLM => true
    case _ => false

  def requiresLLM: Boolean = graderType match
    case GraderType.LLM | GraderType.FileAndLLM => true
    case _ => false

object EvaluationSpec:
  val fileOnly: EvaluationSpec = EvaluationSpec(GraderType.File)
  val llmOnly: EvaluationSpec = EvaluationSpec(GraderType.LLM)
  val both: EvaluationSpec = EvaluationSpec(GraderType.FileAndLLM)

  def forFile(answerPosition: String): EvaluationSpec =
    EvaluationSpec(GraderType.File, answerPosition = Some(answerPosition))

  def forLLM(expectedAnswer: String): EvaluationSpec =
    EvaluationSpec(GraderType.LLM, expectedAnswer = Some(expectedAnswer))

  def forBoth(answerPosition: String, expectedAnswer: String): EvaluationSpec =
    EvaluationSpec(
      GraderType.FileAndLLM,
      answerPosition = Some(answerPosition),
      expectedAnswer = Some(expectedAnswer)
    )

  given Encoder[EvaluationSpec] = deriveEncoder
  given Decoder[EvaluationSpec] = deriveDecoder

// ============================================================================
// Task Metadata
// ============================================================================

/** Additional metadata for a task */
case class TaskMetadata(
  source: Option[String] = None,
  difficulty: Option[String] = None,
  tags: Set[String] = Set.empty,
  notes: Option[String] = None,
  custom: Map[String, String] = Map.empty
):
  def withTag(tag: String): TaskMetadata = copy(tags = tags + tag)
  def withSource(s: String): TaskMetadata = copy(source = Some(s))
  def withDifficulty(d: String): TaskMetadata = copy(difficulty = Some(d))

object TaskMetadata:
  val empty: TaskMetadata = TaskMetadata()

  def apply(source: String): TaskMetadata = TaskMetadata(source = Some(source))

  given Encoder[TaskMetadata] = deriveEncoder
  given Decoder[TaskMetadata] = deriveDecoder
