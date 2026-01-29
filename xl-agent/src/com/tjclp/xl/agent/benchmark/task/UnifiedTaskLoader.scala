package com.tjclp.xl.agent.benchmark.task

import cats.effect.IO
import cats.syntax.all.*
import com.tjclp.xl.agent.benchmark.common.SampleGenerator
import com.tjclp.xl.agent.benchmark.grading.GraderType
import com.tjclp.xl.agent.benchmark.token.{TokenTask, TokenTasks}
import com.tjclp.xl.agent.benchmark.{
  BenchConfig,
  BenchTask,
  DatasetTask,
  TaskLoader as LegacyTaskLoader
}
import com.tjclp.xl.agent.error.AgentError
import io.circe.*
import io.circe.parser.*

import java.nio.file.{Files, Path}
import scala.jdk.CollectionConverters.*

// ============================================================================
// Unified Task Loader
// ============================================================================

/** Unified task loader that supports multiple benchmark formats */
object UnifiedTaskLoader:

  /** Benchmark source types */
  enum BenchmarkSource:
    case SpreadsheetBench(dataDir: Path, dataset: String)
    case TokenBenchmark(samplePath: Option[Path])
    case CustomFile(path: Path)

  /** Task filter configuration */
  case class TaskFilter(
    taskIds: Option[List[String]] = None,
    skipIds: Set[String] = Set.empty,
    category: Option[String] = None,
    limit: Option[Int] = None,
    includeVba: Boolean = false
  ):
    def matches(task: BenchmarkTask): Boolean =
      val idMatches = taskIds.map(_.contains(task.taskIdValue)).getOrElse(true)
      val notSkipped = !skipIds.contains(task.taskIdValue)
      val categoryMatches =
        category.map(c => TaskCategory.fromString(c) == task.category).getOrElse(true)
      val vbaOk = includeVba || task.category != TaskCategory.VBA

      idMatches && notSkipped && categoryMatches && vbaOk

  object TaskFilter:
    val default: TaskFilter = TaskFilter()

  /** Load tasks from any supported source */
  def load(
    source: BenchmarkSource,
    filter: TaskFilter = TaskFilter.default
  ): IO[List[BenchmarkTask]] =
    source match
      case BenchmarkSource.SpreadsheetBench(dataDir, dataset) =>
        loadSpreadsheetBench(dataDir, dataset, filter)
      case BenchmarkSource.TokenBenchmark(samplePath) =>
        loadTokenBenchmark(samplePath, filter)
      case BenchmarkSource.CustomFile(path) =>
        loadCustomFile(path, filter)

  // --------------------------------------------------------------------------
  // SpreadsheetBench Loading
  // --------------------------------------------------------------------------

  /** Load tasks from SpreadsheetBench dataset */
  def loadSpreadsheetBench(
    dataDir: Path,
    dataset: String,
    filter: TaskFilter
  ): IO[List[BenchmarkTask]] =
    // Use the legacy TaskLoader to get BenchTasks, then convert
    // Create a BenchConfig to pass to the legacy loader
    val legacyConfig = BenchConfig(
      dataset = dataset,
      dataDir = dataDir,
      taskIds = filter.taskIds,
      skipIds = filter.skipIds,
      category = filter.category,
      limit = filter.limit
    )
    LegacyTaskLoader.load(legacyConfig).map { benchTasks =>
      benchTasks
        .map(convertFromBenchTask)
        .filter(filter.matches)
    }

  /** Convert legacy BenchTask to unified BenchmarkTask */
  private def convertFromBenchTask(bt: BenchTask): BenchmarkTask =
    BenchmarkTask(
      id = TaskId(bt.id),
      instruction = bt.instruction,
      category = TaskCategory.fromInstructionType(bt.instructionType),
      inputSource = InputSource.TestCases(
        bt.testCases.map { tc =>
          TestCaseFile(tc.caseNum, tc.inputPath, tc.answerPath)
        }.toVector
      ),
      evaluation = EvaluationSpec.forFile(bt.answerPosition),
      metadata = TaskMetadata(source = Some("spreadsheetbench"))
    )

  // --------------------------------------------------------------------------
  // TokenBenchmark Loading
  // --------------------------------------------------------------------------

  /** Load tasks from TokenBenchmark definitions */
  def loadTokenBenchmark(
    samplePath: Option[Path],
    filter: TaskFilter
  ): IO[List[BenchmarkTask]] =
    for
      // Use provided path or ensure default sample exists (creating it if needed)
      effectivePath <- samplePath match
        case Some(p) =>
          // User provided explicit path - verify it exists
          IO.blocking(Files.exists(p)).flatMap {
            case true => IO.pure(p)
            case false =>
              IO.raiseError(
                AgentError.ConfigError(s"Sample file not found: $p")
              )
          }
        case None =>
          // Use default path, creating sample if needed
          SampleGenerator.ensureSampleExists()
    yield TokenTasks
      .allTasks(includeLarge = false)
      .map(convertFromTokenTask(_, Some(effectivePath)))
      .filter(filter.matches)
      .take(filter.limit.getOrElse(Int.MaxValue))

  /** Convert TokenTask to unified BenchmarkTask */
  private def convertFromTokenTask(tt: TokenTask, samplePath: Option[Path]): BenchmarkTask =
    BenchmarkTask(
      id = TaskId(tt.id),
      instruction = tt.xlPrompt, // Use xl prompt as the canonical instruction
      category = TaskCategory.Analysis,
      inputSource = samplePath match
        case Some(p) => InputSource.SingleFile(p)
        case None => InputSource.NoInput,
      evaluation = tt.expectedAnswer match
        case Some(expected) => EvaluationSpec.forLLM(expected)
        case None => EvaluationSpec(GraderType.LLM),
      metadata = TaskMetadata(
        source = Some("token_benchmark"),
        custom = Map(
          "name" -> tt.name,
          "description" -> tt.description,
          "xlsxPrompt" -> tt.xlsxPrompt
        )
      )
    )

  // --------------------------------------------------------------------------
  // Custom File Loading
  // --------------------------------------------------------------------------

  /** Load tasks from a custom JSON file */
  def loadCustomFile(path: Path, filter: TaskFilter): IO[List[BenchmarkTask]] =
    for
      content <- IO.blocking(Files.readString(path))
      tasks <- IO.fromEither(
        decode[List[BenchmarkTask]](content)
          .leftMap(e => AgentError.ParseError(content.take(200), e.getMessage))
      )
    yield tasks.filter(filter.matches).take(filter.limit.getOrElse(Int.MaxValue))

  // --------------------------------------------------------------------------
  // Utility Methods
  // --------------------------------------------------------------------------

  /** Get test case files for a task */
  def getTestCases(task: BenchmarkTask): List[TestCaseFile] =
    task.inputSource match
      case InputSource.TestCases(cases) => cases.toList
      case _ => Nil

  /** Get the single input file for a task */
  def getSingleInputFile(task: BenchmarkTask): Option[Path] =
    task.inputSource match
      case InputSource.SingleFile(path) => Some(path)
      case InputSource.TestCases(cases) => cases.headOption.map(_.inputPath)
      case _ => None

  /** Check if task needs file-based evaluation */
  def needsFileGrading(task: BenchmarkTask): Boolean =
    task.evaluation.requiresFile

  /** Check if task needs LLM-based evaluation */
  def needsLLMGrading(task: BenchmarkTask): Boolean =
    task.evaluation.requiresLLM

// ============================================================================
// Conversion Utilities
// ============================================================================

object TaskConversions:
  import com.tjclp.xl.agent.benchmark.{BenchTask, TestCase}

  /** Convert unified task back to legacy BenchTask (for backwards compat) */
  def toBenchTask(task: BenchmarkTask): Option[BenchTask] =
    task.inputSource match
      case InputSource.TestCases(cases) =>
        Some(
          BenchTask(
            id = task.taskIdValue,
            instruction = task.instruction,
            instructionType = task.category match
              case TaskCategory.CellLevel => "cell_level"
              case TaskCategory.SheetLevel => "sheet_level"
              case TaskCategory.WorkbookLevel => "workbook_level"
              case TaskCategory.Analysis => "analysis"
              case TaskCategory.VBA => "vba"
              case TaskCategory.Custom(name) => name,
            answerPosition = task.evaluation.answerPosition.getOrElse(""),
            testCases = cases.map(c => TestCase(c.caseNum, c.inputPath, c.answerPath)).toList
          )
        )
      case _ => None

  /** Convert unified task back to TokenTask (for backwards compat) */
  def toTokenTask(task: BenchmarkTask): TokenTask =
    TokenTask(
      id = task.taskIdValue,
      name = task.metadata.custom.getOrElse("name", task.taskIdValue),
      description = task.metadata.custom.getOrElse("description", task.instruction),
      xlPrompt = task.instruction,
      xlsxPrompt = task.metadata.custom.getOrElse("xlsxPrompt", task.instruction),
      expectedAnswer = task.evaluation.expectedAnswer
    )
