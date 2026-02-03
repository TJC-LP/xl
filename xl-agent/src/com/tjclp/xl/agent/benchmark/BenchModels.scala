package com.tjclp.xl.agent.benchmark

import io.circe.*
import io.circe.generic.semiauto.*
import java.nio.file.Path
import com.tjclp.xl.agent.TokenUsage

// ============================================================================
// Dataset JSON Models
// ============================================================================

/** Raw task from dataset.json. Note: id can be Int or String */
case class DatasetTask(
  id: Json,
  instruction: String,
  spreadsheet_path: String,
  instruction_type: String,
  answer_position: String
)

object DatasetTask:
  given Decoder[DatasetTask] = deriveDecoder

// ============================================================================
// Benchmark Task Models
// ============================================================================

/** A single test case (each task has 3) */
case class TestCase(
  caseNum: Int,
  inputPath: Path,
  answerPath: Path
)

/** A benchmark task with all its test cases */
case class BenchTask(
  id: String,
  instruction: String,
  instructionType: String,
  answerPosition: String,
  testCases: List[TestCase]
):
  def isVba: Boolean =
    instruction.toLowerCase.contains("vba") ||
      instruction.toLowerCase.contains("macro")

// ============================================================================
// Evaluation Models
// ============================================================================

/** Normalized value for comparison */
enum ComparableValue:
  case Empty
  case Number(value: BigDecimal)
  case Text(value: String)
  case Bool(value: Boolean)
  case Error(value: String)

object ComparableValue:
  given Encoder[ComparableValue] = Encoder.instance {
    case Empty => Json.fromString("empty")
    case Number(v) => Json.fromString(v.toString)
    case Text(v) => Json.fromString(s"\"$v\"")
    case Bool(v) => Json.fromString(v.toString)
    case Error(v) => Json.fromString(s"#$v")
  }

/** A cell that didn't match */
case class CellMismatch(
  ref: String,
  expected: ComparableValue,
  actual: ComparableValue
)

object CellMismatch:
  given Encoder[CellMismatch] = deriveEncoder

/** Result of comparing a single range */
case class RangeResult(
  position: String,
  passed: Boolean,
  mismatches: List[CellMismatch]
)

object RangeResult:
  given Encoder[RangeResult] = deriveEncoder

/** Result of evaluating a single test case */
case class TestCaseResult(
  caseNum: Int,
  passed: Boolean,
  rangeResults: List[RangeResult],
  error: Option[String] = None
)

object TestCaseResult:
  given Encoder[TestCaseResult] = deriveEncoder

/** Score for a task (per SpreadsheetBench spec) */
case class TaskScore(
  softRestriction: Double, // passing_cases / 3 (0.0 - 1.0)
  hardRestriction: Int // 1 if all_pass else 0
)

object TaskScore:
  given Encoder[TaskScore] = deriveEncoder

  def fromResults(results: List[TestCaseResult]): TaskScore =
    val passingCount = results.count(_.passed)
    val total = results.length.max(1)
    TaskScore(
      softRestriction = passingCount.toDouble / total,
      hardRestriction = if passingCount == total then 1 else 0
    )

// ============================================================================
// Task Result Models
// ============================================================================

/** Result of running a single benchmark task */
case class TaskResult(
  taskId: String,
  instruction: String,
  instructionType: String,
  score: TaskScore,
  testCaseResults: List[TestCaseResult],
  usage: TokenUsage,
  latencyMs: Long,
  error: Option[String] = None,
  skipped: Boolean = false,
  skipReason: Option[String] = None,
  responseText: Option[String] = None
)

object TaskResult:
  given Encoder[TaskResult] = deriveEncoder

  def skipped(task: BenchTask, reason: String): TaskResult =
    TaskResult(
      taskId = task.id,
      instruction = task.instruction,
      instructionType = task.instructionType,
      score = TaskScore(0.0, 0),
      testCaseResults = Nil,
      usage = TokenUsage.zero,
      latencyMs = 0,
      skipped = true,
      skipReason = Some(reason)
    )

// ============================================================================
// Benchmark Report Models
// ============================================================================

/** Aggregate statistics */
case class AggregateStats(
  totalTasks: Int,
  completedTasks: Int,
  skippedTasks: Int,
  avgSoftRestriction: Double,
  avgHardRestriction: Double,
  totalInputTokens: Long,
  totalOutputTokens: Long,
  avgLatencyMs: Long
)

object AggregateStats:
  given Encoder[AggregateStats] = deriveEncoder

  def fromResults(results: List[TaskResult]): AggregateStats =
    val completed = results.filterNot(_.skipped)
    val skipped = results.filter(_.skipped)
    AggregateStats(
      totalTasks = results.length,
      completedTasks = completed.length,
      skippedTasks = skipped.length,
      avgSoftRestriction =
        if completed.isEmpty then 0.0
        else completed.map(_.score.softRestriction).sum / completed.length,
      avgHardRestriction =
        if completed.isEmpty then 0.0
        else completed.map(_.score.hardRestriction).sum.toDouble / completed.length,
      totalInputTokens = completed.map(_.usage.inputTokens).sum,
      totalOutputTokens = completed.map(_.usage.outputTokens).sum,
      avgLatencyMs =
        if completed.isEmpty then 0
        else completed.map(_.latencyMs).sum / completed.length
    )

/** Full benchmark report */
case class BenchmarkReport(
  timestamp: String,
  dataset: String,
  model: String,
  approach: String,
  aggregate: AggregateStats,
  results: List[TaskResult]
):
  def toJsonPretty: String =
    import io.circe.syntax.*
    this.asJson.spaces2

object BenchmarkReport:
  given Encoder[BenchmarkReport] = deriveEncoder

// ============================================================================
// Approach Configuration
// ============================================================================

/** Benchmark approach for spreadsheet operations */
enum Approach derives CanEqual:
  case Xl // Custom xl-cli skill with uploaded binary
  case Xlsx // Built-in Anthropic xlsx skill (openpyxl)

object Approach:
  given Encoder[Approach] = Encoder.encodeString.contramap(_.toString.toLowerCase)

// ============================================================================
// CLI Config
// ============================================================================

case class BenchConfig(
  dataset: String = "sample_data_200",
  dataDir: Path =
    Path.of(Option(java.lang.System.getenv("HOME")).getOrElse("."), "git/SpreadsheetBench/data"),
  approach: Approach = Approach.Xl,
  limit: Option[Int] = None,
  taskIds: Option[List[String]] = None,
  skipIds: Set[String] = Set.empty,
  category: Option[String] = None,
  includeVba: Boolean = false,
  parallelism: Int = 4,
  verbose: Boolean = false,
  outputDir: Path = Path.of("results"),
  xlBinary: Option[Path] = None,
  xlSkill: Option[Path] = None,
  // Comparison mode options
  compare: Boolean = false,
  xlOnly: Boolean = false,
  xlsxOnly: Boolean = false,
  // N-skill comparison (new)
  skills: Option[List[String]] = None,
  stream: Boolean = false,
  listSkills: Boolean = false
)
