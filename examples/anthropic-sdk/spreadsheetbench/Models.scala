package spreadsheetbench

import zio.json.*
import java.nio.file.Path
import java.math.RoundingMode

// ============================================================================
// Dataset JSON Models
// ============================================================================

/** Raw task from dataset.json. Note: id can be Int or String */
case class DatasetTask(
    id: zio.json.ast.Json,
    instruction: String,
    spreadsheet_path: String,
    instruction_type: String,
    answer_position: String
)

object DatasetTask:
  given JsonDecoder[DatasetTask] = DeriveJsonDecoder.gen[DatasetTask]

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
  given JsonEncoder[ComparableValue] = JsonEncoder[String].contramap {
    case Empty         => "empty"
    case Number(v)     => v.toString
    case Text(v)       => s"\"$v\""
    case Bool(v)       => v.toString
    case Error(v)      => s"#$v"
  }

/** A cell that didn't match */
case class CellMismatch(
    ref: String,
    expected: ComparableValue,
    actual: ComparableValue
)

object CellMismatch:
  given JsonEncoder[CellMismatch] = DeriveJsonEncoder.gen[CellMismatch]

/** Result of comparing a single range */
case class RangeResult(
    position: String,
    passed: Boolean,
    mismatches: List[CellMismatch]
)

object RangeResult:
  given JsonEncoder[RangeResult] = DeriveJsonEncoder.gen[RangeResult]

/** Result of evaluating a single test case */
case class TestCaseResult(
    caseNum: Int,
    passed: Boolean,
    rangeResults: List[RangeResult],
    error: Option[String] = None
)

object TestCaseResult:
  given JsonEncoder[TestCaseResult] = DeriveJsonEncoder.gen[TestCaseResult]

/** Score for a task (per SpreadsheetBench spec) */
case class TaskScore(
    softRestriction: Double, // passing_cases / 3 (0.0 - 1.0)
    hardRestriction: Int     // 1 if all_pass else 0
)

object TaskScore:
  given JsonEncoder[TaskScore] = DeriveJsonEncoder.gen[TaskScore]

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

/** Token usage from API */
case class TokenUsage(
    inputTokens: Long,
    outputTokens: Long
):
  def totalTokens: Long = inputTokens + outputTokens

object TokenUsage:
  given JsonEncoder[TokenUsage] = DeriveJsonEncoder.gen[TokenUsage]

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
    responseText: Option[String] = None  // Full model response for logging
)

object TaskResult:
  given JsonEncoder[TaskResult] = DeriveJsonEncoder.gen[TaskResult]

  def skipped(task: BenchTask, reason: String): TaskResult =
    TaskResult(
      taskId = task.id,
      instruction = task.instruction,
      instructionType = task.instructionType,
      score = TaskScore(0.0, 0),
      testCaseResults = Nil,
      usage = TokenUsage(0, 0),
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
  given JsonEncoder[AggregateStats] = DeriveJsonEncoder.gen[AggregateStats]

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
    aggregate: AggregateStats,
    results: List[TaskResult]
)

object BenchmarkReport:
  given JsonEncoder[BenchmarkReport] = DeriveJsonEncoder.gen[BenchmarkReport]

// ============================================================================
// CLI Config
// ============================================================================

case class BenchConfig(
    dataset: String = "sample_data_200",
    dataDir: Path = Path.of(Option(java.lang.System.getenv("HOME")).getOrElse("."), "git/SpreadsheetBench/data"),
    limit: Option[Int] = None,
    taskIds: Option[List[String]] = None,  // Specific task IDs to run
    skipIds: Set[String] = Set.empty,      // Task IDs to skip
    category: Option[String] = None,       // Filter by instruction_type
    parallelism: Int = 4,
    verbose: Boolean = false,
    outputDir: Path = Path.of("results")
)
