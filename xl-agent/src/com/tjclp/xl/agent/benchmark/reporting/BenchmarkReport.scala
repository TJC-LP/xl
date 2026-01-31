package com.tjclp.xl.agent.benchmark.reporting

import com.tjclp.xl.agent.TokenUsage
import com.tjclp.xl.agent.benchmark.grading.*
import com.tjclp.xl.agent.benchmark.task.*
import io.circe.*
import io.circe.generic.semiauto.*
import io.circe.syntax.*

import java.time.{Duration, Instant}

// ============================================================================
// Unified Benchmark Report Model
// ============================================================================

/** Unified benchmark report supporting all benchmark types */
case class BenchmarkReport(
  header: ReportHeader,
  summary: ReportSummary,
  skillResults: Map[String, SkillSummary],
  taskResults: List[TaskResultEntry],
  metadata: ReportMetadata = ReportMetadata.empty
):
  def toJsonPretty: String = this.asJson.spaces2

object BenchmarkReport:
  given Encoder[BenchmarkReport] = deriveEncoder

// ============================================================================
// Report Header
// ============================================================================

/** Header information for the report */
case class ReportHeader(
  title: String,
  benchmarkType: String,
  timestamp: Instant,
  duration: Duration,
  model: String,
  skills: List[String]
):
  def formattedTimestamp: String =
    java.time.format.DateTimeFormatter.ISO_INSTANT.format(timestamp)

  def formattedDuration: String =
    val secs = duration.toSeconds
    val mins = secs / 60
    val remSecs = secs % 60
    if mins > 0 then f"${mins}m ${remSecs}s" else s"${secs}s"

object ReportHeader:
  given Encoder[ReportHeader] = Encoder.instance { h =>
    Json.obj(
      "title" -> h.title.asJson,
      "benchmarkType" -> h.benchmarkType.asJson,
      "timestamp" -> h.formattedTimestamp.asJson,
      "durationMs" -> h.duration.toMillis.asJson,
      "durationFormatted" -> h.formattedDuration.asJson,
      "model" -> h.model.asJson,
      "skills" -> h.skills.asJson
    )
  }

// ============================================================================
// Report Summary
// ============================================================================

/** Aggregate summary statistics */
case class ReportSummary(
  totalTasks: Int,
  completedTasks: Int,
  skippedTasks: Int,
  passedTasks: Int,
  failedTasks: Int,
  averageScore: Double,
  totalTokens: TokenSummary,
  averageLatencyMs: Long
):
  def passRate: Double = if completedTasks > 0 then passedTasks.toDouble / completedTasks else 0.0
  def passRatePercent: Int = (passRate * 100).toInt

object ReportSummary:
  given Encoder[ReportSummary] = deriveEncoder

  def fromResults(results: List[TaskResultEntry]): ReportSummary =
    val completed = results.filterNot(_.skipped)
    val passed = completed.count(_.passed)
    val totalUsage = completed.foldLeft(TokenSummary.zero)((acc, r) => acc + r.usage)
    val avgScore =
      if completed.isEmpty then 0.0
      else completed.map(_.score.normalized).sum / completed.size

    ReportSummary(
      totalTasks = results.length,
      completedTasks = completed.length,
      skippedTasks = results.count(_.skipped),
      passedTasks = passed,
      failedTasks = completed.length - passed,
      averageScore = avgScore,
      totalTokens = totalUsage,
      averageLatencyMs =
        if completed.isEmpty then 0 else completed.map(_.latencyMs).sum / completed.length
    )

// ============================================================================
// Token Summary
// ============================================================================

/** Token usage summary */
case class TokenSummary(
  inputTokens: Long,
  outputTokens: Long,
  cacheCreation: Long = 0,
  cacheRead: Long = 0
):
  def total: Long = inputTokens + outputTokens

  def +(other: TokenSummary): TokenSummary = TokenSummary(
    inputTokens + other.inputTokens,
    outputTokens + other.outputTokens,
    cacheCreation + other.cacheCreation,
    cacheRead + other.cacheRead
  )

object TokenSummary:
  val zero: TokenSummary = TokenSummary(0, 0, 0, 0)

  def fromUsage(usage: TokenUsage): TokenSummary = TokenSummary(
    inputTokens = usage.inputTokens,
    outputTokens = usage.outputTokens,
    cacheCreation = 0,
    cacheRead = 0
  )

  given Encoder[TokenSummary] = deriveEncoder

// ============================================================================
// Skill Summary
// ============================================================================

/** Summary for a single skill */
case class SkillSummary(
  skillName: String,
  displayName: String,
  taskCount: Int,
  passCount: Int,
  failCount: Int,
  averageScore: Double,
  totalTokens: TokenSummary,
  averageLatencyMs: Long
):
  def passRate: Double = if taskCount > 0 then passCount.toDouble / taskCount else 0.0
  def passRatePercent: Int = (passRate * 100).toInt

object SkillSummary:
  given Encoder[SkillSummary] = deriveEncoder

  def fromResults(
    skillName: String,
    displayName: String,
    results: List[TaskResultEntry]
  ): SkillSummary =
    val skillResults = results.filter(_.skill == skillName)
    val completed = skillResults.filterNot(_.skipped)
    val passed = completed.count(_.passed)
    val totalUsage = completed.foldLeft(TokenSummary.zero)((acc, r) => acc + r.usage)
    val avgScore =
      if completed.isEmpty then 0.0
      else completed.map(_.score.normalized).sum / completed.size

    SkillSummary(
      skillName = skillName,
      displayName = displayName,
      taskCount = completed.length,
      passCount = passed,
      failCount = completed.length - passed,
      averageScore = avgScore,
      totalTokens = totalUsage,
      averageLatencyMs =
        if completed.isEmpty then 0 else completed.map(_.latencyMs).sum / completed.length
    )

// ============================================================================
// Task Result Entry
// ============================================================================

/** Result for a single task execution */
case class TaskResultEntry(
  taskId: String,
  skill: String,
  caseNum: Option[Int],
  instruction: String,
  category: String,
  score: Score,
  gradeDetails: GradeDetails,
  usage: TokenSummary,
  latencyMs: Long,
  skipped: Boolean = false,
  skipReason: Option[String] = None,
  error: Option[String] = None,
  traceFile: Option[String] = None
):
  def passed: Boolean = gradeDetails.passed && !skipped
  def normalizedScore: Double = score.normalized

object TaskResultEntry:
  given Encoder[TaskResultEntry] = Encoder.instance { e =>
    Json.obj(
      "taskId" -> e.taskId.asJson,
      "skill" -> e.skill.asJson,
      "caseNum" -> e.caseNum.asJson,
      "instruction" -> e.instruction.take(200).asJson,
      "category" -> e.category.asJson,
      "score" -> e.score.normalized.asJson,
      "passed" -> e.passed.asJson,
      "reasoning" -> e.gradeDetails.reasoning.asJson,
      "mismatches" -> e.gradeDetails.mismatches.asJson,
      "inputTokens" -> e.usage.inputTokens.asJson,
      "outputTokens" -> e.usage.outputTokens.asJson,
      "latencyMs" -> e.latencyMs.asJson,
      "skipped" -> e.skipped.asJson,
      "skipReason" -> e.skipReason.asJson,
      "error" -> e.error.asJson,
      "traceFile" -> e.traceFile.asJson
    )
  }

  def skipped(taskId: String, skill: String, reason: String): TaskResultEntry =
    TaskResultEntry(
      taskId = taskId,
      skill = skill,
      caseNum = None,
      instruction = "",
      category = "",
      score = Score.BinaryScore.Fail,
      gradeDetails = GradeDetails.fail(reason),
      usage = TokenSummary.zero,
      latencyMs = 0,
      skipped = true,
      skipReason = Some(reason)
    )

// ============================================================================
// Report Metadata
// ============================================================================

/** Additional metadata for the report */
case class ReportMetadata(
  dataset: Option[String] = None,
  dataDir: Option[String] = None,
  config: Map[String, String] = Map.empty,
  notes: Option[String] = None
)

object ReportMetadata:
  val empty: ReportMetadata = ReportMetadata()
  given Encoder[ReportMetadata] = deriveEncoder

// ============================================================================
// Report Builder (Immutable)
// ============================================================================

/** Immutable builder for constructing reports */
case class ReportBuilder(
  title: String = "Benchmark Report",
  benchmarkType: String = "unified",
  startTime: Instant = Instant.now(),
  model: String = "",
  skills: List[String] = Nil,
  results: List[TaskResultEntry] = Nil,
  metadata: ReportMetadata = ReportMetadata.empty
):
  def withTitle(t: String): ReportBuilder = copy(title = t)
  def withBenchmarkType(bt: String): ReportBuilder = copy(benchmarkType = bt)
  def withStartTime(t: Instant): ReportBuilder = copy(startTime = t)
  def withModel(m: String): ReportBuilder = copy(model = m)
  def withSkills(s: List[String]): ReportBuilder = copy(skills = s)
  def withMetadata(m: ReportMetadata): ReportBuilder = copy(metadata = m)

  def addResult(result: TaskResultEntry): ReportBuilder =
    copy(results = results :+ result)

  def addResults(rs: List[TaskResultEntry]): ReportBuilder =
    copy(results = results ++ rs)

  def build(): BenchmarkReport =
    val endTime = Instant.now()
    val duration = Duration.between(startTime, endTime)

    val header = ReportHeader(
      title = title,
      benchmarkType = benchmarkType,
      timestamp = startTime,
      duration = duration,
      model = model,
      skills = skills
    )

    val summary = ReportSummary.fromResults(results)

    val skillSummaries = skills.map { skill =>
      skill -> SkillSummary.fromResults(skill, skill, results)
    }.toMap

    BenchmarkReport(
      header = header,
      summary = summary,
      skillResults = skillSummaries,
      taskResults = results,
      metadata = metadata
    )

object ReportBuilder:
  def apply(): ReportBuilder = ReportBuilder(startTime = Instant.now())
