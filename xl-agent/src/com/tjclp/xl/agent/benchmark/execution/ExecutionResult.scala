package com.tjclp.xl.agent.benchmark.execution

import com.tjclp.xl.agent.benchmark.grading.*
import com.tjclp.xl.agent.benchmark.task.TaskId
import io.circe.*
import io.circe.generic.semiauto.*
import io.circe.syntax.*

import java.nio.file.Path

// ============================================================================
// Unified Execution Result Model
// ============================================================================

/**
 * Canonical execution result type for all benchmarks.
 *
 * This is the universal result that both SpreadsheetBench and TokenBenchmark produce, enabling
 * unified reporting and comparison.
 *
 * @param taskId
 *   The task identifier
 * @param skill
 *   The skill used for execution
 * @param caseResults
 *   Individual case results (1 for token tasks, N for spreadsheet bench)
 * @param aggregateScore
 *   Combined score across all cases
 * @param usage
 *   Token usage with cache tracking
 * @param latencyMs
 *   Total execution time in milliseconds
 * @param traces
 *   Paths to conversation trace files
 * @param gradeResult
 *   Optional LLM grade result
 * @param error
 *   Error message if execution failed
 */
case class ExecutionResult(
  taskId: TaskId,
  skill: String,
  caseResults: Vector[CaseResult],
  aggregateScore: Score,
  usage: TokenUsage,
  latencyMs: Long,
  traces: Vector[Path] = Vector.empty,
  gradeResult: Option[GradeResult[? <: Score]] = None,
  error: Option[String] = None
):
  /** Check if the execution passed (score >= 0.5) */
  def passed: Boolean = error.isEmpty && aggregateScore.normalized >= 0.5

  /** Total tokens used (input + output) */
  def totalTokens: Long = usage.total

  /** Number of cases that passed */
  def passedCases: Int = caseResults.count(_.passed)

  /** Total number of cases */
  def totalCases: Int = caseResults.length

  /** Pass rate as a fraction */
  def passRate: Double =
    if totalCases == 0 then 0.0 else passedCases.toDouble / totalCases

  /** Get the task ID value as a string */
  def taskIdValue: String = taskId.value

  /** Add a grade result to this execution result */
  def withGrade(grade: GradeResult[? <: Score]): ExecutionResult =
    copy(gradeResult = Some(grade))

  /** Add trace paths */
  def withTraces(paths: Vector[Path]): ExecutionResult =
    copy(traces = paths)

  /** Mark as failed with an error */
  def withError(message: String): ExecutionResult =
    copy(error = Some(message))

object ExecutionResult:
  given Encoder[ExecutionResult] = Encoder.instance { r =>
    Json.obj(
      "taskId" -> r.taskIdValue.asJson,
      "skill" -> r.skill.asJson,
      "caseResults" -> r.caseResults.asJson,
      "aggregateScore" -> r.aggregateScore.normalized.asJson,
      "passed" -> r.passed.asJson,
      "usage" -> r.usage.asJson,
      "latencyMs" -> r.latencyMs.asJson,
      "traces" -> r.traces.map(_.toString).asJson,
      "error" -> r.error.asJson
    )
  }

  /** Create a failed result for an execution error */
  def failed(
    taskId: TaskId,
    skill: String,
    error: String,
    usage: TokenUsage = TokenUsage.zero,
    latencyMs: Long = 0
  ): ExecutionResult =
    ExecutionResult(
      taskId = taskId,
      skill = skill,
      caseResults = Vector.empty,
      aggregateScore = Score.BinaryScore.Fail,
      usage = usage,
      latencyMs = latencyMs,
      error = Some(error)
    )

  /** Create a skipped result */
  def skipped(
    taskId: TaskId,
    skill: String,
    reason: String
  ): ExecutionResult =
    failed(taskId, skill, s"Skipped: $reason")

  /** Create from a single case result */
  def fromSingleCase(
    taskId: TaskId,
    skill: String,
    caseResult: CaseResult,
    usage: TokenUsage,
    latencyMs: Long
  ): ExecutionResult =
    ExecutionResult(
      taskId = taskId,
      skill = skill,
      caseResults = Vector(caseResult),
      aggregateScore = Score.BinaryScore.fromBoolean(caseResult.passed),
      usage = usage,
      latencyMs = latencyMs
    )

  /** Create from multiple case results */
  def fromCases(
    taskId: TaskId,
    skill: String,
    caseResults: Vector[CaseResult],
    usage: TokenUsage,
    latencyMs: Long
  ): ExecutionResult =
    val passed = caseResults.count(_.passed)
    ExecutionResult(
      taskId = taskId,
      skill = skill,
      caseResults = caseResults,
      aggregateScore = Score.FractionalScore(passed, caseResults.length),
      usage = usage,
      latencyMs = latencyMs
    )

// ============================================================================
// Case Result (Per-Test-Case)
// ============================================================================

/**
 * Result for a single test case within a task.
 *
 * @param caseNum
 *   The test case number (1-indexed)
 * @param passed
 *   Whether the case passed
 * @param usage
 *   Token usage for this case
 * @param latencyMs
 *   Execution time for this case
 * @param details
 *   Additional details about the result
 * @param tracePath
 *   Path to the trace file for this case
 */
case class CaseResult(
  caseNum: Int,
  passed: Boolean,
  usage: TokenUsage,
  latencyMs: Long,
  details: CaseDetails = CaseDetails.NoDetails,
  tracePath: Option[Path] = None
):
  /** Get mismatches if this is a file comparison result */
  def mismatches: List[CellMismatch] = details match
    case CaseDetails.FileComparison(ms) => ms
    case _ => Nil

  /** Get response text if this is a token comparison result */
  def responseText: Option[String] = details match
    case CaseDetails.TokenComparison(text, _) => Some(text)
    case _ => None

object CaseResult:
  given Encoder[CaseResult] = Encoder.instance { r =>
    Json.obj(
      "caseNum" -> r.caseNum.asJson,
      "passed" -> r.passed.asJson,
      "usage" -> r.usage.asJson,
      "latencyMs" -> r.latencyMs.asJson,
      "details" -> r.details.asJson,
      "tracePath" -> r.tracePath.map(_.toString).asJson
    )
  }

  /** Create a passed case result */
  def passed(
    caseNum: Int,
    usage: TokenUsage,
    latencyMs: Long,
    details: CaseDetails = CaseDetails.NoDetails
  ): CaseResult =
    CaseResult(caseNum, passed = true, usage, latencyMs, details)

  /** Create a failed case result */
  def failed(
    caseNum: Int,
    usage: TokenUsage,
    latencyMs: Long,
    details: CaseDetails = CaseDetails.NoDetails
  ): CaseResult =
    CaseResult(caseNum, passed = false, usage, latencyMs, details)

// ============================================================================
// Case Details (Discriminated Union)
// ============================================================================

/**
 * Details about how a case was evaluated.
 *
 * Different benchmark types produce different details:
 *   - SpreadsheetBench produces FileComparison with cell mismatches
 *   - TokenBenchmark produces TokenComparison with response text
 */
enum CaseDetails:
  /** File comparison result (SpreadsheetBench style) */
  case FileComparison(mismatches: List[CellMismatch])

  /** Token comparison result (TokenBenchmark style) */
  case TokenComparison(responseText: String, expectedAnswer: Option[String])

  /** No additional details */
  case NoDetails

object CaseDetails:
  given Encoder[CaseDetails] = Encoder.instance {
    case FileComparison(mismatches) =>
      Json.obj(
        "type" -> "file_comparison".asJson,
        "mismatches" -> mismatches.asJson
      )
    case TokenComparison(response, expected) =>
      Json.obj(
        "type" -> "token_comparison".asJson,
        "responseText" -> response.asJson,
        "expectedAnswer" -> expected.asJson
      )
    case NoDetails =>
      Json.obj("type" -> "none".asJson)
  }

  /** Create a file comparison with no mismatches (passed) */
  def filePass: CaseDetails = FileComparison(Nil)

  /** Create a file comparison with mismatches (failed) */
  def fileFail(mismatches: List[CellMismatch]): CaseDetails =
    FileComparison(mismatches)

  /** Create a token comparison result */
  def token(response: String, expected: Option[String] = None): CaseDetails =
    TokenComparison(response, expected)
