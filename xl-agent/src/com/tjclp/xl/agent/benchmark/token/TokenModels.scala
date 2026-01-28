package com.tjclp.xl.agent.benchmark.token

import io.circe.*
import io.circe.generic.semiauto.*
import com.tjclp.xl.agent.TokenUsage

import java.nio.file.Path

// ============================================================================
// Approach Enum
// ============================================================================

/** The approach used to complete a task */
enum Approach derives CanEqual:
  case Xl // Custom xl-cli skill with binary
  case Xlsx // Built-in Anthropic xlsx skill (openpyxl)

object Approach:
  given Encoder[Approach] = Encoder.encodeString.contramap(_.toString.toLowerCase)
  given Decoder[Approach] = Decoder.decodeString.emap {
    case "xl" => Right(Xl)
    case "xlsx" => Right(Xlsx)
    case other => Left(s"Unknown approach: $other")
  }

// ============================================================================
// Grade Enum
// ============================================================================

/** Letter grade for task correctness */
enum Grade derives CanEqual:
  case A, B, C, D, F

object Grade:
  given Encoder[Grade] = Encoder.encodeString.contramap(_.toString)
  given Decoder[Grade] = Decoder.decodeString.emap { s =>
    s.toUpperCase match
      case "A" => Right(A)
      case "B" => Right(B)
      case "C" => Right(C)
      case "D" => Right(D)
      case "F" => Right(F)
      case other => Left(s"Unknown grade: $other")
  }

  /** Convert grade to numeric value (A=4, B=3, C=2, D=1, F=0) */
  def toNumeric(grade: Grade): Int = grade match
    case A => 4
    case B => 3
    case C => 2
    case D => 1
    case F => 0

  /** Calculate average grade from a list */
  def average(grades: List[Grade]): Option[Grade] =
    if grades.isEmpty then None
    else
      val avg = grades.map(toNumeric).sum.toDouble / grades.size
      Some(
        if avg >= 3.5 then A
        else if avg >= 2.5 then B
        else if avg >= 1.5 then C
        else if avg >= 0.5 then D
        else F
      )

// ============================================================================
// Task Definition
// ============================================================================

/** A task definition for the token efficiency benchmark */
case class TokenTask(
  id: String,
  name: String,
  description: String,
  xlPrompt: String,
  xlsxPrompt: String,
  expectedAnswer: Option[String] = None
)

object TokenTask:
  given Encoder[TokenTask] = deriveEncoder
  given Decoder[TokenTask] = deriveDecoder

// ============================================================================
// Task Result
// ============================================================================

/** Result of executing a single task with one approach */
case class TokenTaskResult(
  taskId: String,
  taskName: String,
  approach: Approach,
  success: Boolean,
  inputTokens: Long,
  outputTokens: Long,
  totalTokens: Long,
  latencyMs: Long,
  error: Option[String] = None,
  responseText: Option[String] = None,
  grade: Option[Grade] = None,
  gradeReasoning: Option[String] = None
)

object TokenTaskResult:
  given Encoder[TokenTaskResult] = deriveEncoder
  given Decoder[TokenTaskResult] = deriveDecoder

  def failed(
    taskId: String,
    taskName: String,
    approach: Approach,
    error: String,
    latencyMs: Long
  ): TokenTaskResult =
    TokenTaskResult(
      taskId = taskId,
      taskName = taskName,
      approach = approach,
      success = false,
      inputTokens = 0,
      outputTokens = 0,
      totalTokens = 0,
      latencyMs = latencyMs,
      error = Some(error)
    )

// ============================================================================
// Aggregate Statistics
// ============================================================================

/** Aggregate statistics for one approach */
case class ApproachStats(
  approach: Approach,
  taskCount: Int,
  successCount: Int,
  totalInputTokens: Long,
  totalOutputTokens: Long,
  totalTokens: Long,
  avgLatencyMs: Long,
  grades: List[Grade]
):
  def avgGrade: Option[Grade] = Grade.average(grades)

object ApproachStats:
  given Encoder[ApproachStats] = deriveEncoder

  def fromResults(approach: Approach, results: List[TokenTaskResult]): ApproachStats =
    val filtered = results.filter(_.approach == approach)
    val successful = filtered.filter(_.success)
    ApproachStats(
      approach = approach,
      taskCount = filtered.size,
      successCount = successful.size,
      totalInputTokens = successful.map(_.inputTokens).sum,
      totalOutputTokens = successful.map(_.outputTokens).sum,
      totalTokens = successful.map(_.totalTokens).sum,
      avgLatencyMs =
        if successful.isEmpty then 0 else successful.map(_.latencyMs).sum / successful.size,
      grades = successful.flatMap(_.grade)
    )

// ============================================================================
// Benchmark Report
// ============================================================================

/** Full benchmark report */
case class TokenBenchmarkReport(
  timestamp: String,
  model: String,
  sampleFile: String,
  xlStats: ApproachStats,
  xlsxStats: ApproachStats,
  results: List[TokenTaskResult]
):
  def toJsonPretty: String =
    import io.circe.syntax.*
    this.asJson.spaces2

object TokenBenchmarkReport:
  given Encoder[TokenBenchmarkReport] = deriveEncoder

// ============================================================================
// Configuration
// ============================================================================

/** Configuration for the token efficiency benchmark */
case class TokenConfig(
  taskId: Option[String] = None,
  xlOnly: Boolean = false,
  xlsxOnly: Boolean = false,
  sequential: Boolean = false,
  grading: Boolean = true,
  includeLarge: Boolean = false,
  outputPath: Option[Path] = None,
  xlBinary: Option[Path] = None,
  xlSkill: Option[Path] = None,
  samplePath: Option[Path] = None,
  verbose: Boolean = false
)
