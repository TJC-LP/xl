package com.tjclp.xl.agent.benchmark.grading

import cats.effect.IO
import cats.effect.syntax.concurrent.*
import cats.syntax.all.*
import com.anthropic.client.AnthropicClient as JAnthropicClient
import com.anthropic.core.JsonValue
import com.anthropic.models.beta.AnthropicBeta
import com.anthropic.models.beta.messages.{BetaJsonOutputFormat, MessageCreateParams}
import io.circe.*
import io.circe.parser.*

import scala.jdk.CollectionConverters.*
import scala.jdk.OptionConverters.*

// ============================================================================
// LLM-Based Grader (Opus 4.5 as Judge)
// ============================================================================

/**
 * Grader that uses Opus 4.5 with structured outputs for evaluation.
 *
 * Supports parallel grading via `gradeAll` with configurable concurrency.
 */
class OpusLLMGrader(
  client: JAnthropicClient,
  model: String = OpusLLMGrader.DefaultModel,
  parallelism: Int = OpusLLMGrader.DefaultParallelism
) extends Grader[Score.LetterGrade]:

  val name: String = "llm"

  def canHandle(ctx: GraderContext): Boolean =
    ctx.responseText.isDefined && ctx.expectedAnswer.isDefined

  def grade(ctx: GraderContext): IO[GradeResult[Score.LetterGrade]] =
    (ctx.responseText, ctx.expectedAnswer) match
      case (Some(response), Some(expected)) =>
        gradeInternal(ctx.taskId, ctx.skill, response, expected)
      case _ =>
        IO.pure(
          GradeResult(
            graderName = name,
            score = Score.LetterGrade.F,
            details = GradeDetails.fail("Missing response or expected answer")
          )
        )

  private def gradeInternal(
    taskId: String,
    skill: String,
    responseText: String,
    expectedAnswer: String
  ): IO[GradeResult[Score.LetterGrade]] =
    IO.blocking {
      val prompt = buildGradingPrompt(taskId, responseText, expectedAnswer)

      // Build JSON schema for structured output
      val gradeSchema = BetaJsonOutputFormat.Schema
        .builder()
        .putAdditionalProperty("type", JsonValue.from("object"))
        .putAdditionalProperty(
          "properties",
          JsonValue.from(
            java.util.Map.of(
              "grade",
              java.util.Map.of(
                "type",
                "string",
                "enum",
                java.util.List.of("A", "B", "C", "D", "F"),
                "description",
                "Letter grade for correctness"
              ),
              "reason",
              java.util.Map.of(
                "type",
                "string",
                "description",
                "Brief explanation of the grade"
              )
            )
          )
        )
        .putAdditionalProperty("required", JsonValue.from(java.util.List.of("grade", "reason")))
        .putAdditionalProperty("additionalProperties", JsonValue.from(false))
        .build()

      val outputFormat = BetaJsonOutputFormat
        .builder()
        .schema(gradeSchema)
        .build()

      val params = MessageCreateParams
        .builder()
        .model(model)
        .maxTokens(256L)
        .addUserMessage(prompt)
        .outputFormat(outputFormat)
        .addBeta(AnthropicBeta.of("structured-outputs-2025-11-13"))
        .build()

      val response = client.beta().messages().create(params)

      // Extract text content and parse as JSON
      val textContent = response
        .content()
        .asScala
        .flatMap(_.text().toScala)
        .map(_.text())
        .mkString

      decode[LLMGradeResponse](textContent) match
        case Right(result) =>
          Score.LetterGrade.fromString(result.grade) match
            case Right(grade) =>
              GradeResult(
                graderName = name,
                score = grade,
                details = GradeDetails(
                  passed = grade.isPassing,
                  reasoning = Some(result.reason)
                ),
                metadata = Map(
                  "taskId" -> taskId,
                  "skill" -> skill,
                  "model" -> model
                )
              )
            case Left(e) =>
              GradeResult(
                graderName = name,
                score = Score.LetterGrade.F,
                details = GradeDetails.fail(s"Invalid grade: $e"),
                metadata = Map("taskId" -> taskId, "rawGrade" -> result.grade)
              )
        case Left(e) =>
          GradeResult(
            graderName = name,
            score = Score.LetterGrade.F,
            details = GradeDetails.fail(s"Failed to parse grading response: ${e.getMessage}"),
            metadata = Map("taskId" -> taskId, "rawResponse" -> textContent.take(500))
          )
    }.handleError { e =>
      GradeResult(
        graderName = name,
        score = Score.LetterGrade.F,
        details = GradeDetails.fail(s"LLM grading failed: ${e.getMessage}"),
        metadata = Map("taskId" -> taskId, "error" -> e.getMessage)
      )
    }

  /**
   * Grade multiple contexts in parallel with configured concurrency.
   *
   * @param contexts
   *   List of grading contexts to evaluate
   * @return
   *   Aggregated results with summary statistics
   */
  def gradeAll(contexts: List[GraderContext]): IO[ParallelGradeResults] =
    gradeAllWithParallelism(contexts, parallelism)

  /**
   * Grade multiple contexts with explicit parallelism level.
   *
   * @param contexts
   *   List of grading contexts to evaluate
   * @param n
   *   Maximum number of concurrent API calls
   * @return
   *   Aggregated results with summary statistics
   */
  def gradeAllWithParallelism(contexts: List[GraderContext], n: Int): IO[ParallelGradeResults] =
    contexts
      .parTraverseN(n.max(1))(grade)
      .map { results =>
        val byGrade = results.groupBy(_.score)
        val passing = results.count(_.passed)
        val total = results.length

        ParallelGradeResults(
          results = results,
          summary = GradeSummary(
            total = total,
            passing = passing,
            failing = total - passing,
            byGrade = byGrade.map { case (grade, rs) => grade.toString -> rs.length }
          )
        )
      }

  private def buildGradingPrompt(
    taskId: String,
    responseText: String,
    expectedAnswer: String
  ): String =
    s"""You are grading an AI's response to an Excel analysis task.

TASK ID: $taskId

EXPECTED ANSWER (ground truth):
$expectedAnswer

AI'S ACTUAL RESPONSE:
$responseText

Grade the response on correctness:
- A: Fully correct, all key information present
- B: Mostly correct, minor details missing or slight inaccuracies
- C: Partially correct, some key information wrong or missing
- D: Mostly incorrect, but shows some understanding
- F: Completely wrong or didn't answer the question"""

object OpusLLMGrader:
  val DefaultModel: String = "claude-opus-4-5-20251101"
  val DefaultParallelism: Int = 4

  def apply(client: JAnthropicClient): OpusLLMGrader =
    new OpusLLMGrader(client)

  def apply(client: JAnthropicClient, model: String): OpusLLMGrader =
    new OpusLLMGrader(client, model)

  def apply(client: JAnthropicClient, model: String, parallelism: Int): OpusLLMGrader =
    new OpusLLMGrader(client, model, parallelism)

  def withParallelism(client: JAnthropicClient, parallelism: Int): OpusLLMGrader =
    new OpusLLMGrader(client, DefaultModel, parallelism)

// ============================================================================
// Parallel Grading Results
// ============================================================================

/** Results from parallel grading with summary statistics */
case class ParallelGradeResults(
  results: List[GradeResult[Score.LetterGrade]],
  summary: GradeSummary
):
  def passRate: Double =
    if summary.total == 0 then 0.0
    else summary.passing.toDouble / summary.total

  def averageGrade: Option[Score.LetterGrade] =
    if results.isEmpty then None
    else
      val avg = results.map(_.score.toNumeric).sum.toDouble / results.length
      Some(Score.LetterGrade.fromNumeric(avg.round.toInt))

/** Summary statistics for grading results */
case class GradeSummary(
  total: Int,
  passing: Int,
  failing: Int,
  byGrade: Map[String, Int]
)

/** Internal response structure for LLM grading */
private case class LLMGradeResponse(
  grade: String,
  reason: String
)

private object LLMGradeResponse:
  given Decoder[LLMGradeResponse] = Decoder.instance { c =>
    for
      grade <- c.get[String]("grade")
      reason <- c.get[String]("reason")
    yield LLMGradeResponse(grade, reason)
  }

// ============================================================================
// Sonnet LLM Grader (Faster, cheaper alternative)
// ============================================================================

/** Grader that uses Sonnet for faster/cheaper evaluation */
class SonnetLLMGrader(client: JAnthropicClient, parallelism: Int = OpusLLMGrader.DefaultParallelism)
    extends OpusLLMGrader(client, SonnetLLMGrader.Model, parallelism):
  override val name: String = "llm-sonnet"

object SonnetLLMGrader:
  val Model: String = "claude-sonnet-4-20250514"

  def apply(client: JAnthropicClient): SonnetLLMGrader =
    new SonnetLLMGrader(client)

  def apply(client: JAnthropicClient, parallelism: Int): SonnetLLMGrader =
    new SonnetLLMGrader(client, parallelism)

// ============================================================================
// No-Op LLM Grader (for when LLM grading is disabled)
// ============================================================================

/** A grader that skips LLM evaluation and returns a neutral result */
object NoOpLLMGrader extends Grader[Score.LetterGrade]:
  val name: String = "llm-noop"

  def canHandle(ctx: GraderContext): Boolean = true

  def grade(ctx: GraderContext): IO[GradeResult[Score.LetterGrade]] =
    IO.pure(
      GradeResult(
        graderName = name,
        score = Score.LetterGrade.C, // Neutral grade
        details = GradeDetails(
          passed = true,
          reasoning = Some("LLM grading disabled")
        ),
        metadata = Map("taskId" -> ctx.taskId, "skipped" -> "true")
      )
    )
