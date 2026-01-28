package com.tjclp.xl.agent.benchmark.token

import cats.effect.IO
import cats.syntax.all.*
import com.anthropic.client.AnthropicClient as JAnthropicClient
import com.anthropic.core.JsonValue
import com.anthropic.models.beta.AnthropicBeta
import com.anthropic.models.beta.messages.{BetaJsonOutputFormat, MessageCreateParams}
import com.tjclp.xl.agent.error.AgentError
import io.circe.*
import io.circe.parser.*

import scala.jdk.CollectionConverters.*
import scala.jdk.OptionConverters.*

/** Grading result from Opus 4.5 */
case class GradeResult(
  grade: Grade,
  reason: String
)

object GradeResult:
  given Decoder[GradeResult] = Decoder.instance { c =>
    for
      gradeStr <- c.get[String]("grade")
      grade <- Grade.given_Decoder_Grade.decodeJson(Json.fromString(gradeStr))
      reason <- c.get[String]("reason")
    yield GradeResult(grade, reason)
  }

/** Grades task responses using Opus 4.5 with structured outputs */
object TokenGrader:

  private val Model = "claude-opus-4-5-20251101"

  /** Grade a task response */
  def grade(
    client: JAnthropicClient,
    task: TokenTask,
    responseText: String
  ): IO[GradeResult] =
    IO.blocking {
      val prompt = buildGradingPrompt(task, responseText)

      // Build the JSON schema for structured output
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
        .model(Model)
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

      decode[GradeResult](textContent) match
        case Right(result) => result
        case Left(e) =>
          throw AgentError.GradingFailed(s"Failed to parse grading response: ${e.getMessage}")
    }.adaptError { case e: Exception =>
      AgentError.GradingFailed(e.getMessage)
    }

  /** Grade a list of task results */
  def gradeAll(
    client: JAnthropicClient,
    tasks: List[TokenTask],
    results: List[TokenTaskResult]
  ): IO[List[TokenTaskResult]] =
    results.traverse { result =>
      // Only grade successful results with response text
      if result.success && result.responseText.isDefined then
        tasks.find(_.id == result.taskId) match
          case Some(task) if task.expectedAnswer.isDefined =>
            for
              gradeResult <- grade(client, task, result.responseText.getOrElse(""))
              _ <- IO.println(s"   [${result.taskId}] ${result.approach}: ${gradeResult.grade}")
            yield result.copy(
              grade = Some(gradeResult.grade),
              gradeReasoning = Some(gradeResult.reason)
            )
          case _ =>
            // No expected answer to grade against
            IO.pure(result)
      else IO.pure(result)
    }

  /** Build the grading prompt */
  private def buildGradingPrompt(task: TokenTask, responseText: String): String =
    val expectedAnswer = task.expectedAnswer.getOrElse("No expected answer provided")

    s"""You are grading an AI's response to an Excel analysis task.

TASK: ${task.name}
PROMPT: ${task.xlPrompt}

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
