package com.tjclp.xl.agent.benchmark.tracing

import cats.effect.{IO, Ref}
import cats.syntax.all.*
import com.tjclp.xl.agent.{AgentEvent, TokenUsage}
import io.circe.*
import io.circe.syntax.*

import java.nio.file.{Files, Path}
import java.time.{Duration, Instant}

/** A traced event with timing information */
case class TracedEvent(
  timestamp: Instant,
  relativeMs: Long,
  event: AgentEvent
)

object TracedEvent:
  given Encoder[TracedEvent] = Encoder.instance { te =>
    Json.obj(
      "timestamp" -> te.timestamp.toString.asJson,
      "relativeMs" -> te.relativeMs.asJson,
      "event" -> Encoder[AgentEvent].apply(te.event)
    )
  }

/** Metadata about a conversation run */
case class ConversationMetadata(
  taskId: String,
  skillName: String,
  caseNum: Int,
  startTime: Instant,
  endTime: Option[Instant] = None,
  usage: Option[TokenUsage] = None,
  passed: Option[Boolean] = None,
  error: Option[String] = None
):
  def durationMs: Long =
    endTime.map(e => Duration.between(startTime, e).toMillis).getOrElse(0L)

object ConversationMetadata:
  given Encoder[ConversationMetadata] = Encoder.instance { m =>
    Json.obj(
      "taskId" -> m.taskId.asJson,
      "skillName" -> m.skillName.asJson,
      "caseNum" -> m.caseNum.asJson,
      "startTime" -> m.startTime.toString.asJson,
      "endTime" -> m.endTime.map(_.toString).asJson,
      "durationMs" -> m.durationMs.asJson,
      "usage" -> m.usage.asJson,
      "passed" -> m.passed.asJson,
      "error" -> m.error.asJson
    )
  }

/** Complete conversation trace */
case class ConversationTrace(
  metadata: ConversationMetadata,
  events: Vector[TracedEvent]
):
  def turnCount: Int = events.count {
    case TracedEvent(_, _, AgentEvent.ToolInvocation(_, _, _, _)) => true
    case _ => false
  }

  def toolCallCount: Int = events.count {
    case TracedEvent(_, _, AgentEvent.ToolInvocation(_, _, _, _)) => true
    case _ => false
  }

  def errorCount: Int = events.count {
    case TracedEvent(_, _, AgentEvent.Error(_)) => true
    case _ => false
  }

object ConversationTrace:
  given Encoder[ConversationTrace] = Encoder.instance { ct =>
    Json.obj(
      "metadata" -> ct.metadata.asJson,
      "summary" -> Json.obj(
        "turnCount" -> ct.turnCount.asJson,
        "toolCallCount" -> ct.toolCallCount.asJson,
        "errorCount" -> ct.errorCount.asJson
      ),
      "events" -> ct.events.asJson
    )
  }

/**
 * Captures and optionally streams conversation events
 *
 * Usage:
 * {{{
 * val tracer = ConversationTracer.create(...)
 * agent.runStreaming(task, tracer.onEvent)
 * tracer.complete(usage, passed)
 * tracer.save()
 * }}}
 */
class ConversationTracer private (
  outputDir: Path,
  initialMetadata: ConversationMetadata,
  streaming: Boolean,
  events: Ref[IO, Vector[TracedEvent]],
  metadata: Ref[IO, ConversationMetadata]
):

  /** Callback for Agent.runStreaming */
  def onEvent(event: AgentEvent): IO[Unit] =
    for
      now <- IO.realTimeInstant
      meta <- metadata.get
      relativeMs = Duration.between(meta.startTime, now).toMillis
      traced = TracedEvent(now, relativeMs, event)
      _ <- events.update(_ :+ traced)
      _ <- IO.whenA(streaming)(StreamingConsole.print(traced))
    yield ()

  /** Mark the conversation as complete with final results */
  def complete(
    usage: TokenUsage,
    passed: Boolean,
    error: Option[String] = None
  ): IO[Unit] =
    for
      now <- IO.realTimeInstant
      _ <- metadata.update(
        _.copy(
          endTime = Some(now),
          usage = Some(usage),
          passed = Some(passed),
          error = error
        )
      )
    yield ()

  /** Get the current trace */
  def getTrace: IO[ConversationTrace] =
    for
      meta <- metadata.get
      evts <- events.get
    yield ConversationTrace(meta, evts)

  /** Save the conversation to disk */
  def save(): IO[Path] =
    for
      trace <- getTrace
      dir = outputDir
        .resolve("tasks")
        .resolve(trace.metadata.taskId)
        .resolve(trace.metadata.skillName)
        .resolve(s"case${trace.metadata.caseNum}")
      _ <- IO.blocking(Files.createDirectories(dir))

      // Save markdown
      mdPath = dir.resolve("conversation.md")
      mdContent = TraceFormat.toMarkdown(trace)
      _ <- IO.blocking(Files.writeString(mdPath, mdContent))

      // Save JSON
      jsonPath = dir.resolve("conversation.json")
      jsonContent = trace.asJson.spaces2
      _ <- IO.blocking(Files.writeString(jsonPath, jsonContent))
    yield dir

object ConversationTracer:
  /** Create a new conversation tracer */
  def create(
    outputDir: Path,
    taskId: String,
    skillName: String,
    caseNum: Int,
    streaming: Boolean = false
  ): IO[ConversationTracer] =
    for
      now <- IO.realTimeInstant
      initialMeta = ConversationMetadata(taskId, skillName, caseNum, now)
      eventsRef <- Ref.of[IO, Vector[TracedEvent]](Vector.empty)
      metaRef <- Ref.of[IO, ConversationMetadata](initialMeta)
    yield new ConversationTracer(outputDir, initialMeta, streaming, eventsRef, metaRef)
