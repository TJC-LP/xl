package com.tjclp.xl.agent.benchmark.skills

import cats.effect.{Clock, IO}
import munit.CatsEffectSuite
import com.tjclp.xl.agent.{AgentEvent, TurnUsage}
import com.tjclp.xl.agent.benchmark.tracing.ConversationTracer

import java.nio.file.{Files, Path}

/**
 * Pins the trace-on-failure path (issue #334): when a case dies mid-run, the conversation trace
 * collected up to the failure must land on disk with passed=false and the error recorded, and the
 * handler must be total (never raise, even when saving the trace itself fails).
 */
class CaseFailureSpec extends CatsEffectSuite:

  private val tempDir = FunFixture[Path](
    setup = _ => Files.createTempDirectory("case-failure-spec"),
    teardown = { dir =>
      import scala.jdk.CollectionConverters.*
      import scala.util.Using
      Using.resource(Files.walk(dir)) { stream =>
        stream.iterator().asScala.toList.reverse.foreach(p => Files.deleteIfExists(p))
      }
    }
  )

  tempDir.test("withPartialTrace saves the partial trace with passed=false and the error") { dir =>
    for
      tracer <- ConversationTracer.create(
        outputDir = dir,
        taskId = "13894",
        skillName = "xl",
        caseNum = 3,
        streaming = false,
        model = Some("claude-test")
      )
      _ <- tracer.onEvent(AgentEvent.TextOutput("partial progress before the crash"))
      start <- Clock[IO].monotonic.map(_.toMillis)
      result <- CaseFailure.withPartialTrace(
        tracer,
        caseNum = 3,
        startTimeMs = start - 1234,
        error = new RuntimeException("boom mid-run")
      )
      traceJson <- IO.blocking {
        val traceDir = result.tracePath.getOrElse(fail("expected a saved trace path"))
        Files.readString(traceDir.resolve("conversation.json"))
      }
    yield
      assertEquals(result.caseNum, 3)
      assertEquals(result.passed, false)
      assertEquals(result.error, Some("RuntimeException: boom mid-run"))
      assert(result.latencyMs >= 1234, s"latency should cover the run so far: ${result.latencyMs}")

      val parsed = io.circe.parser.parse(traceJson).getOrElse(fail("unparseable trace JSON"))
      val meta = parsed.hcursor.downField("metadata")
      assertEquals(meta.get[Boolean]("passed"), Right(false))
      assertEquals(meta.get[String]("error"), Right("RuntimeException: boom mid-run"))
      assert(
        traceJson.contains("partial progress before the crash"),
        "events collected before the failure must survive in the trace"
      )
  }

  tempDir.test("withPartialTrace sums the partial trace's per-turn usage (issue #340)") { dir =>
    for
      tracer <- ConversationTracer.create(
        outputDir = dir,
        taskId = "13894",
        skillName = "xl",
        caseNum = 3,
        streaming = false,
        model = Some("claude-test")
      )
      // Two completed turns before the crash: per-turn deltas must be summed
      // (not the cumulative fields, which would double-count turn 1)
      _ <- tracer.onEvent(
        AgentEvent.TurnComplete(TurnUsage(1, 1000L, 200L, 1000L, 200L, 1500L))
      )
      _ <- tracer.onEvent(
        AgentEvent.TurnComplete(TurnUsage(2, 2500L, 300L, 3500L, 500L, 900L))
      )
      start <- Clock[IO].monotonic.map(_.toMillis)
      result <- CaseFailure.withPartialTrace(
        tracer,
        caseNum = 3,
        startTimeMs = start,
        error = new RuntimeException("boom mid-run")
      )
      traceJson <- IO.blocking {
        val traceDir = result.tracePath.getOrElse(fail("expected a saved trace path"))
        Files.readString(traceDir.resolve("conversation.json"))
      }
    yield
      assertEquals(result.usage.inputTokens, 3500L)
      assertEquals(result.usage.outputTokens, 500L)

      val parsed = io.circe.parser.parse(traceJson).getOrElse(fail("unparseable trace JSON"))
      val usage = parsed.hcursor.downField("metadata").downField("usage")
      assertEquals(usage.get[Long]("inputTokens"), Right(3500L))
      assertEquals(usage.get[Long]("outputTokens"), Right(500L))
  }

  tempDir.test("withPartialTrace records zero usage when no turn completed") { dir =>
    for
      tracer <- ConversationTracer.create(
        outputDir = dir,
        taskId = "13894",
        skillName = "xl",
        caseNum = 1
      )
      _ <- tracer.onEvent(AgentEvent.TextOutput("no TurnComplete before the crash"))
      start <- Clock[IO].monotonic.map(_.toMillis)
      result <- CaseFailure.withPartialTrace(
        tracer,
        caseNum = 1,
        startTimeMs = start,
        error = new RuntimeException("early crash")
      )
    yield
      assertEquals(result.usage.inputTokens, 0L)
      assertEquals(result.usage.outputTokens, 0L)
  }

  tempDir.test("withPartialTrace is total when the trace cannot be saved") { dir =>
    val blocker = dir.resolve("blocker")
    for
      // A regular file where save() needs a directory makes the save fail deterministically
      _ <- IO.blocking(Files.createFile(blocker))
      tracer <- ConversationTracer.create(
        outputDir = blocker,
        taskId = "13894",
        skillName = "xl",
        caseNum = 1
      )
      start <- Clock[IO].monotonic.map(_.toMillis)
      result <- CaseFailure.withPartialTrace(
        tracer,
        caseNum = 1,
        startTimeMs = start,
        error = new IllegalStateException("kaput")
      )
    yield
      assertEquals(result.passed, false)
      assertEquals(result.error, Some("IllegalStateException: kaput"))
      assertEquals(result.tracePath, None)
  }

  test("withoutTrace records the error when no tracer exists") {
    CaseFailure.withoutTrace(2, new IllegalArgumentException("no tracer yet")).map { result =>
      assertEquals(result.caseNum, 2)
      assertEquals(result.passed, false)
      assertEquals(result.error, Some("IllegalArgumentException: no tracer yet"))
      assertEquals(result.tracePath, None)
    }
  }

  test("describe is total for exceptions without a message") {
    assertEquals(CaseFailure.describe(new RuntimeException()), "RuntimeException: (no message)")
  }

  test("describe includes exception class and message") {
    assertEquals(
      CaseFailure.describe(new IllegalStateException("oops")),
      "IllegalStateException: oops"
    )
  }
