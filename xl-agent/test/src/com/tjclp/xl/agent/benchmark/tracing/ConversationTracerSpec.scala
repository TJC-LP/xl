package com.tjclp.xl.agent.benchmark.tracing

import cats.effect.IO
import munit.CatsEffectSuite
import com.tjclp.xl.agent.TokenUsage

import java.nio.file.{Files, Path}

/** Pins stop_reason capture in the trace metadata (issue #340) */
class ConversationTracerSpec extends CatsEffectSuite:

  private val tempDir = FunFixture[Path](
    setup = _ => Files.createTempDirectory("conversation-tracer-spec"),
    teardown = { dir =>
      import scala.jdk.CollectionConverters.*
      import scala.util.Using
      Using.resource(Files.walk(dir)) { stream =>
        stream.iterator().asScala.toList.reverse.foreach(p => Files.deleteIfExists(p))
      }
    }
  )

  tempDir.test("complete records the stop reason in the saved metadata") { dir =>
    for
      tracer <- ConversationTracer.create(
        outputDir = dir,
        taskId = "13894",
        skillName = "xl",
        caseNum = 3,
        streaming = false,
        model = Some("claude-test")
      )
      _ <- tracer.complete(
        TokenUsage(100, 16819),
        passed = false,
        error = Some("turn truncated: stop_reason=max_tokens (raise --max-tokens)"),
        stopReason = Some("max_tokens")
      )
      traceDir <- tracer.save()
      traceJson <- IO.blocking(Files.readString(traceDir.resolve("conversation.json")))
    yield
      val parsed = io.circe.parser.parse(traceJson).getOrElse(fail("unparseable trace JSON"))
      val meta = parsed.hcursor.downField("metadata")
      assertEquals(meta.get[String]("stopReason"), Right("max_tokens"))
      assertEquals(
        meta.get[String]("error"),
        Right("turn truncated: stop_reason=max_tokens (raise --max-tokens)")
      )
  }

  tempDir.test("stop reason is null in metadata when never provided") { dir =>
    for
      tracer <- ConversationTracer.create(
        outputDir = dir,
        taskId = "13894",
        skillName = "xl",
        caseNum = 1
      )
      _ <- tracer.complete(TokenUsage.zero, passed = true)
      traceDir <- tracer.save()
      traceJson <- IO.blocking(Files.readString(traceDir.resolve("conversation.json")))
    yield
      val parsed = io.circe.parser.parse(traceJson).getOrElse(fail("unparseable trace JSON"))
      val meta = parsed.hcursor.downField("metadata")
      assertEquals(meta.get[Option[String]]("stopReason"), Right(None))
  }
