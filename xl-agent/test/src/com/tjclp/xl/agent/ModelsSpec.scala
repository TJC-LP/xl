package com.tjclp.xl.agent

import munit.CatsEffectSuite
import io.circe.syntax.*

class ModelsSpec extends CatsEffectSuite:

  test("TokenUsage addition") {
    val a = TokenUsage(100, 50)
    val b = TokenUsage(200, 100)
    val result = a + b
    assertEquals(result.inputTokens, 300L)
    assertEquals(result.outputTokens, 150L)
    assertEquals(result.totalTokens, 450L)
  }

  test("TokenUsage addition combines cache tokens") {
    val a = TokenUsage(100, 50, cacheCreationTokens = 10, cacheReadTokens = 5)
    val b = TokenUsage(200, 100, cacheCreationTokens = 20, cacheReadTokens = 15)
    val result = a + b
    assertEquals(result.cacheCreationTokens, 30L)
    assertEquals(result.cacheReadTokens, 20L)
  }

  test("TokenUsage.zero") {
    assertEquals(TokenUsage.zero.inputTokens, 0L)
    assertEquals(TokenUsage.zero.outputTokens, 0L)
    assertEquals(TokenUsage.zero.cacheCreationTokens, 0L)
    assertEquals(TokenUsage.zero.cacheReadTokens, 0L)
  }

  test("TokenUsage decodes legacy JSON without cache fields") {
    val json = io.circe.Json.obj(
      "inputTokens" -> io.circe.Json.fromLong(100),
      "outputTokens" -> io.circe.Json.fromLong(50)
    )
    assertEquals(json.as[TokenUsage], Right(TokenUsage(100, 50)))
  }

  test("AgentEvent JSON encoding") {
    val event: AgentEvent = AgentEvent.TextOutput("Hello")
    val json = event.asJson
    assertEquals(json.hcursor.get[String]("type").toOption, Some("TextOutput"))
    assertEquals(json.hcursor.get[String]("text").toOption, Some("Hello"))
  }

  test("AgentEvent.ToolInvocation JSON encoding") {
    val event: AgentEvent = AgentEvent.ToolInvocation(
      name = "code_execution",
      toolUseId = "tool-123",
      input = io.circe.Json.obj(),
      command = Some("ls -la")
    )
    val json = event.asJson
    assertEquals(json.hcursor.get[String]("type").toOption, Some("ToolInvocation"))
    assertEquals(json.hcursor.get[String]("name").toOption, Some("code_execution"))
    assertEquals(json.hcursor.get[String]("command").toOption, Some("ls -la"))
  }

  test("AgentResult JSON carries the stop reason") {
    val result = AgentResult(
      success = true,
      outputFileId = None,
      outputPath = None,
      usage = TokenUsage.zero,
      latencyMs = 191_000L,
      transcript = Vector.empty,
      responseText = Some("done"),
      error = Some("turn truncated: stop_reason=max_tokens (raise --max-tokens)"),
      stopReason = Some("max_tokens")
    )
    val json = result.asJson
    assertEquals(json.hcursor.get[String]("stopReason").toOption, Some("max_tokens"))

    val absent = result.copy(stopReason = None, error = None)
    assertEquals(absent.asJson.hcursor.get[Option[String]]("stopReason"), Right(None))
  }

  test("AgentConfig defaults") {
    val config = AgentConfig()
    // Default is Sonnet 5 (near-Opus coding/agentic quality at Sonnet cost)
    assertEquals(config.model, "claude-sonnet-5")
    // 32K leaves headroom for adaptive thinking, which counts against maxTokens
    // (GH-340: a 16K cap truncated a real benchmark turn at 16,819 output tokens)
    assertEquals(config.maxTokens, 32768)
    assertEquals(config.verbose, false)
  }
