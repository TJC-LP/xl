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

  test("AgentConfig defaults") {
    val config = AgentConfig()
    // Default is Sonnet 5 (near-Opus coding/agentic quality at Sonnet cost)
    assertEquals(config.model, "claude-sonnet-5")
    assertEquals(config.maxTokens, 16384)
    assertEquals(config.verbose, false)
  }
