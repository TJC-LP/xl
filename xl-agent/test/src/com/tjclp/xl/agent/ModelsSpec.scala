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

  test("TokenUsage.zero") {
    assertEquals(TokenUsage.zero.inputTokens, 0L)
    assertEquals(TokenUsage.zero.outputTokens, 0L)
  }

  test("AgentEvent JSON encoding") {
    val event: AgentEvent = AgentEvent.TextOutput("Hello")
    val json = event.asJson
    assertEquals(json.hcursor.get[String]("type").toOption, Some("text"))
    assertEquals(json.hcursor.get[String]("text").toOption, Some("Hello"))
  }

  test("AgentEvent.ToolInvocation JSON encoding") {
    val event: AgentEvent = AgentEvent.ToolInvocation("code_execution", "ls -la")
    val json = event.asJson
    assertEquals(json.hcursor.get[String]("type").toOption, Some("tool_invocation"))
    assertEquals(json.hcursor.get[String]("name").toOption, Some("code_execution"))
    assertEquals(json.hcursor.get[String]("command").toOption, Some("ls -la"))
  }

  test("AgentConfig defaults") {
    val config = AgentConfig()
    assertEquals(config.model, "claude-opus-4-5-20251101")
    assertEquals(config.maxTokens, 8192)
    assertEquals(config.verbose, false)
  }
