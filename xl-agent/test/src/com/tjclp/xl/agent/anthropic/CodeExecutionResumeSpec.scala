package com.tjclp.xl.agent.anthropic

import java.time.OffsetDateTime
import java.util.Optional

import com.anthropic.models.beta.messages.*
import munit.FunSuite
import com.tjclp.xl.agent.{AgentConfig, UploadedFile}
import com.tjclp.xl.agent.approach.{XlApproachStrategy, XlsxApproachStrategy}

import scala.jdk.CollectionConverters.*
import scala.jdk.OptionConverters.*

/**
 * Pins the pause_turn re-send shape (issue #344), per the Anthropic contract for server-side tools:
 * the paused assistant message goes back into `messages` verbatim (no synthetic user message — the
 * API detects the trailing server-tool state and resumes), and the paused turn's container is
 * reused so files from earlier bash cycles stay visible.
 */
class CodeExecutionResumeSpec extends FunSuite:

  private val config = AgentConfig(model = "claude-test", maxTokens = 1024)

  /**
   * Wire role of a message param. Typed-builder messages carry a known Role; toParam-converted
   * responses carry the raw wire value — read whichever is present.
   */
  private def roleOf(message: BetaMessageParam): String =
    message._role().asString().toScala.getOrElse(message.role().toString)

  private def usage: BetaUsage =
    BetaUsage
      .builder()
      .inputTokens(1L)
      .outputTokens(1L)
      .cacheCreation(Optional.empty[BetaCacheCreation]())
      .cacheCreationInputTokens(Optional.empty[java.lang.Long]())
      .cacheReadInputTokens(Optional.empty[java.lang.Long]())
      .serverToolUse(Optional.empty[BetaServerToolUsage]())
      .serviceTier(Optional.empty[BetaUsage.ServiceTier]())
      .inferenceGeo(Optional.empty[String]())
      .speed(Optional.empty[BetaUsage.Speed]())
      .iterations(Optional.empty[java.util.List[BetaUsage.Iteration]]())
      .outputTokensDetails(Optional.empty[BetaOutputTokensDetails]())
      .build()

  private def pausedResponse(containerId: Option[String]): BetaMessage =
    val builder = BetaMessage
      .builder()
      .id("msg_paused")
      .content(
        java.util.List.of(
          BetaContentBlock.ofText(
            BetaTextBlock
              .builder()
              .text("partial work")
              .citations(java.util.List.of[BetaTextCitation]())
              .build()
          )
        )
      )
      .model("claude-test")
      .stopReason(BetaStopReason.PAUSE_TURN)
      .stopSequence(Optional.empty[String]())
      .contextManagement(Optional.empty[BetaContextManagementResponse]())
      .diagnostics(Optional.empty[BetaDiagnostics]())
      .stopDetails(Optional.empty[BetaRefusalStopDetails]())
      .usage(usage)
    containerId match
      case Some(id) =>
        builder.container(
          BetaContainer
            .builder()
            .id(id)
            .expiresAt(OffsetDateTime.parse("2026-07-15T00:00:00Z"))
            .skills(Optional.empty[java.util.List[BetaSkill]]())
            .build()
        )
      case None => builder.container(Optional.empty[BetaContainer]())
    builder.build()

  test("the initial request carries only the user message and no container") {
    val params = CodeExecution.buildParams(
      config = config,
      systemPrompt = "system",
      userPrompt = "do the task",
      containerUploads = List("file_1"),
      resumeTurns = Vector.empty
    )
    val messages = params.messages().asScala.toList
    assertEquals(messages.map(roleOf), List("user"))
    assertEquals(params.container().toScala, None)
  }

  test("a resume appends the paused assistant turn verbatim and reuses its container") {
    val paused = pausedResponse(containerId = Some("cont_123"))
    val params = CodeExecution.buildParams(
      config = config,
      systemPrompt = "system",
      userPrompt = "do the task",
      containerUploads = List("file_1"),
      resumeTurns = Vector(paused)
    )
    val messages = params.messages().asScala.toList
    assertEquals(messages.map(roleOf), List("user", "assistant"))
    // Verbatim: exactly the paused message's toParam conversion, content untouched
    assertEquals(messages(1), paused.toParam())
    assertEquals(params.container().toScala.flatMap(_.string().toScala), Some("cont_123"))
  }

  test("a resume of a container-less paused turn sets no container") {
    val paused = pausedResponse(containerId = None)
    val params = CodeExecution.buildParams(
      config = config,
      systemPrompt = "system",
      userPrompt = "do the task",
      containerUploads = Nil,
      resumeTurns = Vector(paused)
    )
    assertEquals(params.container().toScala, None)
    assertEquals(params.messages().asScala.toList.lastOption.map(roleOf), Some("assistant"))
  }

  // The production strategies set builder.container(BetaContainerParams with skills) inside
  // configureRequest, and the last .container() write on the builder wins. These tests pin the
  // wire shape after composition with a REAL strategy — not the identity default — so the paused
  // container id can never be silently clobbered into a fresh container again (issue #344).

  private val realStrategies = List(
    "xl" -> new XlApproachStrategy(UploadedFile("file_bin", "xl"), "skill_1").configureRequest,
    "xlsx" -> new XlsxApproachStrategy().configureRequest
  )

  realStrategies.foreach { case (name, strategyConfigure) =>
    test(s"a resume through the real $name strategy still reuses the paused container (#344)") {
      val paused = pausedResponse(containerId = Some("cont_123"))
      val params = CodeExecution.buildParams(
        config = config,
        systemPrompt = "system",
        userPrompt = "do the task",
        containerUploads = List("file_1"),
        resumeTurns = Vector(paused),
        configureRequest = strategyConfigure
      )
      // The paused id must be what reaches the wire — the documented reuse shape: the skill
      // files already live in the existing container, so no skills params are re-declared.
      assertEquals(params.container().toScala.flatMap(_.string().toScala), Some("cont_123"))
      // The rest of the strategy configuration must survive the resume unchanged
      assert(params.tools().toScala.exists(!_.isEmpty), "strategy tools must survive a resume")
    }
  }

  test("an initial request through a real strategy declares the skills container with no id") {
    val params = CodeExecution.buildParams(
      config = config,
      systemPrompt = "system",
      userPrompt = "do the task",
      containerUploads = List("file_1"),
      resumeTurns = Vector.empty,
      configureRequest =
        new XlApproachStrategy(UploadedFile("file_bin", "xl"), "skill_1").configureRequest
    )
    val containerParams = params.container().toScala.flatMap(_.betaContainerParams().toScala)
    assert(containerParams.isDefined, "initial request must carry the strategy's container params")
    assertEquals(containerParams.flatMap(_.id().toScala), None)
    assert(
      containerParams.exists(_.skills().toScala.exists(!_.isEmpty)),
      "initial request must declare the strategy's skills"
    )
  }
