package com.tjclp.xl.agent

import java.util.Optional

import cats.effect.{IO, Ref}
import com.anthropic.models.beta.messages.*
import munit.CatsEffectSuite

import scala.jdk.OptionConverters.*

/**
 * Pins the pause_turn auto-resume loop (issue #344): a turn the API pauses mid-server-tool-loop is
 * re-sent with the paused assistant turns appended verbatim until it finishes cleanly (or the
 * resume cap trips), producing ONE merged result with usage summed across every billed request.
 */
class AgentPauseTurnResumeSpec extends CatsEffectSuite:

  private def usage(inputTokens: Long, outputTokens: Long): BetaUsage =
    BetaUsage
      .builder()
      .inputTokens(inputTokens)
      .outputTokens(outputTokens)
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

  private def response(
    id: String,
    stopReason: BetaStopReason,
    inputTokens: Long,
    outputTokens: Long,
    text: Option[String] = None
  ): BetaMessage =
    val content = text.fold(java.util.List.of[BetaContentBlock]()) { t =>
      java.util.List.of(
        BetaContentBlock.ofText(
          BetaTextBlock.builder().text(t).citations(java.util.List.of[BetaTextCitation]()).build()
        )
      )
    }
    BetaMessage
      .builder()
      .id(id)
      .content(content)
      .model("claude-test")
      .stopReason(stopReason)
      .stopSequence(Optional.empty[String]())
      .container(Optional.empty[BetaContainer]())
      .contextManagement(Optional.empty[BetaContextManagementResponse]())
      .diagnostics(Optional.empty[BetaDiagnostics]())
      .stopDetails(Optional.empty[BetaRefusalStopDetails]())
      .usage(usage(inputTokens, outputTokens))
      .build()

  private def finalStop(turn: Agent.TurnResponses): Option[String] =
    turn.last.stopReason().toScala.map(_.asString)

  test("pause_turn -> end_turn resumes once and merges into one clean result") {
    val paused =
      response("msg_1", BetaStopReason.PAUSE_TURN, 100L, 50L, text = Some("partial work"))
    val done = response("msg_2", BetaStopReason.END_TURN, 200L, 75L, text = Some("final answer"))
    for
      calls <- Ref.of[IO, Vector[Vector[BetaMessage]]](Vector.empty)
      turn <- Agent.sendWithPauseTurnResume { priorTurns =>
        calls.update(_ :+ priorTurns) *>
          calls.get.map(seen => if seen.size == 1 then paused else done)
      }
      recorded <- calls.get
    yield
      // Initial send has no prior turns; the resume re-sends with the paused turn appended
      assertEquals(recorded.map(_.map(_.id())), Vector(Vector(), Vector("msg_1")))
      assertEquals(turn.paused.map(_.id()), Vector("msg_1"))
      assertEquals(turn.last.id(), "msg_2")
      assertEquals(turn.resumesUsed, 1)
      // One merged result: summed usage, concatenated text, and no error for a clean finish
      assertEquals(Agent.mergedUsage(turn), TokenUsage(300L, 125L))
      assertEquals(Agent.mergedResponseText(turn), "partial work\nfinal answer")
      assertEquals(StopReasonPolicy.errorFor(finalStop(turn), turn.resumesUsed), None)
  }

  test("a clean first response never resumes") {
    val done = response("msg_1", BetaStopReason.END_TURN, 10L, 5L)
    for
      count <- Ref.of[IO, Int](0)
      turn <- Agent.sendWithPauseTurnResume(_ => count.update(_ + 1).as(done))
      sends <- count.get
    yield
      assertEquals(sends, 1)
      assertEquals(turn.paused, Vector.empty[BetaMessage])
      assertEquals(turn.resumesUsed, 0)
      assertEquals(Agent.mergedUsage(turn), TokenUsage(10L, 5L))
      assertEquals(StopReasonPolicy.errorFor(finalStop(turn), turn.resumesUsed), None)
  }

  test("persistent pause_turn stops after the resume cap with an exhaustion error") {
    val paused = response("msg_p", BetaStopReason.PAUSE_TURN, 10L, 5L)
    for
      count <- Ref.of[IO, Int](0)
      turn <- Agent.sendWithPauseTurnResume(_ => count.update(_ + 1).as(paused))
      sends <- count.get
    yield
      // 1 initial send + MaxPauseTurnResumes resumes, then give up
      assertEquals(sends, 1 + StopReasonPolicy.MaxPauseTurnResumes)
      assertEquals(turn.resumesUsed, StopReasonPolicy.MaxPauseTurnResumes)
      assertEquals(finalStop(turn), Some("pause_turn"))
      // Usage still sums across every billed request, including the abandoned ones
      assertEquals(Agent.mergedUsage(turn), TokenUsage(40L, 20L))
      val error = StopReasonPolicy.errorFor(finalStop(turn), turn.resumesUsed)
      assert(
        error.exists(_.contains("auto-resume exhausted after 3 resumes")),
        s"cap error must name the exhausted resumes: $error"
      )
  }

  test("max_tokens after a resume is still the truncation error, not resume exhaustion") {
    assertEquals(
      StopReasonPolicy.errorFor(Some("max_tokens"), resumesUsed = 2),
      Some("turn truncated: stop_reason=max_tokens (raise --max-tokens)")
    )
  }
