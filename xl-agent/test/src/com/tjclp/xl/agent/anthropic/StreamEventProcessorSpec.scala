package com.tjclp.xl.agent.anthropic

import java.time.Instant
import java.util.Optional

import cats.effect.{IO, Ref}
import cats.effect.std.Queue
import com.anthropic.models.beta.messages.*
import munit.CatsEffectSuite
import com.tjclp.xl.agent.AgentEvent

/**
 * Truncated-stream usage recovery (issue #340): TurnComplete only fires on message_stop, so a
 * stream that dies mid-turn loses the open turn's token accounting unless flushPartialTurn emits it
 * — and a flush after a completed turn must never double-count.
 */
class StreamEventProcessorSpec extends CatsEffectSuite:

  private def messageStart(inputTokens: Long): BetaRawMessageStreamEvent =
    BetaRawMessageStreamEvent.ofMessageStart(
      BetaRawMessageStartEvent
        .builder()
        .message(
          BetaMessage
            .builder()
            .id("msg_test")
            .content(java.util.List.of())
            .model("claude-test")
            .stopReason(Optional.empty[BetaStopReason]())
            .stopSequence(Optional.empty[String]())
            .container(Optional.empty[BetaContainer]())
            .contextManagement(Optional.empty[BetaContextManagementResponse]())
            .diagnostics(Optional.empty[BetaDiagnostics]())
            .stopDetails(Optional.empty[BetaRefusalStopDetails]())
            .usage(
              BetaUsage
                .builder()
                .inputTokens(inputTokens)
                .outputTokens(0L)
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
            )
            .build()
        )
        .build()
    )

  private def messageDelta(outputTokens: Long): BetaRawMessageStreamEvent =
    BetaRawMessageStreamEvent.ofMessageDelta(
      BetaRawMessageDeltaEvent
        .builder()
        .contextManagement(Optional.empty[BetaContextManagementResponse]())
        .delta(
          BetaRawMessageDeltaEvent.Delta
            .builder()
            .stopReason(Optional.empty[BetaStopReason]())
            .stopSequence(Optional.empty[String]())
            .container(Optional.empty[BetaContainer]())
            .stopDetails(Optional.empty[BetaRefusalStopDetails]())
            .build()
        )
        .usage(
          BetaMessageDeltaUsage
            .builder()
            .outputTokens(outputTokens)
            .inputTokens(Optional.empty[java.lang.Long]())
            .cacheCreationInputTokens(Optional.empty[java.lang.Long]())
            .cacheReadInputTokens(Optional.empty[java.lang.Long]())
            .serverToolUse(Optional.empty[BetaServerToolUsage]())
            .iterations(Optional.empty[java.util.List[BetaMessageDeltaUsage.Iteration]]())
            .outputTokensDetails(Optional.empty[BetaOutputTokensDetails]())
            .build()
        )
        .build()
    )

  private def messageStop: BetaRawMessageStreamEvent =
    BetaRawMessageStreamEvent.ofMessageStop(BetaRawMessageStopEvent.builder().build())

  private def newProcessor
    : IO[(StreamEventProcessor, Ref[IO, Vector[AgentEvent]], Queue[IO, AgentEvent])] =
    for
      queue <- Queue.unbounded[IO, AgentEvent]
      seen <- Ref.of[IO, Vector[AgentEvent]](Vector.empty)
      processor <- StreamEventProcessor.create(queue, e => seen.update(_ :+ e))
    yield (processor, seen, queue)

  private def turnCompletes(events: Vector[AgentEvent]) =
    events.collect { case AgentEvent.TurnComplete(usage) => usage }

  test("flushPartialTurn emits a synthetic TurnComplete for a turn cut off mid-stream") {
    for
      (processor, seen, _) <- newProcessor
      _ <- processor.process(messageStart(1000L))
      _ <- processor.process(messageDelta(500L))
      // The stream dies here: no message_stop arrives
      _ <- processor.flushPartialTurn
      events <- seen.get
    yield
      val turns = turnCompletes(events)
      assertEquals(turns.map(u => (u.inputTokens, u.outputTokens)), Vector((1000L, 500L)))
      assertEquals(turns.headOption.map(_.turnNum), Some(1))
  }

  test("flushPartialTurn is a no-op after message_stop reported the turn") {
    for
      (processor, seen, _) <- newProcessor
      _ <- processor.process(messageStart(1000L))
      _ <- processor.process(messageDelta(500L))
      _ <- processor.process(messageStop)
      _ <- processor.flushPartialTurn
      events <- seen.get
    yield
      val turns = turnCompletes(events)
      assertEquals(
        turns.map(u => (u.inputTokens, u.outputTokens)),
        Vector((1000L, 500L)),
        "a completed turn must never be double-counted"
      )
  }

  test("flushPartialTurn emits at most once for the same open turn") {
    for
      (processor, seen, _) <- newProcessor
      _ <- processor.process(messageStart(1000L))
      _ <- processor.process(messageDelta(500L))
      _ <- processor.flushPartialTurn
      _ <- processor.flushPartialTurn
      events <- seen.get
    yield assertEquals(turnCompletes(events).size, 1)
  }

  test("flushPartialTurn is a no-op before any turn started") {
    for
      (processor, seen, _) <- newProcessor
      _ <- processor.flushPartialTurn
      events <- seen.get
    yield assertEquals(turnCompletes(events), Vector.empty)
  }

  test("openTurnUsage is pure: open turns report deltas, closed turns report nothing") {
    val now = Instant.parse("2026-07-09T00:00:00Z")
    val open = ProcessorState(
      turnNum = 2,
      turnStartTime = Some(now.minusSeconds(5)),
      prevCumulativeInput = 1000L,
      prevCumulativeOutput = 200L,
      lastCumulativeInput = 3500L,
      lastCumulativeOutput = 700L,
      turnOpen = true
    )
    val usage = open.openTurnUsage(now).getOrElse(fail("open turn must report usage"))
    assertEquals(usage.turnNum, 2)
    assertEquals(usage.inputTokens, 2500L)
    assertEquals(usage.outputTokens, 500L)
    assertEquals(usage.cumulativeInputTokens, 3500L)
    assertEquals(usage.cumulativeOutputTokens, 700L)
    assertEquals(usage.durationMs, 5000L)

    assertEquals(open.copy(turnOpen = false).openTurnUsage(now), None)
    assertEquals(ProcessorState().openTurnUsage(now), None)
  }
