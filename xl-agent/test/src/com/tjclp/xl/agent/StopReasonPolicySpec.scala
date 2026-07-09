package com.tjclp.xl.agent

import com.anthropic.models.beta.messages.BetaStopReason
import munit.FunSuite

/**
 * Pins the stop_reason -> error mapping (issue #340): non-clean stops must become loud agent
 * errors, clean stops and absent stop reasons must stay silent.
 */
class StopReasonPolicySpec extends FunSuite:

  private def errorForConstant(reason: BetaStopReason): Option[String] =
    StopReasonPolicy.errorFor(Some(reason.asString))

  test("clean stops map to no error") {
    assertEquals(errorForConstant(BetaStopReason.END_TURN), None)
    assertEquals(errorForConstant(BetaStopReason.TOOL_USE), None)
    assertEquals(errorForConstant(BetaStopReason.STOP_SEQUENCE), None)
  }

  test("absent stop reason maps to no error") {
    assertEquals(StopReasonPolicy.errorFor(None), None)
  }

  test("max_tokens maps to a truncation error naming the remedy") {
    assertEquals(
      errorForConstant(BetaStopReason.MAX_TOKENS),
      Some("turn truncated: stop_reason=max_tokens (raise --max-tokens)")
    )
  }

  test("refusal maps to a refusal error") {
    assertEquals(
      errorForConstant(BetaStopReason.REFUSAL),
      Some("model refused: stop_reason=refusal")
    )
  }

  test("pause_turn maps to an incomplete-turn error") {
    assertEquals(
      errorForConstant(BetaStopReason.PAUSE_TURN),
      Some("turn incomplete: stop_reason=pause_turn")
    )
  }

  test("remaining SDK constants are flagged, never silent") {
    assertEquals(
      errorForConstant(BetaStopReason.COMPACTION),
      Some("turn incomplete: stop_reason=compaction")
    )
    assertEquals(
      errorForConstant(BetaStopReason.MODEL_CONTEXT_WINDOW_EXCEEDED),
      Some("turn incomplete: stop_reason=model_context_window_exceeded")
    )
  }

  test("every SDK stop reason constant has a deliberate mapping") {
    // If the SDK grows a new constant this fails, forcing a policy decision
    val known = BetaStopReason.Value.values.toList.filterNot(_ == BetaStopReason.Value._UNKNOWN)
    assertEquals(known.size, 8, s"unexpected BetaStopReason constants: $known")
  }

  test("unknown future stop reasons are flagged rather than ignored") {
    assertEquals(
      StopReasonPolicy.errorFor(Some("some_future_reason")),
      Some("turn incomplete: stop_reason=some_future_reason")
    )
  }
