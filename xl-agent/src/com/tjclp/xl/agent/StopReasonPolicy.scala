package com.tjclp.xl.agent

/**
 * Pure mapping from the API's final `stop_reason` to an agent-level error (issue #340).
 *
 * A sampling iteration that ends with `max_tokens` (or `refusal`) silently terminates the
 * code-execution loop and previously looked like a clean completion: the run xl/13894 case 3 hit
 * the 16K cap at 16,819 output tokens, produced no file, and recorded no error. Deriving
 * `AgentResult.error` from the stop reason makes every non-clean stop loud through the existing
 * failure plumbing (trace metadata, console lines, reports).
 */
object StopReasonPolicy:

  /** Stop reasons that mark a cleanly finished turn (wire values from `BetaStopReason`) */
  private val CleanStops: Set[String] = Set("end_turn", "tool_use", "stop_sequence")

  /**
   * Error text for a final stop reason; None when the turn finished cleanly.
   *
   * Unknown reasons are flagged rather than ignored: a stop reason this policy has never seen is by
   * definition not a known-clean completion, and a false alarm is cheaper than a silent truncation.
   */
  def errorFor(stopReason: Option[String]): Option[String] =
    stopReason.filterNot(CleanStops.contains).map {
      case "max_tokens" => "turn truncated: stop_reason=max_tokens (raise --max-tokens)"
      case "refusal" => "model refused: stop_reason=refusal"
      case other => s"turn incomplete: stop_reason=$other"
    }
