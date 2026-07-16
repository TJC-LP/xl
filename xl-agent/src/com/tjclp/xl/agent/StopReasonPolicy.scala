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

  /** Wire value of the resumable pause stop: the server-side tool loop hit its iteration cap */
  val PauseTurn: String = "pause_turn"

  /**
   * Cap on automatic pause_turn resumes per logical turn (issue #344). Each resume re-sends the
   * request with the paused assistant turns appended (see `CodeExecution.buildParams`); the cap
   * keeps a stuck server loop from consuming budget forever. A turn still paused past the cap
   * surfaces through [[errorFor(stopReason:Option[String],resumesUsed:Int)*]].
   */
  val MaxPauseTurnResumes: Int = 3

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

  /**
   * Error text for a final stop reason after `resumesUsed` pause_turn auto-resumes (issue #344): a
   * pause_turn that survived the resume loop is flagged as exhausted, not merely incomplete. All
   * other reasons map exactly as [[errorFor(stopReason:Option[String])*]].
   */
  def errorFor(stopReason: Option[String], resumesUsed: Int): Option[String] =
    errorFor(stopReason).map { message =>
      if stopReason.contains(PauseTurn) && resumesUsed > 0 then
        s"$message (auto-resume exhausted after $resumesUsed resumes)"
      else message
    }
