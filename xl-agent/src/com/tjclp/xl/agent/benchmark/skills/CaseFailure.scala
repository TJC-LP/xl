package com.tjclp.xl.agent.benchmark.skills

import cats.effect.{Clock, IO}
import com.tjclp.xl.agent.TokenUsage as AgentTokenUsage
import com.tjclp.xl.agent.benchmark.execution.{CaseDetails, CaseResult, TokenUsage}
import com.tjclp.xl.agent.benchmark.tracing.ConversationTracer

/**
 * Total failure handlers for benchmark case execution (issue #334).
 *
 * When a case dies mid-run the conversation trace collected so far is the primary diagnostic, so
 * these handlers save it best-effort before returning the failing CaseResult. They never raise:
 * tracer completion/save failures are swallowed (the trace is lost but the error is preserved).
 */
object CaseFailure:

  /** One-line description of a failure, total even for exceptions without a message */
  def describe(e: Throwable): String =
    val message = Option(e.getMessage).getOrElse("(no message)")
    s"${e.getClass.getSimpleName}: $message"

  /**
   * Build the failing CaseResult for a case that raised after its tracer was created.
   *
   * Completes the tracer with the error and saves the partial trace; a successfully saved path is
   * wired into the CaseResult so reports can point at the surviving trace.
   */
  def withPartialTrace(
    tracer: ConversationTracer,
    caseNum: Int,
    startTimeMs: Long,
    error: Throwable
  ): IO[CaseResult] =
    val message = describe(error)
    for
      endTimeMs <- Clock[IO].monotonic.map(_.toMillis)
      _ <- tracer.complete(AgentTokenUsage.zero, passed = false, error = Some(message)).attempt
      saved <- tracer.save().attempt
    yield CaseResult(
      caseNum = caseNum,
      passed = false,
      usage = TokenUsage.zero,
      latencyMs = math.max(0L, endTimeMs - startTimeMs),
      details = CaseDetails.NoDetails,
      tracePath = saved.toOption,
      error = Some(message)
    )

  /** Fallback for failures before a tracer exists (there is no trace to save) */
  def withoutTrace(caseNum: Int, error: Throwable): IO[CaseResult] =
    IO.pure(
      CaseResult(
        caseNum = caseNum,
        passed = false,
        usage = TokenUsage.zero,
        latencyMs = 0,
        details = CaseDetails.NoDetails,
        error = Some(describe(error))
      )
    )
