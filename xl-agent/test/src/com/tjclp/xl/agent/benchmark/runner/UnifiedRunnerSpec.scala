package com.tjclp.xl.agent.benchmark.runner

import munit.CatsEffectSuite
import com.tjclp.xl.agent.benchmark.execution.{
  BenchmarkRun,
  CaseDetails,
  CaseResult,
  EngineConfig,
  ExecutionResult,
  SkillRunResult,
  TokenUsage
}
import com.tjclp.xl.agent.benchmark.execution.SkillSummary as ExecSkillSummary
import com.tjclp.xl.agent.benchmark.task.*

import java.time.Instant

/** Pins error surfacing in the streaming console line and legacy report entries (issue #334) */
class UnifiedRunnerSpec extends CatsEffectSuite:

  private val caseError = "RuntimeException: boom mid-run"

  private def failedCase(caseNum: Int): CaseResult =
    CaseResult(
      caseNum = caseNum,
      passed = false,
      usage = TokenUsage.zero,
      latencyMs = 570_000L,
      details = CaseDetails.NoDetails,
      tracePath = None,
      error = Some(caseError)
    )

  test("streaming console line surfaces the first per-case error") {
    val result = ExecutionResult.fromCases(
      TaskId("13894"),
      "xl",
      Vector(CaseResult.passed(1, TokenUsage.zero, 1000), failedCase(3)),
      TokenUsage.zero,
      571_000L
    )
    val line = UnifiedRunner.formatExecutionResult(result)
    assert(line.contains(caseError), line)
  }

  test("streaming console line surfaces a task-level error") {
    val result = ExecutionResult.failed(TaskId("2768"), "xl", "Skill setup exploded")
    val line = UnifiedRunner.formatExecutionResult(result)
    assert(line.contains("Skill setup exploded"), line)
  }

  test("streaming console line omits the error suffix when there is none") {
    val result = ExecutionResult.fromCases(
      TaskId("2768"),
      "xl",
      Vector(CaseResult.passed(1, TokenUsage.zero, 1000)),
      TokenUsage.zero,
      1000L
    )
    val line = UnifiedRunner.formatExecutionResult(result)
    assert(!line.contains("error="), line)
  }

  test("streaming console line keeps errors single-line and bounded") {
    val messy = "X" * 500 + "\nsecond line"
    val result = ExecutionResult.failed(TaskId("2768"), "xl", messy)
    val line = UnifiedRunner.formatExecutionResult(result)
    assert(!line.contains("\n"), "console line must stay a single line")
    assert(!line.contains("X" * 201), "long errors must be truncated")
  }

  test("legacy report entries carry the per-case error") {
    val result = ExecutionResult.fromCases(
      TaskId("13894"),
      "xl",
      Vector(failedCase(3)),
      TokenUsage.zero,
      570_000L
    )
    val task = BenchmarkTask(
      id = TaskId("13894"),
      instruction = "Fill the summary table",
      category = TaskCategory.CellLevel,
      inputSource = InputSource.NoInput,
      evaluation = EvaluationSpec.fileOnly
    )
    val start = Instant.parse("2026-07-08T00:00:00Z")
    val run = BenchmarkRun(
      startTime = start,
      endTime = start.plusSeconds(600),
      config = EngineConfig.default,
      skillResults = Map(
        "xl" -> SkillRunResult(
          "xl",
          "xl-cli",
          Vector(result),
          ExecSkillSummary.fromResults(Vector(result))
        )
      ),
      tasks = List(task)
    )

    UnifiedRunner.buildReportFromRun(UnifiedConfig(), run).map { report =>
      val entry = report.taskResults
        .find(e => e.taskId == "13894" && e.caseNum.contains(3))
        .getOrElse(fail("missing entry for case 3"))
      assertEquals(entry.error, Some(caseError))
    }
  }
