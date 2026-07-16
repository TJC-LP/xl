package com.tjclp.xl.agent.benchmark.reporting

import cats.effect.IO
import munit.CatsEffectSuite
import com.tjclp.xl.agent.benchmark.execution.{
  BenchmarkRun,
  CaseResult,
  EngineConfig,
  ExecutionResult,
  SkillRunResult,
  TokenUsage as ExecTokenUsage
}
import com.tjclp.xl.agent.benchmark.execution.SkillSummary as ExecSkillSummary
import com.tjclp.xl.agent.benchmark.grading.{GradeDetails, Score}
import com.tjclp.xl.agent.benchmark.task.*

import java.nio.file.{Files, Path}
import java.time.Instant

/**
 * Pins the all-errored accounting decision (issue #344 item 8): errored tasks consumed budget and
 * represent real failures, so they COUNT in totals, usage sums, and the pass-rate denominator — but
 * are reported distinctly via an errored count, and never dilute the average score (which is over
 * graded, non-errored results only). Wave 11's task-level error roll-up previously dropped
 * all-errored tasks from the execution summary's total/passed/failed/usage sums entirely.
 */
class SkillSummarySpec extends CatsEffectSuite:

  private val tempDir = FunFixture[Path](
    setup = _ => Files.createTempDirectory("skill-summary-spec"),
    teardown = { dir =>
      import scala.jdk.CollectionConverters.*
      import scala.util.Using
      Using.resource(Files.walk(dir)) { stream =>
        stream.iterator().asScala.toList.reverse.foreach(p => Files.deleteIfExists(p))
      }
    }
  )

  // --------------------------------------------------------------------------
  // Reporting-layer SkillSummary (BenchmarkReport)
  // --------------------------------------------------------------------------

  private def entry(
    taskId: String,
    passed: Boolean,
    error: Option[String] = None,
    skipped: Boolean = false
  ): TaskResultEntry =
    TaskResultEntry(
      taskId = taskId,
      skill = "xl",
      caseNum = None,
      instruction = "",
      category = "",
      score = if passed then Score.BinaryScore.Pass else Score.BinaryScore.Fail,
      gradeDetails = if passed then GradeDetails.pass("ok") else GradeDetails.fail("nope"),
      usage = TokenSummary(100, 50),
      latencyMs = 1000,
      skipped = skipped,
      skipReason = Option.when(skipped)("skipped"),
      error = error
    )

  test("reporting summary counts errored tasks in totals but scores over graded only") {
    val results = List(
      entry("1", passed = true),
      entry("2", passed = true),
      entry("3", passed = false), // graded fail
      entry("4", passed = false, error = Some("Execution failed: boom")), // errored
      entry("5", passed = false, skipped = true) // skipped stays out of every sum
    )
    val summary = SkillSummary.fromResults("xl", "xl-cli", results)
    assertEquals(summary.taskCount, 4)
    assertEquals(summary.passCount, 2)
    assertEquals(summary.failCount, 2) // graded fail + errored
    assertEquals(summary.erroredCount, 1)
    // Average score over the 3 graded results only: (1.0 + 1.0 + 0.0) / 3
    assertEqualsDouble(summary.averageScore, 2.0 / 3.0, 1e-9)
    // Pass-rate denominator includes the errored task, which can never pass
    assertEqualsDouble(summary.passRate, 0.5, 1e-9)
    // Errored usage consumed budget and counts; skipped does not
    assertEquals(summary.totalTokens, TokenSummary(400, 200))
  }

  test("reporting summary of only errored tasks is all-errored, zero-scored, never passing") {
    val results = List(
      entry("1", passed = false, error = Some("Execution failed: a")),
      entry("2", passed = false, error = Some("Execution failed: b"))
    )
    val summary = SkillSummary.fromResults("xl", "xl-cli", results)
    assertEquals(summary.taskCount, 2)
    assertEquals(summary.passCount, 0)
    assertEquals(summary.failCount, 2)
    assertEquals(summary.erroredCount, 2)
    assertEqualsDouble(summary.averageScore, 0.0, 1e-9)
    assertEqualsDouble(summary.passRate, 0.0, 1e-9)
    assertEquals(summary.totalTokens, TokenSummary(200, 100))
  }

  // --------------------------------------------------------------------------
  // Execution-layer SkillSummary (BenchmarkEngine)
  // --------------------------------------------------------------------------

  private def passedResult(taskId: String): ExecutionResult =
    ExecutionResult.fromCases(
      taskId = TaskId(taskId),
      skill = "xl",
      caseResults = Vector(CaseResult.passed(1, ExecTokenUsage(100, 50), latencyMs = 10)),
      usage = ExecTokenUsage(100, 50),
      latencyMs = 10
    )

  private def failedResult(taskId: String): ExecutionResult =
    ExecutionResult.fromCases(
      taskId = TaskId(taskId),
      skill = "xl",
      caseResults = Vector(CaseResult.failed(1, ExecTokenUsage(100, 50), latencyMs = 10)),
      usage = ExecTokenUsage(100, 50),
      latencyMs = 10
    )

  private def erroredResult(taskId: String): ExecutionResult =
    ExecutionResult.failed(
      taskId = TaskId(taskId),
      skill = "xl",
      error = "Execution failed: skill exploded",
      usage = ExecTokenUsage(100, 50),
      latencyMs = 10
    )

  test("execution summary keeps errored results in total, failed, and usage sums") {
    val summary = ExecSkillSummary.fromResults(
      Vector(passedResult("1"), failedResult("2"), erroredResult("3"))
    )
    assertEquals(summary.total, 3)
    assertEquals(summary.passed, 1)
    assertEquals(summary.failed, 2) // graded fail + errored
    assertEquals(summary.errored, 1)
    assertEquals(summary.totalUsage, ExecTokenUsage(300, 150))
    assertEqualsDouble(summary.passRate, 1.0 / 3.0, 1e-9)
  }

  test("execution summary of an all-errored run no longer collapses to zero totals") {
    val summary = ExecSkillSummary.fromResults(Vector(erroredResult("1"), erroredResult("2")))
    assertEquals(summary.total, 2)
    assertEquals(summary.passed, 0)
    assertEquals(summary.failed, 2)
    assertEquals(summary.errored, 2)
    assertEquals(summary.totalUsage, ExecTokenUsage(200, 100))
    assertEqualsDouble(summary.passRate, 0.0, 1e-9)
  }

  // --------------------------------------------------------------------------
  // summary.json rendering (how the accounting is consumed)
  // --------------------------------------------------------------------------

  tempDir.test("summary.json skill entries carry the errored count") { dir =>
    val results = Vector(passedResult("1"), erroredResult("2"))
    val run = BenchmarkRun(
      startTime = Instant.parse("2026-07-15T00:00:00Z"),
      endTime = Instant.parse("2026-07-15T00:10:00Z"),
      config = EngineConfig.default,
      skillResults = Map(
        "xl" -> SkillRunResult(
          skill = "xl",
          displayName = "xl-cli",
          results = results,
          summary = ExecSkillSummary.fromResults(results)
        )
      ),
      tasks = Nil
    )
    UnifiedReportWriter.write(run, dir).flatMap { _ =>
      IO.blocking(Files.readString(dir.resolve("summary.json"))).map { raw =>
        val json = io.circe.parser.parse(raw).getOrElse(fail("unparseable summary.json"))
        val skill = json.hcursor
          .downField("skills")
          .focus
          .flatMap(_.asArray)
          .flatMap(_.headOption)
          .getOrElse(fail(s"no skills entry: $raw"))
        assertEquals(skill.hcursor.get[Int]("total"), Right(2))
        assertEquals(skill.hcursor.get[Int]("passed"), Right(1))
        assertEquals(skill.hcursor.get[Int]("failed"), Right(1))
        assertEquals(skill.hcursor.get[Int]("errored"), Right(1))
        // Errored usage stays in the token sums
        assertEquals(skill.hcursor.get[Long]("inputTokens"), Right(200L))
      }
    }
  }
