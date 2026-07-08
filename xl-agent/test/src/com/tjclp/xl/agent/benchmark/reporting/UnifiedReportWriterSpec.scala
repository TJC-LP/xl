package com.tjclp.xl.agent.benchmark.reporting

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
import io.circe.Json

import java.nio.file.{Files, Path}
import java.time.Instant

/**
 * Pins per-case error diagnostics in report output (issue #334): a case that dies mid-run must
 * leave its error string in summary.json, the per-skill summary.json, and summary.md.
 */
class UnifiedReportWriterSpec extends CatsEffectSuite:

  private val tempDir = FunFixture[Path](
    setup = _ => Files.createTempDirectory("unified-report-spec"),
    teardown = { dir =>
      import scala.jdk.CollectionConverters.*
      import scala.util.Using
      Using.resource(Files.walk(dir)) { stream =>
        stream.iterator().asScala.toList.reverse.foreach(p => Files.deleteIfExists(p))
      }
    }
  )

  // Distinctive error longer than the 40-char reasoning truncation: the tail must survive
  // into the markdown so a failed 9-minute case is diagnosable from the report alone.
  private val caseError =
    "IllegalStateException: connection reset during sub-turn 20 (partial trace saved)"

  private val savedTracePath = Path.of("results", "tasks", "13894", "xl", "case3")

  private def sampleRun: BenchmarkRun =
    val okCase = CaseResult.passed(1, TokenUsage(100, 50), latencyMs = 1000)
    val failedCase = CaseResult(
      caseNum = 3,
      passed = false,
      usage = TokenUsage.zero,
      latencyMs = 570_000L,
      details = CaseDetails.NoDetails,
      tracePath = Some(savedTracePath),
      error = Some(caseError)
    )
    val result = ExecutionResult.fromCases(
      taskId = TaskId("13894"),
      skill = "xl",
      caseResults = Vector(okCase, failedCase),
      usage = TokenUsage(100, 50),
      latencyMs = 571_000L
    )
    val task = BenchmarkTask(
      id = TaskId("13894"),
      instruction = "Fill the summary table",
      category = TaskCategory.CellLevel,
      inputSource = InputSource.NoInput,
      evaluation = EvaluationSpec.fileOnly
    )
    val start = Instant.parse("2026-07-08T00:00:00Z")
    BenchmarkRun(
      startTime = start,
      endTime = start.plusSeconds(600),
      config = EngineConfig.default,
      skillResults = Map(
        "xl" -> SkillRunResult(
          skill = "xl",
          displayName = "xl-cli",
          results = Vector(result),
          summary = ExecSkillSummary.fromResults(Vector(result))
        )
      ),
      tasks = List(task)
    )

  private def parseJson(path: Path): Json =
    io.circe.parser
      .parse(Files.readString(path))
      .getOrElse(fail(s"unparseable JSON at $path"))

  private def caseEntries(resultJson: Json): Vector[Json] =
    resultJson.hcursor
      .downField("caseResults")
      .focus
      .flatMap(_.asArray)
      .getOrElse(fail(s"no caseResults array in $resultJson"))

  private def firstResult(json: Json): Json =
    json.hcursor
      .downField("results")
      .focus
      .flatMap(_.asArray)
      .flatMap(_.headOption)
      .getOrElse(fail(s"no results array in $json"))

  private def entryFor(entries: Vector[Json], caseNum: Int): Json =
    entries
      .find(_.hcursor.get[Int]("caseNum").contains(caseNum))
      .getOrElse(fail(s"no caseResults entry for case $caseNum"))

  tempDir.test("summary.json caseResults entries carry error and tracePath") { dir =>
    UnifiedReportWriter.write(sampleRun, dir).map { _ =>
      val json = parseJson(dir.resolve("summary.json"))
      val entries = caseEntries(firstResult(json))

      val failed = entryFor(entries, 3)
      assertEquals(failed.hcursor.get[String]("error"), Right(caseError))
      assertEquals(failed.hcursor.get[String]("tracePath"), Right(savedTracePath.toString))
      assertEquals(failed.hcursor.get[Long]("latencyMs"), Right(570_000L))

      val ok = entryFor(entries, 1)
      assert(ok.hcursor.downField("error").focus.exists(_.isNull), s"expected null error: $ok")
    }
  }

  tempDir.test("per-skill summary.json caseResults entries carry error") { dir =>
    UnifiedReportWriter.write(sampleRun, dir).map { _ =>
      val json = parseJson(dir.resolve("xl").resolve("summary.json"))
      val entries = caseEntries(firstResult(json))
      val failed = entryFor(entries, 3)
      assertEquals(failed.hcursor.get[String]("error"), Right(caseError))
    }
  }

  tempDir.test("summary.md records the full per-case error") { dir =>
    UnifiedReportWriter.write(sampleRun, dir).map { _ =>
      val md = Files.readString(dir.resolve("summary.md"))
      assert(md.contains(caseError), s"markdown must contain the case error:\n$md")
    }
  }
