package com.tjclp.xl.agent.benchmark.reporting

import cats.effect.IO
import cats.syntax.all.*
import com.tjclp.xl.agent.benchmark.execution.{BenchmarkRun, ExecutionResult, SkillRunResult}
import com.tjclp.xl.agent.benchmark.grading.Score
import io.circe.syntax.*

import java.nio.file.{Files, Path}
import java.time.Duration

// ============================================================================
// Report Writer
// ============================================================================

/** Writes benchmark reports in JSON and Markdown formats */
object ReportWriter:

  /** Write a report to the specified directory */
  def write(outputDir: Path, report: BenchmarkReport): IO[ReportPaths] =
    for
      _ <- IO.blocking(Files.createDirectories(outputDir))
      jsonPath = outputDir.resolve("report.json")
      mdPath = outputDir.resolve("report.md")
      _ <- writeJson(jsonPath, report)
      _ <- writeMarkdown(mdPath, report)
    yield ReportPaths(jsonPath, mdPath)

  /** Write JSON report */
  def writeJson(path: Path, report: BenchmarkReport): IO[Path] =
    IO.blocking {
      Files.writeString(path, report.toJsonPretty)
      path
    }

  /** Write Markdown report */
  def writeMarkdown(path: Path, report: BenchmarkReport): IO[Path] =
    IO.blocking {
      val content = formatMarkdown(report)
      Files.writeString(path, content)
      path
    }

  /** Format report as Markdown */
  def formatMarkdown(report: BenchmarkReport): String =
    val sb = new StringBuilder

    // Title and header
    sb.append(s"# ${report.header.title}\n\n")
    sb.append(s"**Benchmark Type:** ${report.header.benchmarkType}  \n")
    sb.append(s"**Timestamp:** ${report.header.formattedTimestamp}  \n")
    sb.append(s"**Duration:** ${report.header.formattedDuration}  \n")
    sb.append(s"**Model:** ${report.header.model}  \n")
    sb.append(s"**Skills:** ${report.header.skills.mkString(", ")}  \n\n")

    // Summary section
    sb.append("## Summary\n\n")
    sb.append(formatSummaryTable(report.summary))
    sb.append("\n")

    // Skill comparison (if multiple skills)
    if report.skillResults.size > 1 then
      sb.append("## Skill Comparison\n\n")
      sb.append(formatSkillComparisonTable(report.skillResults))
      sb.append("\n")

    // Per-skill results
    report.skillResults.foreach { case (skill, summary) =>
      sb.append(s"## Results: ${summary.displayName}\n\n")
      sb.append(formatSkillSummary(summary))
      sb.append("\n")

      val skillResults = report.taskResults.filter(_.skill == skill)
      if skillResults.nonEmpty then
        sb.append(formatResultsTable(skillResults))
        sb.append("\n")
    }

    // Failed tasks detail
    val failedTasks = report.taskResults.filter(r => !r.passed && !r.skipped)
    if failedTasks.nonEmpty then
      sb.append("## Failed Tasks Details\n\n")
      failedTasks.foreach { result =>
        sb.append(s"### Task ${result.taskId} (${result.skill})\n\n")
        result.gradeDetails.reasoning.foreach(r => sb.append(s"**Reason:** $r\n\n"))
        if result.gradeDetails.mismatches.nonEmpty then
          sb.append("**Mismatches:**\n\n")
          sb.append("| Cell | Expected | Actual |\n")
          sb.append("|------|----------|--------|\n")
          result.gradeDetails.mismatches.foreach { m =>
            sb.append(
              s"| ${m.ref} | ${escapeMarkdown(m.expected)} | ${escapeMarkdown(m.actual)} |\n"
            )
          }
          sb.append("\n")
      }

    // Metadata
    report.metadata.notes.foreach { notes =>
      sb.append("## Notes\n\n")
      sb.append(notes)
      sb.append("\n")
    }

    sb.toString()

  private def formatSummaryTable(summary: ReportSummary): String =
    val sb = new StringBuilder
    sb.append("| Metric | Value |\n")
    sb.append("|--------|-------|\n")
    sb.append(s"| Total Tasks | ${summary.totalTasks} |\n")
    sb.append(s"| Completed | ${summary.completedTasks} |\n")
    sb.append(s"| Skipped | ${summary.skippedTasks} |\n")
    sb.append(s"| Passed | ${summary.passedTasks} |\n")
    sb.append(s"| Failed | ${summary.failedTasks} |\n")
    sb.append(s"| Pass Rate | ${summary.passRatePercent}% |\n")
    sb.append(s"| Average Score | ${f"${summary.averageScore}%.2f"} |\n")
    sb.append(s"| Total Tokens | ${formatNumber(summary.totalTokens.total)} |\n")
    sb.append(s"| Avg Latency | ${formatDuration(summary.averageLatencyMs)} |\n")
    sb.toString()

  private def formatSkillComparisonTable(skills: Map[String, SkillSummary]): String =
    val sb = new StringBuilder
    sb.append("| Skill | Pass Rate | Avg Score | Total Tokens | Avg Latency |\n")
    sb.append("|-------|-----------|-----------|--------------|-------------|\n")
    skills.values.toList.sortBy(-_.passRate).foreach { s =>
      sb.append(s"| ${s.displayName} | ${s.passRatePercent}% | ${f"${s.averageScore}%.2f"} | ")
      sb.append(s"${formatNumber(s.totalTokens.total)} | ${formatDuration(s.averageLatencyMs)} |\n")
    }
    sb.toString()

  private def formatSkillSummary(summary: SkillSummary): String =
    s"""- **Tasks:** ${summary.taskCount}
       |- **Passed:** ${summary.passCount} (${summary.passRatePercent}%)
       |- **Failed:** ${summary.failCount}
       |- **Avg Score:** ${f"${summary.averageScore}%.2f"}
       |- **Total Tokens:** ${formatNumber(summary.totalTokens.total)}
       |- **Avg Latency:** ${formatDuration(summary.averageLatencyMs)}
       |""".stripMargin

  private def formatResultsTable(results: List[TaskResultEntry]): String =
    val sb = new StringBuilder
    sb.append("| Task ID | Category | Score | Passed | Tokens | Latency |\n")
    sb.append("|---------|----------|-------|--------|--------|----------|\n")
    results.foreach { r =>
      val status = if r.passed then "✓" else if r.skipped then "⊘" else "✗"
      sb.append(s"| ${r.taskId}${r.caseNum.map(c => s".$c").getOrElse("")} | ${r.category} | ")
      sb.append(s"${f"${r.normalizedScore}%.2f"} | $status | ")
      sb.append(s"${formatNumber(r.usage.total)} | ${formatDuration(r.latencyMs)} |\n")
    }
    sb.toString()

  private def formatNumber(n: Long): String =
    if n >= 1_000_000 then f"${n / 1_000_000.0}%.1fM"
    else if n >= 1_000 then f"${n / 1_000.0}%.1fk"
    else n.toString

  private def formatDuration(ms: Long): String =
    if ms >= 60_000 then f"${ms / 60_000.0}%.1fm"
    else if ms >= 1_000 then f"${ms / 1_000.0}%.1fs"
    else s"${ms}ms"

  private def escapeMarkdown(s: String): String =
    s.replace("|", "\\|").replace("\n", " ").take(50)

/** Paths to generated report files */
case class ReportPaths(
  json: Path,
  markdown: Path
)

// ============================================================================
// Streaming Report Writer
// ============================================================================

/** Writer for streaming results to console during execution */
object StreamingReportWriter:

  /** Format a single result for console output */
  def formatResult(result: TaskResultEntry): String =
    val status =
      if result.passed then Console.GREEN + "✓" + Console.RESET
      else if result.skipped then Console.YELLOW + "⊘" + Console.RESET
      else Console.RED + "✗" + Console.RESET

    val taskLabel = result.caseNum match
      case Some(c) => s"${result.taskId}.$c"
      case None => result.taskId

    f"$status ${result.skill}%-8s ${taskLabel}%-12s score=${result.normalizedScore}%.2f tokens=${result.usage.total}%6d latency=${result.latencyMs}%5dms"

  /** Format a progress update */
  def formatProgress(completed: Int, total: Int, skill: String): String =
    val pct = if total > 0 then (completed * 100) / total else 0
    s"[$completed/$total] ($pct%) - $skill"

  /** Write a streaming header */
  def writeHeader(skills: List[String], totalTasks: Int): IO[Unit] =
    IO.println(
      s"${Console.BOLD}Starting benchmark with ${skills.mkString(", ")} - $totalTasks tasks${Console.RESET}"
    )

  /** Write a streaming result */
  def writeResult(result: TaskResultEntry): IO[Unit] =
    IO.println(formatResult(result))

  /** Write a summary line */
  def writeSummary(summary: ReportSummary): IO[Unit] =
    IO.println(
      s"\n${Console.BOLD}Summary:${Console.RESET} ${summary.passedTasks}/${summary.completedTasks} passed (${summary.passRatePercent}%)"
    )

// ============================================================================
// Unified Report Writer for BenchmarkRun
// ============================================================================

/** Writer for BenchmarkRun results from BenchmarkEngine */
object UnifiedReportWriter:

  /**
   * Write a benchmark run to JSON and Markdown files.
   *
   * @param run
   *   The benchmark run to write
   * @param outputDir
   *   Output directory (usually run's outputDir)
   * @return
   *   Paths to written files
   */
  def write(run: BenchmarkRun, outputDir: Path): IO[ReportPaths] =
    for
      _ <- IO.blocking(Files.createDirectories(outputDir))
      jsonPath = outputDir.resolve("summary.json")
      mdPath = outputDir.resolve("summary.md")
      _ <- writeJson(jsonPath, run)
      _ <- writeMarkdown(mdPath, run)
      _ <- writePerSkillSummaries(run, outputDir)
    yield ReportPaths(jsonPath, mdPath)

  /** Write JSON summary */
  private def writeJson(path: Path, run: BenchmarkRun): IO[Unit] =
    IO.blocking {
      import io.circe.Json

      val skillsJson = run.skillResults.toList.map { case (skillName, sr) =>
        Json.obj(
          "name" -> Json.fromString(skillName),
          "displayName" -> Json.fromString(sr.displayName),
          "passed" -> Json.fromInt(sr.summary.passed),
          "failed" -> Json.fromInt(sr.summary.failed),
          "total" -> Json.fromInt(sr.summary.total),
          "passRate" -> Json.fromDoubleOrNull(sr.summary.passRate),
          "inputTokens" -> Json.fromLong(sr.summary.totalUsage.inputTokens),
          "outputTokens" -> Json.fromLong(sr.summary.totalUsage.outputTokens),
          "avgLatencyMs" -> Json.fromDoubleOrNull(sr.summary.avgLatencyMs),
          "estimatedCost" -> sr.summary.estimatedCost
            .map(c => Json.fromBigDecimal(c))
            .getOrElse(Json.Null)
        )
      }

      val resultsJson = run.allResults.map { r =>
        val gradeJson = r.gradeResult
          .map { gr =>
            Json.obj(
              "graderName" -> Json.fromString(gr.graderName),
              "grade" -> Json.fromString(gr.score.toString), // Letter grade (A, B, C, D, F)
              "score" -> Json.fromDoubleOrNull(gr.score.normalized),
              "passed" -> Json.fromBoolean(gr.passed),
              "reasoning" -> gr.details.reasoning.map(Json.fromString).getOrElse(Json.Null),
              "mismatches" -> (if gr.details.mismatches.isEmpty then Json.Null
                               else
                                 Json.arr(gr.details.mismatches.map { m =>
                                   Json.obj(
                                     "ref" -> Json.fromString(m.ref),
                                     "expected" -> Json.fromString(m.expected),
                                     "actual" -> Json.fromString(m.actual)
                                   )
                                 }*))
            )
          }
          .getOrElse(Json.Null)

        Json.obj(
          "taskId" -> Json.fromString(r.taskIdValue),
          "skill" -> Json.fromString(r.skill),
          "passed" -> Json.fromBoolean(r.passed),
          "score" -> Json.fromDoubleOrNull(r.aggregateScore.normalized),
          "passRate" -> Json.fromDoubleOrNull(r.passRate),
          "passedCases" -> Json.fromInt(r.passedCases),
          "totalCases" -> Json.fromInt(r.totalCases),
          "inputTokens" -> Json.fromLong(r.usage.inputTokens),
          "outputTokens" -> Json.fromLong(r.usage.outputTokens),
          "latencyMs" -> Json.fromLong(r.latencyMs),
          "gradeResult" -> gradeJson,
          "error" -> r.error.map(Json.fromString).getOrElse(Json.Null)
        )
      }

      val json = Json.obj(
        "startTime" -> Json.fromString(run.startTime.toString),
        "endTime" -> Json.fromString(run.endTime.toString),
        "durationMs" -> Json.fromLong(run.duration.toMillis),
        "totalTasks" -> Json.fromInt(run.tasks.length),
        "overallPassRate" -> Json.fromDoubleOrNull(run.overallPassRate),
        "totalInputTokens" -> Json.fromLong(run.totalUsage.inputTokens),
        "totalOutputTokens" -> Json.fromLong(run.totalUsage.outputTokens),
        "skills" -> Json.arr(skillsJson*),
        "results" -> Json.arr(resultsJson*)
      )

      Files.writeString(path, json.spaces2)
    }

  /** Write Markdown summary */
  private def writeMarkdown(path: Path, run: BenchmarkRun): IO[Unit] =
    IO.blocking {
      val sb = new StringBuilder

      // Header
      sb.append("# Benchmark Summary\n\n")
      sb.append(s"**Start:** ${run.startTime}  \n")
      sb.append(s"**Duration:** ${formatDuration(run.duration)}  \n")
      sb.append(s"**Tasks:** ${run.tasks.length}  \n")
      sb.append(s"**Overall Pass Rate:** ${(run.overallPassRate * 100).toInt}%  \n\n")

      // Skill comparison table
      sb.append("## Results by Skill\n\n")
      sb.append(
        "| Skill | Pass Rate | Passed | Failed | Input Tokens | Output Tokens | Avg Latency | Est. Cost |\n"
      )
      sb.append(
        "|-------|-----------|--------|--------|--------------|---------------|-------------|----------|\n"
      )

      run.skillResults.values.toList.sortBy(-_.summary.passRate).foreach { sr =>
        val cost = sr.summary.estimatedCost.map(c => f"$$$c%.2f").getOrElse("-")
        sb.append(
          s"| ${sr.displayName} | ${sr.summary.passRatePercent}% | ${sr.summary.passed} | ${sr.summary.failed} | "
        )
        sb.append(
          s"${formatNumber(sr.summary.totalUsage.inputTokens)} | ${formatNumber(sr.summary.totalUsage.outputTokens)} | "
        )
        sb.append(s"${formatDuration(sr.summary.avgLatencyMs.toLong)} | $cost |\n")
      }
      sb.append("\n")

      // Per-task breakdown (always shown)
      sb.append("## Per-Task Results\n\n")
      sb.append("| Task | Skill | Grade | Score | Tokens | Latency | Reasoning |\n")
      sb.append("|------|-------|-------|-------|--------|---------|----------|\n")
      run.allResults.sortBy(_.taskIdValue).foreach { r =>
        val grade = r.gradeResult.map(_.score.toString).getOrElse("-")
        val score = f"${r.aggregateScore.normalized}%.2f"
        val tokens = formatNumber(r.totalTokens)
        val latency = formatDuration(r.latencyMs)
        val reasoning = r.gradeResult
          .flatMap(_.details.reasoning)
          .map(s => truncateText(s, 40))
          .getOrElse("-")
        sb.append(
          s"| ${r.taskIdValue} | ${r.skill} | $grade | $score | $tokens | $latency | $reasoning |\n"
        )
      }
      sb.append("\n")

      // Per-task comparison (multi-skill only)
      if run.skillResults.size > 1 then
        sb.append("## Per-Task Comparison\n\n")
        sb.append("| Task ID | " + run.skillResults.keys.mkString(" | ") + " |\n")
        sb.append("|---------|" + run.skillResults.keys.map(_ => "-----").mkString("|") + "|\n")

        run.tasks.foreach { task =>
          val results = run.skillResults.map { case (skill, sr) =>
            sr.results
              .find(_.taskIdValue == task.taskIdValue)
              .map { r =>
                if r.passed then s"✓ ${(r.passRate * 100).toInt}%"
                else s"✗ ${(r.passRate * 100).toInt}%"
              }
              .getOrElse("-")
          }
          sb.append(s"| ${task.taskIdValue} | ${results.mkString(" | ")} |\n")
        }
        sb.append("\n")

      // Failed tasks
      val failed = run.allResults.filter(!_.passed)
      if failed.nonEmpty then
        sb.append("## Failed Tasks\n\n")
        failed.foreach { r =>
          sb.append(s"### ${r.taskIdValue} (${r.skill})\n\n")
          r.error.foreach(e => sb.append(s"**Error:** $e\n\n"))
          r.caseResults.filter(!_.passed).foreach { cr =>
            sb.append(
              s"- Case ${cr.caseNum}: ${cr.mismatches.take(3).map(m => s"${m.ref}").mkString(", ")}\n"
            )
          }
          sb.append("\n")
        }

      Files.writeString(path, sb.toString)
    }

  /** Write per-skill summary files */
  private def writePerSkillSummaries(run: BenchmarkRun, outputDir: Path): IO[Unit] =
    run.skillResults.toList.traverse_ { case (skill, sr) =>
      val skillDir = outputDir.resolve(skill)
      IO.blocking(Files.createDirectories(skillDir)) *>
        writeSkillSummary(skillDir.resolve("summary.json"), sr)
    }

  /** Write a single skill's summary */
  private def writeSkillSummary(path: Path, sr: SkillRunResult): IO[Unit] =
    IO.blocking {
      import io.circe.Json

      val resultsJson = sr.results.map { r =>
        val gradeJson = r.gradeResult
          .map { gr =>
            Json.obj(
              "graderName" -> Json.fromString(gr.graderName),
              "grade" -> Json.fromString(gr.score.toString), // Letter grade (A, B, C, D, F)
              "score" -> Json.fromDoubleOrNull(gr.score.normalized),
              "passed" -> Json.fromBoolean(gr.passed),
              "reasoning" -> gr.details.reasoning.map(Json.fromString).getOrElse(Json.Null)
            )
          }
          .getOrElse(Json.Null)

        Json.obj(
          "taskId" -> Json.fromString(r.taskIdValue),
          "passed" -> Json.fromBoolean(r.passed),
          "score" -> Json.fromDoubleOrNull(r.aggregateScore.normalized),
          "passedCases" -> Json.fromInt(r.passedCases),
          "totalCases" -> Json.fromInt(r.totalCases),
          "inputTokens" -> Json.fromLong(r.usage.inputTokens),
          "outputTokens" -> Json.fromLong(r.usage.outputTokens),
          "latencyMs" -> Json.fromLong(r.latencyMs),
          "gradeResult" -> gradeJson,
          "caseResults" -> Json.arr(r.caseResults.map { cr =>
            Json.obj(
              "caseNum" -> Json.fromInt(cr.caseNum),
              "passed" -> Json.fromBoolean(cr.passed),
              "latencyMs" -> Json.fromLong(cr.latencyMs)
            )
          }*)
        )
      }

      val json = Json.obj(
        "skill" -> Json.fromString(sr.skill),
        "displayName" -> Json.fromString(sr.displayName),
        "summary" -> Json.obj(
          "total" -> Json.fromInt(sr.summary.total),
          "passed" -> Json.fromInt(sr.summary.passed),
          "failed" -> Json.fromInt(sr.summary.failed),
          "passRate" -> Json.fromDoubleOrNull(sr.summary.passRate),
          "avgLatencyMs" -> Json.fromDoubleOrNull(sr.summary.avgLatencyMs),
          "totalInputTokens" -> Json.fromLong(sr.summary.totalUsage.inputTokens),
          "totalOutputTokens" -> Json.fromLong(sr.summary.totalUsage.outputTokens)
        ),
        "results" -> Json.arr(resultsJson*)
      )

      Files.writeString(path, json.spaces2)
    }

  private def formatNumber(n: Long): String =
    if n >= 1_000_000 then f"${n / 1_000_000.0}%.1fM"
    else if n >= 1_000 then f"${n / 1_000.0}%.1fk"
    else n.toString

  private def formatDuration(duration: Duration): String =
    formatDuration(duration.toMillis)

  private def formatDuration(ms: Long): String =
    if ms >= 60_000 then f"${ms / 60_000.0}%.1fm"
    else if ms >= 1_000 then f"${ms / 1_000.0}%.1fs"
    else s"${ms}ms"

  private def truncateText(s: String, maxLen: Int): String =
    val clean = s.replace("|", "\\|").replace("\n", " ").trim
    if clean.length <= maxLen then clean
    else clean.take(maxLen) + "..."
