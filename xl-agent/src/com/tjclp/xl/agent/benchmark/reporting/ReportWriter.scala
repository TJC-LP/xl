package com.tjclp.xl.agent.benchmark.reporting

import cats.effect.IO
import cats.syntax.all.*

import java.nio.file.{Files, Path}

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
