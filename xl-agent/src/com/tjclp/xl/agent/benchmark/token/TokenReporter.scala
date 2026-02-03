package com.tjclp.xl.agent.benchmark.token

import cats.effect.IO
import cats.syntax.all.*
import com.tjclp.xl.agent.benchmark.common.BenchmarkUtils

import java.nio.file.{Files, Path}

/** Output formatting for the token efficiency benchmark */
object TokenReporter:

  /** Print the comparison table to console */
  def printComparisonTable(results: List[TokenTaskResult]): IO[Unit] =
    val hasGrades = results.exists(_.grade.isDefined)

    // Group results by task
    val taskResults = results.groupBy(_.taskId)

    for
      _ <- IO.println("")
      _ <- IO.println("=" * 120)
      _ <- IO.println("TOKEN EFFICIENCY COMPARISON: xl CLI vs Anthropic xlsx Skill")
      _ <- IO.println("=" * 120)

      // Print header
      _ <-
        if hasGrades then
          IO.println(
            f"${"Task"}%-20s | ${"xl Tokens"}%10s | ${"xl Grade"}%8s | ${"xlsx Tokens"}%11s | ${"xlsx Grade"}%10s | ${"Winner"}%8s | ${"Savings"}%8s"
          )
        else
          IO.println(
            f"${"Task"}%-20s | ${"xl Input"}%10s | ${"xl Output"}%10s | ${"xlsx Input"}%11s | ${"xlsx Output"}%12s | ${"Winner"}%8s | ${"Savings"}%8s"
          )

      _ <- IO.println("-" * 120)

      // Print each task row
      _ <- taskResults.toList.sortBy(_._1).traverse_ { case (taskId, taskRes) =>
        printTaskRow(taskRes, hasGrades)
      }

      _ <- IO.println("-" * 120)

      // Print totals
      _ <- printTotals(results, hasGrades)

      _ <- IO.println("=" * 120)

      // Print summary
      _ <- printSummary(results)
    yield ()

  /** Print a single task row */
  private def printTaskRow(taskResults: List[TokenTaskResult], hasGrades: Boolean): IO[Unit] =
    val xl = taskResults.find(_.approach == Approach.Xl)
    val xlsx = taskResults.find(_.approach == Approach.Xlsx)

    val taskName = taskResults.headOption.map(_.taskName).getOrElse("Unknown")
    val truncatedName = BenchmarkUtils.truncate(taskName, 20)

    (xl, xlsx) match
      case (Some(xlRes), Some(xlsxRes)) if xlRes.success && xlsxRes.success =>
        val winner =
          if xlRes.totalTokens < xlsxRes.totalTokens then "xl"
          else if xlsxRes.totalTokens < xlRes.totalTokens then "xlsx"
          else "tie"

        val savings =
          if xlRes.totalTokens < xlsxRes.totalTokens then
            BenchmarkUtils.calculateSavings(xlsxRes.totalTokens, xlRes.totalTokens)
          else if xlsxRes.totalTokens < xlRes.totalTokens then
            BenchmarkUtils.calculateSavings(xlRes.totalTokens, xlsxRes.totalTokens)
          else 0.0

        if hasGrades then
          val xlGrade = xlRes.grade.map(_.toString).getOrElse("-")
          val xlsxGrade = xlsxRes.grade.map(_.toString).getOrElse("-")
          IO.println(
            f"$truncatedName%-20s | ${xlRes.totalTokens}%,10d | $xlGrade%8s | ${xlsxRes.totalTokens}%,11d | $xlsxGrade%10s | $winner%8s | ${BenchmarkUtils.formatPercent(savings)}%8s"
          )
        else
          IO.println(
            f"$truncatedName%-20s | ${xlRes.inputTokens}%,10d | ${xlRes.outputTokens}%,10d | ${xlsxRes.inputTokens}%,11d | ${xlsxRes.outputTokens}%,12d | $winner%8s | ${BenchmarkUtils.formatPercent(savings)}%8s"
          )

      case _ =>
        val xlTok = xl.filter(_.success).map(_.totalTokens.toString).getOrElse("ERR")
        val xlsxTok = xlsx.filter(_.success).map(_.totalTokens.toString).getOrElse("ERR")
        if hasGrades then
          IO.println(
            f"$truncatedName%-20s | $xlTok%10s | ${"-"}%8s | $xlsxTok%11s | ${"-"}%10s | ${"N/A"}%8s | ${"N/A"}%8s"
          )
        else
          val xlIn = xl.filter(_.success).map(_.inputTokens.toString).getOrElse("ERR")
          val xlOut = xl.filter(_.success).map(_.outputTokens.toString).getOrElse("ERR")
          val xlsxIn = xlsx.filter(_.success).map(_.inputTokens.toString).getOrElse("ERR")
          val xlsxOut = xlsx.filter(_.success).map(_.outputTokens.toString).getOrElse("ERR")
          IO.println(
            f"$truncatedName%-20s | $xlIn%10s | $xlOut%10s | $xlsxIn%11s | $xlsxOut%12s | ${"N/A"}%8s | ${"N/A"}%8s"
          )

  /** Print totals row */
  private def printTotals(results: List[TokenTaskResult], hasGrades: Boolean): IO[Unit] =
    val xlStats = ApproachStats.fromResults(Approach.Xl, results)
    val xlsxStats = ApproachStats.fromResults(Approach.Xlsx, results)

    val winner =
      if xlStats.totalTokens < xlsxStats.totalTokens then "xl"
      else if xlsxStats.totalTokens < xlStats.totalTokens then "xlsx"
      else "tie"

    val savings =
      if xlStats.totalTokens < xlsxStats.totalTokens then
        BenchmarkUtils.calculateSavings(xlsxStats.totalTokens, xlStats.totalTokens)
      else if xlsxStats.totalTokens < xlStats.totalTokens then
        BenchmarkUtils.calculateSavings(xlStats.totalTokens, xlsxStats.totalTokens)
      else 0.0

    if hasGrades then
      val xlAvg = xlStats.avgGrade.map(_.toString).getOrElse("-")
      val xlsxAvg = xlsxStats.avgGrade.map(_.toString).getOrElse("-")
      IO.println(
        f"${"TOTAL"}%-20s | ${xlStats.totalTokens}%,10d | $xlAvg%8s | ${xlsxStats.totalTokens}%,11d | $xlsxAvg%10s | $winner%8s | ${BenchmarkUtils.formatPercent(savings)}%8s"
      )
    else
      IO.println(
        f"${"TOTAL"}%-20s | ${xlStats.totalInputTokens}%,10d | ${xlStats.totalOutputTokens}%,10d | ${xlsxStats.totalInputTokens}%,11d | ${xlsxStats.totalOutputTokens}%,12d | $winner%8s | ${BenchmarkUtils.formatPercent(savings)}%8s"
      )

  /** Print summary statistics */
  private def printSummary(results: List[TokenTaskResult]): IO[Unit] =
    val xlStats = ApproachStats.fromResults(Approach.Xl, results)
    val xlsxStats = ApproachStats.fromResults(Approach.Xlsx, results)

    // Count wins
    val taskResults = results.groupBy(_.taskId)
    var xlWins = 0
    var xlsxWins = 0

    taskResults.values.foreach { taskRes =>
      val xl = taskRes.find(r => r.approach == Approach.Xl && r.success)
      val xlsx = taskRes.find(r => r.approach == Approach.Xlsx && r.success)
      (xl, xlsx) match
        case (Some(xlRes), Some(xlsxRes)) =>
          if xlRes.totalTokens < xlsxRes.totalTokens then xlWins += 1
          else if xlsxRes.totalTokens < xlRes.totalTokens then xlsxWins += 1
        case _ => ()
    }

    for
      _ <- IO.println("")
      _ <- IO.println(s"Summary: xl wins $xlWins tasks, xlsx wins $xlsxWins tasks")
      _ <- IO.println(
        s"Total tokens: xl=${BenchmarkUtils.formatTokens(xlStats.totalTokens)}, xlsx=${BenchmarkUtils.formatTokens(xlsxStats.totalTokens)}"
      )

      _ <-
        if xlStats.totalTokens < xlsxStats.totalTokens then
          val savings = BenchmarkUtils.calculateSavings(xlsxStats.totalTokens, xlStats.totalTokens)
          IO.println(f"xl CLI uses $savings%.1f%% fewer tokens overall")
        else if xlsxStats.totalTokens < xlStats.totalTokens then
          val savings = BenchmarkUtils.calculateSavings(xlStats.totalTokens, xlsxStats.totalTokens)
          IO.println(f"Anthropic xlsx uses $savings%.1f%% fewer tokens overall")
        else IO.println("Both approaches use the same number of tokens")

      // Print grade distribution if available
      hasGrades = results.exists(_.grade.isDefined)
      _ <- IO.whenA(hasGrades) {
        for
          _ <- IO.println("")
          _ <- IO.println("Grade distribution:")
          _ <- IO.println(s"  xl: ${formatGradeDist(xlStats.grades)}")
          _ <- IO.println(s"  xlsx: ${formatGradeDist(xlsxStats.grades)}")
        yield ()
      }
    yield ()

  /** Format grade distribution */
  private def formatGradeDist(grades: List[Grade]): String =
    if grades.isEmpty then "-"
    else
      val counts = grades.groupBy(identity).view.mapValues(_.size).toMap
      List(Grade.A, Grade.B, Grade.C, Grade.D, Grade.F)
        .flatMap(g => counts.get(g).map(c => s"$g:$c"))
        .mkString(", ")

  /** Save JSON report to file */
  def saveJsonReport(path: Path, report: TokenBenchmarkReport): IO[Unit] =
    IO.blocking {
      Files.createDirectories(path.getParent)
      Files.writeString(path, report.toJsonPretty)
    } *> IO.println(s"\nResults saved to: $path")

  /** Build a benchmark report from results */
  def buildReport(
    results: List[TokenTaskResult],
    model: String,
    sampleFile: String
  ): TokenBenchmarkReport =
    TokenBenchmarkReport(
      timestamp = BenchmarkUtils.formatTimestamp,
      model = model,
      sampleFile = sampleFile,
      xlStats = ApproachStats.fromResults(Approach.Xl, results),
      xlsxStats = ApproachStats.fromResults(Approach.Xlsx, results),
      results = results
    )
