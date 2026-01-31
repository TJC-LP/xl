package com.tjclp.xl.agent.benchmark.common

import cats.effect.IO
import cats.syntax.all.*

/** Shared comparison table reporter for benchmarks */
object ComparisonReporter:

  /** Print comparison table with separate input/output columns */
  def printTable(
    title: String,
    results: List[TaskComparison],
    pricing: ModelPricing
  ): IO[Unit] =
    for
      _ <- BenchmarkUtils.printHeader(title, 120)
      _ <- printTableHeader()
      _ <- IO.println("-" * 120)

      _ <- results.traverse_(printRow)

      _ <- IO.println("-" * 120)
      _ <- printTotals(results)
      _ <- IO.println("=" * 120)
      _ <- printSummary(results, pricing)
    yield ()

  private def printTableHeader(): IO[Unit] =
    IO.println(
      f"${"Task"}%-15s | ${"xl In"}%9s | ${"xl Out"}%8s | ${"xlsx In"}%9s | ${"xlsx Out"}%9s | ${"In Diff"}%9s | ${"Out Diff"}%9s | ${"Pass"}%6s"
    )

  private def printRow(r: TaskComparison): IO[Unit] =
    val xlIn = r.xl.map(x => f"${x.usage.inputTokens}%,d").getOrElse("-")
    val xlOut = r.xl.map(x => f"${x.usage.outputTokens}%,d").getOrElse("-")
    val xlsxIn = r.xlsx.map(x => f"${x.usage.inputTokens}%,d").getOrElse("-")
    val xlsxOut = r.xlsx.map(x => f"${x.usage.outputTokens}%,d").getOrElse("-")

    val inDiff = r.inputTokenDiff
    val outDiff = r.outputTokenDiff
    val inDiffStr = formatDiff(inDiff)
    val outDiffStr = formatDiff(outDiff)

    val passStr = formatPassStatus(r.xl.map(_.passed), r.xlsx.map(_.passed))

    IO.println(
      f"${BenchmarkUtils.truncate(r.taskId, 15)}%-15s | $xlIn%9s | $xlOut%8s | $xlsxIn%9s | $xlsxOut%9s | $inDiffStr%9s | $outDiffStr%9s | $passStr%6s"
    )

  private def formatDiff(diff: Long): String =
    if diff == 0 then "0"
    else if diff < 0 then f"$diff%,d"
    else f"+$diff%,d"

  private def formatPassStatus(xl: Option[Boolean], xlsx: Option[Boolean]): String =
    (xl, xlsx) match
      case (Some(true), Some(true)) => "\u2713/\u2713"
      case (Some(true), Some(false)) => "\u2713/\u2717"
      case (Some(false), Some(true)) => "\u2717/\u2713"
      case (Some(false), Some(false)) => "\u2717/\u2717"
      case (Some(true), None) => "\u2713/-"
      case (Some(false), None) => "\u2717/-"
      case (None, Some(true)) => "-/\u2713"
      case (None, Some(false)) => "-/\u2717"
      case (None, None) => "-/-"

  private def printTotals(results: List[TaskComparison]): IO[Unit] =
    val xlInTotal = results.flatMap(_.xl).map(_.usage.inputTokens).sum
    val xlOutTotal = results.flatMap(_.xl).map(_.usage.outputTokens).sum
    val xlsxInTotal = results.flatMap(_.xlsx).map(_.usage.inputTokens).sum
    val xlsxOutTotal = results.flatMap(_.xlsx).map(_.usage.outputTokens).sum

    val inDiff = xlInTotal - xlsxInTotal
    val outDiff = xlOutTotal - xlsxOutTotal

    IO.println(
      f"${"TOTAL"}%-15s | $xlInTotal%,9d | $xlOutTotal%,8d | $xlsxInTotal%,9d | $xlsxOutTotal%,9d | ${formatDiff(inDiff)}%9s | ${formatDiff(outDiff)}%9s |"
    )

  private def printSummary(results: List[TaskComparison], pricing: ModelPricing): IO[Unit] =
    val stats = ComparisonStats.fromResults(results, pricing)

    val xlCacheHits = results.flatMap(_.xl).map(_.usage.cacheReadInputTokens).sum
    val xlsxCacheHits = results.flatMap(_.xlsx).map(_.usage.cacheReadInputTokens).sum

    for
      _ <- IO.println("")
      _ <- IO.println(
        f"Estimated cost: xl=$$${stats.xlEstimatedCost}%.4f, xlsx=$$${stats.xlsxEstimatedCost}%.4f"
      )
      _ <-
        if stats.xlEstimatedCost < stats.xlsxEstimatedCost then
          IO.println(f"  xl saves ${stats.costSavingsPercent}%.1f%% vs xlsx")
        else if stats.xlsxEstimatedCost < stats.xlEstimatedCost then
          IO.println(f"  xlsx saves ${-stats.costSavingsPercent}%.1f%% vs xl")
        else IO.println("  Both approaches have equal cost")
      _ <- IO.println(
        f"Pass rate: xl=${stats.xlPassed}/${stats.totalTasks}, xlsx=${stats.xlsxPassed}/${stats.totalTasks}"
      )
      _ <- IO.whenA(xlCacheHits > 0 || xlsxCacheHits > 0) {
        IO.println(f"Cache read tokens: xl=$xlCacheHits%,d, xlsx=$xlsxCacheHits%,d")
      }
    yield ()

  /** Print a simple single-approach summary (for non-comparison mode) */
  def printSingleApproachSummary(
    approach: String,
    results: List[ApproachResult],
    pricing: ModelPricing
  ): IO[Unit] =
    val totalIn = results.map(_.usage.inputTokens).sum
    val totalOut = results.map(_.usage.outputTokens).sum
    val totalCost = results.map(_.usage.estimatedCost(pricing)).sum
    val passed = results.count(_.passed)
    val total = results.size

    for
      _ <- IO.println("")
      _ <- IO.println(f"$approach approach summary:")
      _ <- IO.println(f"  Input tokens: $totalIn%,d")
      _ <- IO.println(f"  Output tokens: $totalOut%,d")
      _ <- IO.println(f"  Estimated cost: $$$totalCost%.4f")
      _ <- IO.println(f"  Pass rate: $passed/$total")
    yield ()
