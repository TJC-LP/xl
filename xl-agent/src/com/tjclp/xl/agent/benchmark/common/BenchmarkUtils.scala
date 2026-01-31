package com.tjclp.xl.agent.benchmark.common

import cats.effect.IO
import com.tjclp.xl.agent.TokenUsage

import java.nio.file.{Files, Path}
import java.time.{Instant, ZoneId}
import java.time.format.DateTimeFormatter

/** Shared utility functions for benchmarks */
object BenchmarkUtils:

  /** Format current timestamp for filenames (e.g., 2025-01-28T21-30-45) */
  def formatTimestamp: String =
    DateTimeFormatter
      .ofPattern("yyyy-MM-dd'T'HH-mm-ss")
      .format(Instant.now.atZone(ZoneId.systemDefault()))

  /** Format current timestamp for display (e.g., 2025-01-28 21:30:45) */
  def formatTimestampDisplay: String =
    DateTimeFormatter
      .ofPattern("yyyy-MM-dd HH:mm:ss")
      .format(Instant.now.atZone(ZoneId.systemDefault()))

  /** Aggregate token usage from multiple results */
  def aggregateUsage(usages: List[TokenUsage]): TokenUsage =
    usages.foldLeft(TokenUsage.zero)(_ + _)

  /** Calculate percentage savings (positive = better) */
  def calculateSavings(baseline: Long, actual: Long): Double =
    if baseline == 0 then 0.0
    else (baseline - actual).toDouble / baseline * 100

  /** Format a percentage for display */
  def formatPercent(value: Double): String =
    if value >= 0 then f"-$value%.1f%%"
    else f"+${-value}%.1f%%"

  /** Format token count with thousands separator */
  def formatTokens(count: Long): String =
    f"$count%,d"

  /** Format milliseconds as duration string */
  def formatDuration(ms: Long): String =
    if ms < 1000 then s"${ms}ms"
    else if ms < 60000 then f"${ms / 1000.0}%.1fs"
    else f"${ms / 60000}%dm ${(ms % 60000) / 1000}%ds"

  /** Truncate a string with ellipsis if too long */
  def truncate(s: String, maxLen: Int): String =
    if s.length <= maxLen then s
    else s.take(maxLen - 3) + "..."

  /** Print a horizontal rule */
  def printRule(char: Char = '=', width: Int = 80): IO[Unit] =
    IO.println(char.toString * width)

  /** Print a header with rules */
  def printHeader(title: String, width: Int = 80): IO[Unit] =
    for
      _ <- printRule('=', width)
      _ <- IO.println(title)
      _ <- printRule('=', width)
    yield ()

  /** Print a section separator */
  def printSeparator(width: Int = 80): IO[Unit] =
    printRule('-', width)

  /** Write string content to a file */
  def writeFile(path: Path, content: String): IO[Unit] =
    IO.blocking {
      Files.createDirectories(path.getParent)
      Files.writeString(path, content)
    }

  /** Read file content as string */
  def readFile(path: Path): IO[String] =
    IO.blocking(Files.readString(path))
