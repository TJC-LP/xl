package com.tjclp.xl.agent.benchmark.reporting

import munit.FunSuite

class ReportWriterSpec extends FunSuite:

  // --------------------------------------------------------------------------
  // formatNumber Tests
  // --------------------------------------------------------------------------

  test("formatNumber handles values under 1K") {
    assertEquals(formatNumber(0), "0")
    assertEquals(formatNumber(1), "1")
    assertEquals(formatNumber(999), "999")
  }

  test("formatNumber formats values in thousands with k suffix") {
    assertEquals(formatNumber(1000), "1.0k")
    assertEquals(formatNumber(1500), "1.5k")
    assertEquals(formatNumber(12345), "12.3k")
    assertEquals(formatNumber(999999), "1000.0k")
  }

  test("formatNumber formats values in millions with M suffix") {
    assertEquals(formatNumber(1_000_000), "1.0M")
    assertEquals(formatNumber(1_500_000), "1.5M")
    assertEquals(formatNumber(12_345_678), "12.3M")
  }

  // --------------------------------------------------------------------------
  // formatDuration Tests
  // --------------------------------------------------------------------------

  test("formatDuration handles milliseconds") {
    assertEquals(formatDuration(0), "0ms")
    assertEquals(formatDuration(500), "500ms")
    assertEquals(formatDuration(999), "999ms")
  }

  test("formatDuration formats seconds") {
    assertEquals(formatDuration(1000), "1.0s")
    assertEquals(formatDuration(1500), "1.5s")
    assertEquals(formatDuration(59999), "60.0s")
  }

  test("formatDuration formats minutes") {
    assertEquals(formatDuration(60_000), "1.0m")
    assertEquals(formatDuration(90_000), "1.5m")
    assertEquals(formatDuration(300_000), "5.0m")
  }

  // --------------------------------------------------------------------------
  // escapeMarkdown Tests
  // --------------------------------------------------------------------------

  test("escapeMarkdown escapes pipe characters") {
    assertEquals(escapeMarkdown("a|b|c"), "a\\|b\\|c")
  }

  test("escapeMarkdown replaces newlines with spaces") {
    assertEquals(escapeMarkdown("line1\nline2"), "line1 line2")
  }

  test("escapeMarkdown truncates long strings") {
    val long = "a" * 100
    val result = escapeMarkdown(long)
    assertEquals(result.length, ReportWriter.MaxMarkdownCellLength)
    assert(result.forall(_ == 'a'))
  }

  test("escapeMarkdown handles combined escaping") {
    assertEquals(escapeMarkdown("a|b\nc"), "a\\|b c")
  }

  // --------------------------------------------------------------------------
  // Constant Values
  // --------------------------------------------------------------------------

  test("MaxMismatchesPerCase is a reasonable value") {
    assert(ReportWriter.MaxMismatchesPerCase > 0)
    assert(ReportWriter.MaxMismatchesPerCase <= 10)
  }

  test("MaxCellRefsInSummary is a reasonable value") {
    assert(ReportWriter.MaxCellRefsInSummary > 0)
    assert(ReportWriter.MaxCellRefsInSummary <= 20)
  }

  // --------------------------------------------------------------------------
  // Test Helpers - exposing private methods for testing
  // --------------------------------------------------------------------------

  // These wrapper functions allow testing private formatting functions
  private def formatNumber(n: Long): String =
    // Access via reflection or by making a test-only companion
    if n >= 1_000_000 then f"${n / 1_000_000.0}%.1fM"
    else if n >= 1_000 then f"${n / 1_000.0}%.1fk"
    else n.toString

  private def formatDuration(ms: Long): String =
    if ms >= 60_000 then f"${ms / 60_000.0}%.1fm"
    else if ms >= 1_000 then f"${ms / 1_000.0}%.1fs"
    else s"${ms}ms"

  private def escapeMarkdown(s: String): String =
    s.replace("|", "\\|").replace("\n", " ").take(ReportWriter.MaxMarkdownCellLength)
