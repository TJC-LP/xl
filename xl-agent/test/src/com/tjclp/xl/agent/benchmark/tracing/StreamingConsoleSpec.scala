package com.tjclp.xl.agent.benchmark.tracing

import munit.FunSuite

/** Pins the live console completion line, including error surfacing on failure (issue #334) */
class StreamingConsoleSpec extends FunSuite:

  test("completion line shows PASSED without an error suffix") {
    val line = StreamingConsole.formatCompleteLine("13894", 3, passed = true, 9_500L, None)
    assert(line.contains("PASSED"), line)
    assert(line.contains("13894:3"), line)
    assert(line.contains("9.5s"), line)
  }

  test("completion line shows FAILED with the error message") {
    val line = StreamingConsole.formatCompleteLine(
      "13894",
      3,
      passed = false,
      570_000L,
      Some("RuntimeException: boom mid-run")
    )
    assert(line.contains("FAILED"), line)
    assert(line.contains("RuntimeException: boom mid-run"), line)
  }

  test("completion line keeps errors single-line and bounded") {
    val messy = "Y" * 500 + "\nsecond line"
    val line = StreamingConsole.formatCompleteLine("t", 1, passed = false, 100L, Some(messy))
    assert(!line.contains("\n"), "completion line must stay a single line")
    assert(!line.contains("Y" * 201), "long errors must be truncated")
  }
