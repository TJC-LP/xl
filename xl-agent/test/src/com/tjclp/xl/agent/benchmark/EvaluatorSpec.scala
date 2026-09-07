package com.tjclp.xl.agent.benchmark

import munit.FunSuite
import com.tjclp.xl.agent.error.AgentError

import java.nio.file.Path

/**
 * The grader reads `xl --json sheets` (GH-592): the first sheet is the envelope's `data[0].name`, a
 * failed envelope surfaces its `error.code`, and the text table the CLI prints without `--json` is
 * a parse error — nothing here scrapes markdown.
 */
class EvaluatorSpec extends FunSuite:

  private def envelope(data: String, error: String, ok: Boolean, exit: Int): String =
    s"""{
       |  "ok": $ok,
       |  "exitCode": $exit,
       |  "verb": "sheets",
       |  "version": "0.20.0",
       |  "data": $data,
       |  "warnings": [],
       |  "error": $error
       |}""".stripMargin

  private val twoSheets = envelope(
    """[
      |    {"name": "Data", "index": 1, "state": "visible", "dimension": "A1:C4"},
      |    {"name": "Summary", "index": 2, "state": "visible", "dimension": null}
      |  ]""".stripMargin,
    "null",
    ok = true,
    exit = 0
  )

  test("firstSheet is data[0].name of a successful envelope") {
    assertEquals(Evaluator.firstSheet(twoSheets), Right("Data"))
  }

  test("a failed envelope surfaces error.code and the message, never a guessed sheet") {
    val failed = envelope(
      "null",
      """{"code": "IO_READ", "message": "Failed to open file: nope.xlsx", "hint": null,
        |    "candidates": [], "location": {"file": "nope.xlsx", "sheet": null, "ref": null, "opIndex": null}}""".stripMargin,
      ok = false,
      exit = 3
    )
    Evaluator.firstSheet(failed) match
      case Left(AgentError.EvaluationFailed(cause)) =>
        assert(cause.contains("IO_READ"), cause)
        assert(cause.contains("nope.xlsx"), cause)
      case other => fail(s"expected EvaluationFailed, got $other")
  }

  test("a workbook with no sheets is an evaluation failure, not a default 'Sheet1'") {
    val none = envelope("[]", "null", ok = true, exit = 0)
    Evaluator.firstSheet(none) match
      case Left(AgentError.EvaluationFailed(cause)) => assert(cause.contains("no sheets"), cause)
      case other => fail(s"expected EvaluationFailed, got $other")
  }

  test("the text table (what `sheets` prints without --json) is a parse error") {
    val table =
      """| #   | Name    | Dimension | State |
        ||-----|---------|-----------|-------|
        || 1   | Data    | A1:C4     |       |
        || 2   | Summary | (unknown) |       |""".stripMargin
    Evaluator.firstSheet(table) match
      case Left(AgentError.ParseError(_, _)) => ()
      case other => fail(s"expected ParseError, got $other")
  }

  test("a sheet entry without a name is a parse error, not an empty sheet name") {
    val nameless =
      envelope("""[{"index": 1, "state": "visible", "dimension": null}]""", "null", true, 0)
    Evaluator.firstSheet(nameless) match
      case Left(AgentError.ParseError(_, _)) => ()
      case other => fail(s"expected ParseError, got $other")
  }

  test("the grader asks the binary for the envelope: `-f <file> --json sheets`") {
    assertEquals(
      Evaluator.sheetsCommand("/opt/xl", Path.of("out", "book.xlsx")),
      List("/opt/xl", "-f", Path.of("out", "book.xlsx").toString, "--json", "sheets")
    )
  }
