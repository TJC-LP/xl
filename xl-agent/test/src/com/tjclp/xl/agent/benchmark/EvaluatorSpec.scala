package com.tjclp.xl.agent.benchmark

import munit.FunSuite
import com.tjclp.xl.agent.error.AgentError

import java.nio.file.Path

/**
 * The grader reads `--json` envelopes only (GH-592): the first sheet is `sheets`' `data[0].name`,
 * the graded cells are `view --eval`'s `data.rows[].cells[]`, a failed envelope surfaces its
 * `error.code`, and the text the CLI prints without `--json` is a parse error — nothing here
 * scrapes markdown.
 */
class EvaluatorSpec extends FunSuite:

  private def envelope(
    data: String,
    error: String,
    ok: Boolean,
    exit: Int,
    verb: String = "sheets"
  ): String =
    s"""{
       |  "ok": $ok,
       |  "exitCode": $exit,
       |  "verb": "$verb",
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
    val nameless = envelope(
      """[{"index": 1, "state": "visible", "dimension": null}]""",
      "null",
      ok = true,
      exit = 0
    )
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

  test("a failed envelope without error.code is a parse error, not an invented code") {
    val codeless = envelope("null", """{"message": "something broke"}""", ok = false, exit = 3)
    Evaluator.envelopeData("sheets", codeless) match
      case Left(AgentError.ParseError(_, cause)) => assert(cause.contains("something broke"), cause)
      case other => fail(s"expected ParseError, got $other")
  }

  // --- view --json --eval: the graded cells ------------------------------------------------------

  private val viewData =
    """{
      |    "sheet": "Data",
      |    "range": "A1:C2",
      |    "rows": [
      |      {"row": 1, "cells": [
      |        {"ref": "A1", "type": "text", "value": "Revenue", "formatted": "Revenue"},
      |        {"ref": "B1", "type": "number", "value": 1000000, "formatted": "$1,000,000"},
      |        {"ref": "C1", "type": "formula", "formula": "=B1*2", "value": 2000000, "formatted": "$2,000,000"}
      |      ]},
      |      {"row": 2, "cells": [
      |        {"ref": "A2", "type": "boolean", "value": true, "formatted": "TRUE"},
      |        {"ref": "B2", "type": "error", "value": "#DIV/0!", "formatted": "#DIV/0!"},
      |        {"ref": "C2", "type": "empty", "value": null, "formatted": ""}
      |      ]}
      |    ]
      |  }""".stripMargin

  test("viewCells is data.rows[].cells[] of a successful envelope, normalized and keyed by ref") {
    val cells = Evaluator.viewCells(envelope(viewData, "null", ok = true, exit = 0, verb = "view"))
    assertEquals(
      cells,
      Right(
        Map(
          "A1" -> ComparableValue.Text("Revenue"),
          "B1" -> ComparableValue.Number(BigDecimal("1000000.00")),
          "C1" -> ComparableValue.Number(BigDecimal("2000000.00")),
          "A2" -> ComparableValue.Bool(true),
          "B2" -> ComparableValue.Error("#DIV/0!"),
          "C2" -> ComparableValue.Empty
        )
      )
    )
  }

  test("a failed view envelope surfaces error.code and the message, never an empty range") {
    val failed = envelope(
      "null",
      """{"code": "SHEET_NOT_FOUND", "message": "sheet 'Data' not found", "hint": null,
        |    "candidates": ["Sheet1"], "location": {"file": null, "sheet": "Data", "ref": null, "opIndex": null}}""".stripMargin,
      ok = false,
      exit = 3,
      verb = "view"
    )
    Evaluator.viewCells(failed) match
      case Left(AgentError.EvaluationFailed(cause)) =>
        assert(cause.contains("SHEET_NOT_FOUND"), cause)
        assert(cause.contains("'Data' not found"), cause)
      case other => fail(s"expected EvaluationFailed, got $other")
  }

  test("the text table (what `view` prints without --json) is a parse error") {
    val table =
      """|   | A       | B         |
        ||---|---------|-----------|
        || 1 | Revenue | 1,000,000 |""".stripMargin
    Evaluator.viewCells(table) match
      case Left(AgentError.ParseError(_, _)) => ()
      case other => fail(s"expected ParseError, got $other")
  }

  test("a view envelope whose data carries no rows is a parse error, not an empty range") {
    val rowless = envelope(
      """{"sheet": "Data", "range": "A1:C2"}""",
      "null",
      ok = true,
      exit = 0,
      verb = "view"
    )
    Evaluator.viewCells(rowless) match
      case Left(AgentError.ParseError(_, _)) => ()
      case other => fail(s"expected ParseError, got $other")
  }

  test("the grader asks for a range as `-f <file> -s <sheet> --json view <range> --eval`") {
    assertEquals(
      Evaluator.viewCommand("/opt/xl", Path.of("out", "book.xlsx"), "Data", "B2:D9"),
      List(
        "/opt/xl",
        "-f",
        Path.of("out", "book.xlsx").toString,
        "-s",
        "Data",
        "--json",
        "view",
        "B2:D9",
        "--eval"
      )
    )
  }
