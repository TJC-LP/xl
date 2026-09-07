package com.tjclp.xl.cli.contract

import munit.CatsEffectSuite

import com.tjclp.xl.cli.CliIO
import com.tjclp.xl.error.XLError

class DiagnosticsSpec extends CatsEffectSuite:

  test("render: today's `Error: <message>` first, then the code, nothing else") {
    assertEquals(
      Diagnostics.render(CliError("SHEET_NOT_FOUND", "Sheet not found: Nope. Available: Data")),
      "Error: Sheet not found: Nope. Available: Data\n  code: SHEET_NOT_FOUND"
    )
  }

  test("render: candidates then hint, indented, in that order") {
    val err = CliError
      .fromXLError(XLError.SheetRequired("view", Vector("Data", "Summary")), None)
      .copy(message = "view requires --sheet")
    assertEquals(
      Diagnostics.render(err),
      """Error: view requires --sheet
        |  code: SHEET_REQUIRED
        |  did you mean: Data, Summary
        |  hint: use -s <name> or a qualified ref like 'Name'!A1""".stripMargin
    )
  }

  test("render: a hint without candidates skips the did-you-mean line") {
    val err = CliError.fromXLError(XLError.SheetNotFound("Nope"), None)
    assertEquals(
      Diagnostics.render(err),
      "Error: Sheet not found: 'Nope'\n  code: SHEET_NOT_FOUND\n  hint: list sheets with `xl -f <file> sheets`"
    )
  }

  test("render: a multi-line message keeps the code line after the last message line") {
    val err = CliError("FORMULA_ERROR", "=SUM(\n    ^\nFormula error in '=SUM(': unexpected end")
    assertEquals(
      Diagnostics.render(err),
      "Error: =SUM(\n    ^\nFormula error in '=SUM(': unexpected end\n  code: FORMULA_ERROR"
    )
  }

  test("report writes the rendering to stderr and nothing to stdout") {
    val err = CliError(ErrorCode.OUTPUT_REQUIRED, "put requires -o <out.xlsx>")
    CliIO.capturing("").flatMap { (io, collect) =>
      Diagnostics.report(err, io) *> collect.map { (out, stderr) =>
        assertEquals(out, "")
        assertEquals(stderr, Diagnostics.render(err) + System.lineSeparator())
      }
    }
  }

  test("warnings render as Warning[<CODE>]: <message> on stderr") {
    val warning = Warning(WarningCode.READER_WARNING, "MissingStylesXml")
    assertEquals(Diagnostics.renderWarning(warning), "Warning[READER_WARNING]: MissingStylesXml")
    CliIO.capturing("").flatMap { (io, collect) =>
      Diagnostics.warn(warning, io) *> collect.map { (out, stderr) =>
        assertEquals(out, "")
        assertEquals(stderr, "Warning[READER_WARNING]: MissingStylesXml" + System.lineSeparator())
      }
    }
  }
