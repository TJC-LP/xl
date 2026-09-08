package com.tjclp.xl.error

import munit.FunSuite

/**
 * GH-589: `XLException` is the script-edge twin of the CLI's `CliError` — it carries the structured
 * [[XLError]] and, optionally, a richer message than the error renders on its own (the sync
 * facade's `readSheet` names the candidate sheets this way while keeping `error ==
 * SheetNotFound(name)` so `code`/`hint` stay the domain's).
 */
class XLExceptionSpec extends FunSuite:

  test("the one-argument form keeps the error's own message") {
    val ex = XLException(XLError.SheetNotFound("Sumary"))
    assertEquals(ex.getMessage, "Sheet not found: 'Sumary'")
    assertEquals(ex.error, XLError.SheetNotFound("Sumary"))
    assertEquals(ex.error.code, "SHEET_NOT_FOUND")
  }

  test("the detail form replaces the message and keeps the structured error") {
    val ex = XLException(
      XLError.SheetNotFound("Sumary"),
      "Sheet not found: 'Sumary'. Did you mean: Summary? Available: Data, Summary"
    )
    assertEquals(
      ex.getMessage,
      "Sheet not found: 'Sumary'. Did you mean: Summary? Available: Data, Summary"
    )
    assertEquals(ex.error, XLError.SheetNotFound("Sumary"))
    assertEquals(ex.error.hint, Some("list sheets with `xl -f <file> sheets`"))
  }
