package com.tjclp.xl.error

import munit.FunSuite

/**
 * ADR-017 §2.7: every `XLError` case has a stable SCREAMING_SNAKE `code`, an optional one-line
 * `hint`, `candidates`, `root` and `opIndex`; `XLError.codes` enumerates the vocabulary.
 */
class XLErrorCodesSpec extends FunSuite:

  private val codePattern = "^[A-Z][A-Z0-9_]*$".r

  /**
   * One instance per case. `XLError.codes` is compared against its size so the table stays honest.
   */
  private val samples: Vector[XLError] = Vector(
    XLError.InvalidCellRef("Z0", "row 0"),
    XLError.InvalidRange("A1:", "no end"),
    XLError.InvalidReference("bare"),
    XLError.InvalidSheetName("a[b", "brackets"),
    XLError.OutOfBounds("XFE1", "column"),
    XLError.SheetNotFound("Nope"),
    XLError.DuplicateSheet("Data"),
    XLError.DuplicateCellRef("A1, A1"),
    XLError.InvalidColumn(-1, "negative"),
    XLError.InvalidRow(-1, "negative"),
    XLError.TypeMismatch("Int", "Text", "A1"),
    XLError.FormulaError("=SUM(", "unexpected end"),
    XLError.StyleError("bad"),
    XLError.NumberFormatError("0.0.0", "double point"),
    XLError.MoneyFormatError("$x", "nan"),
    XLError.PercentFormatError("x%", "nan"),
    XLError.DateFormatError("x", "nan"),
    XLError.AccountingFormatError("x", "nan"),
    XLError.ColorError("#GG", "hex"),
    XLError.InvalidWorkbook("no sheets"),
    XLError.IOError("disk"),
    XLError.ParseError("xl/workbook.xml", "unclosed"),
    XLError.SecurityError("too big"),
    XLError.ValueCountMismatch(4, 3, "A1:B2"),
    XLError.UnsupportedType("A1", "Unit"),
    XLError.InvalidTableName("t 1", "space"),
    XLError.InvalidTableDisplayName("t 1", "space"),
    XLError.InvalidTableRange("A1", "one cell"),
    XLError.InvalidTableColumns("dup"),
    XLError.Other("misc"),
    XLError.EditFailed(3, "put", XLError.SheetNotFound("x")),
    XLError.UnsupportedCapability("chart", "--stream", "omit --stream"),
    XLError.SheetRequired("view", Vector("Data", "Summary"))
  )

  test("the sample table covers every case (one code per case)") {
    assertEquals(samples.size, XLError.codes.size)
  }

  test("every case has a non-empty SCREAMING_SNAKE code") {
    samples.foreach { e =>
      assert(e.code.nonEmpty, s"empty code for $e")
      assert(codePattern.matches(e.code), s"code '${e.code}' for $e is not SCREAMING_SNAKE")
    }
    XLError.codes.foreach(c => assert(codePattern.matches(c), c))
  }

  test("XLError.codes is unique and contains every sample's code") {
    assertEquals(XLError.codes.distinct, XLError.codes)
    samples.foreach(e => assert(XLError.codes.contains(e.code), s"${e.code} missing from codes"))
  }

  test("codes are the SCREAMING_SNAKE of the case name") {
    assertEquals(XLError.SheetNotFound("x").code, "SHEET_NOT_FOUND")
    assertEquals(XLError.InvalidCellRef("x", "y").code, "INVALID_CELL_REF")
    assertEquals(XLError.IOError("x").code, "IO_ERROR")
    assertEquals(XLError.Other("x").code, "OTHER")
    assertEquals(XLError.SheetRequired("view", Vector.empty).code, "SHEET_REQUIRED")
    assertEquals(XLError.UnsupportedCapability("a", "b", "c").code, "UNSUPPORTED_CAPABILITY")
    assert(XLError.codes.contains("EDIT_FAILED"))
  }

  test("EditFailed reports its cause's code, root and candidates") {
    val cause = XLError.SheetRequired("put", Vector("Data"))
    val nested = XLError.EditFailed(3, "put", XLError.EditFailed(1, "style", cause))
    assertEquals(XLError.EditFailed(3, "put", XLError.SheetNotFound("x")).code, "SHEET_NOT_FOUND")
    assertEquals(nested.code, "SHEET_REQUIRED")
    assertEquals(nested.root, cause)
    assertEquals(nested.candidates, Vector("Data"))
    assertEquals(nested.opIndex, Some(3))
    assertEquals(cause.opIndex, None)
    assertEquals(cause.root, cause)
  }

  test("hints: the per-case table") {
    assertEquals(
      XLError.SheetNotFound("x").hint,
      Some("list sheets with `xl -f <file> sheets`")
    )
    assertEquals(
      XLError.SheetRequired("view", Vector.empty).hint,
      Some("use -s <name> or a qualified ref like 'Name'!A1")
    )
    assertEquals(
      XLError.ValueCountMismatch(4, 3, "A1:B2").hint,
      Some("provide exactly 4 values, or 1 to fill the range")
    )
    assertEquals(XLError.FormulaError("=", "x").hint, Some("check the formula with `xl eval`"))
    assertEquals(XLError.SecurityError("x").hint, Some("re-run with --max-size 0 or --stream"))
    assertEquals(XLError.OutOfBounds("x", "y").hint, Some("valid cells are A1..XFD1048576"))
    // On the case type itself `.hint` is the field; through XLError it is the extension.
    val capability: XLError = XLError.UnsupportedCapability("a", "b", "do c")
    assertEquals(capability.hint, Some("do c"))
    assertEquals(
      XLError.EditFailed(0, "put", XLError.SheetNotFound("x")).hint,
      XLError.SheetNotFound("x").hint
    )
    assertEquals(XLError.Other("x").hint, None)
    assertEquals(XLError.IOError("x").hint, None)
  }

  test("candidates: SheetRequired exposes the available names, others are empty") {
    assertEquals(
      XLError.SheetRequired("view", Vector("Data", "Summary")).candidates,
      Vector("Data", "Summary")
    )
    assertEquals(XLError.SheetNotFound("x").candidates, Vector.empty[String])
  }

  test("candidates: SheetNotFound offers the nearest of its available sheets (GH-589)") {
    val err = XLError.SheetNotFound("Sumary", Vector("Data", "Summary"))
    assertEquals(err.candidates, Vector("Summary"))
    assertEquals(err.message, "Sheet not found: 'Sumary'. Available: Data, Summary")
    assertEquals(XLError.SheetNotFound("Zebra", Vector("Data", "Summary")).candidates, Vector.empty)
    assertEquals(XLError.SheetNotFound("Zebra").message, "Sheet not found: 'Zebra'")
  }

  test(
    "renderDiagnostic: Error line, indented code, did-you-mean before hint, absent lines skipped"
  ) {
    assertEquals(XLError.Other("boom").renderDiagnostic, "Error: boom\n  code: OTHER")
    assertEquals(
      XLError.SheetRequired("view", Vector("Data", "Summary")).renderDiagnostic,
      "Error: view requires a sheet: pass -s <name> or qualify the ref (Sheet!A1). Available: Data, Summary\n" +
        "  code: SHEET_REQUIRED\n  did you mean: Data, Summary\n  hint: use -s <name> or a qualified ref like 'Name'!A1"
    )
    assertEquals(
      XLError.renderDiagnostic("USAGE", "unknown verb", None, Vector("view", "cell")),
      "Error: unknown verb\n  code: USAGE\n  did you mean: view, cell"
    )
  }

  test("messages of the new cases") {
    assertEquals(
      XLError.EditFailed(3, "put", XLError.SheetNotFound("x")).message,
      "op 3 (put): Sheet not found: 'x'"
    )
    assertEquals(
      XLError.UnsupportedCapability("chart", "--stream", "omit --stream").message,
      "chart requires --stream: omit --stream"
    )
    assertEquals(
      XLError.SheetRequired("view", Vector("Data", "Summary")).message,
      "view requires a sheet: pass -s <name> or qualify the ref (Sheet!A1). Available: Data, Summary"
    )
  }

  test("XLError derives CanEqual: cases compare without a cast") {
    assert(XLError.Other("a") == XLError.Other("a"))
    assert(XLError.Other("a") != XLError.Other("b"))
    val a: XLError = XLError.SheetNotFound("x")
    val b: XLError = XLError.SheetNotFound("x")
    assert(a == b)
  }
