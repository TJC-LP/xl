package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.Sheet
import java.nio.file.{Files, Path}

/**
 * Security tests for formula injection prevention.
 *
 * Formula injection occurs when untrusted data starting with `=`, `+`, `-`, or `@` is written to
 * Excel. These characters can cause Excel to interpret the value as a formula, potentially leading
 * to:
 *   - Data exfiltration via HYPERLINK
 *   - Command execution via DDE
 *   - Remote resource access
 *
 * XL provides `CellValue.escape()` and `WriterConfig.secure` to prevent these attacks.
 */
@SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
class FormulaInjectionSpec extends FunSuite:

  // ========== CellValue.escape() tests ==========

  test("escape() prefixes = with single quote") {
    assertEquals(CellValue.escape("=SUM(A1)"), "'=SUM(A1)")
  }

  test("escape() prefixes + with single quote") {
    assertEquals(CellValue.escape("+1234"), "'+1234")
  }

  test("escape() prefixes - with single quote") {
    assertEquals(CellValue.escape("-100"), "'-100")
  }

  test("escape() prefixes @ with single quote") {
    assertEquals(CellValue.escape("@import"), "'@import")
  }

  test("escape() does not modify normal text") {
    assertEquals(CellValue.escape("Hello World"), "Hello World")
    assertEquals(CellValue.escape("Price: $100"), "Price: $100")
    assertEquals(CellValue.escape("50%"), "50%")
  }

  test("escape() does not modify empty string") {
    assertEquals(CellValue.escape(""), "")
  }

  test("escape() is idempotent (already escaped text unchanged)") {
    assertEquals(CellValue.escape("'=SUM(A1)"), "'=SUM(A1)")
    assertEquals(CellValue.escape("'+100"), "'+100")
  }

  test("escape() handles text starting with single quote but not formula char") {
    assertEquals(CellValue.escape("'Hello"), "'Hello")
  }

  // ========== CellValue.unescape() tests ==========

  test("unescape() removes quote prefix from escaped formula chars") {
    assertEquals(CellValue.unescape("'=SUM(A1)"), "=SUM(A1)")
    assertEquals(CellValue.unescape("'+100"), "+100")
    assertEquals(CellValue.unescape("'-50"), "-50")
    assertEquals(CellValue.unescape("'@test"), "@test")
  }

  test("unescape() does not modify normal text") {
    assertEquals(CellValue.unescape("Hello"), "Hello")
    assertEquals(CellValue.unescape("Price"), "Price")
  }

  test("unescape() does not modify text starting with quote but no formula char") {
    assertEquals(CellValue.unescape("'Hello"), "'Hello")
  }

  test("unescape() handles empty string") {
    assertEquals(CellValue.unescape(""), "")
  }

  test("unescape() only unescapes if next char is formula char") {
    // ''=nested: starts with ', but second char is ' not a formula char
    // So unescape does nothing (correct behavior)
    assertEquals(CellValue.unescape("''=nested"), "''=nested")

    // Only unescapes when the quote is followed by =, +, -, @
    assertEquals(CellValue.unescape("'=test"), "=test")
  }

  test("escape and unescape are inverses") {
    val testCases = List("=SUM(A1)", "+100", "-50", "@import", "Hello")
    testCases.foreach { text =>
      assertEquals(CellValue.unescape(CellValue.escape(text)), text)
    }
  }

  // ========== CellValue.couldBeFormula() tests ==========

  test("couldBeFormula() returns true for formula prefix chars") {
    assert(CellValue.couldBeFormula("=A1"))
    assert(CellValue.couldBeFormula("+100"))
    assert(CellValue.couldBeFormula("-50"))
    assert(CellValue.couldBeFormula("@test"))
  }

  test("couldBeFormula() returns false for normal text") {
    assert(!CellValue.couldBeFormula("Hello"))
    assert(!CellValue.couldBeFormula("$100"))
    assert(!CellValue.couldBeFormula("50%"))
    assert(!CellValue.couldBeFormula("'=escaped"))
  }

  test("couldBeFormula() returns false for empty string") {
    assert(!CellValue.couldBeFormula(""))
  }

  // ========== WriterConfig.secure integration tests ==========

  test("WriterConfig.secure escapes dangerous text in output") {
    val sheet = Sheet("Test")
      .put("A1" -> "=SUM(A2:A10)")
      .put("A2" -> "+1234")
      .put("A3" -> "-dangerous")
      .put("A4" -> "@import")
      .put("A5" -> "Normal text")

    val wb = Workbook(sheet)
    val tempFile = Files.createTempFile("test-injection-", ".xlsx")

    try
      // Write with secure config (escapes formulas)
      XlsxWriter.writeWith(wb, tempFile, WriterConfig.secure) match
        case Left(err) => fail(s"Write failed: $err")
        case Right(()) => ()

      // Read back and verify escaping
      XlsxReader.read(tempFile) match
        case Left(err) => fail(s"Read failed: $err")
        case Right(readWb) =>
          val readSheet = readWb.sheets.head

          // Text values should be escaped (formula injection prevented)
          readSheet.cells.get(ref"A1").map(_.value) match
            case Some(CellValue.Text(t)) =>
              assertEquals(t, "'=SUM(A2:A10)", "= prefix should be escaped")
            case other => fail(s"Expected Text, got: $other")

          readSheet.cells.get(ref"A2").map(_.value) match
            case Some(CellValue.Text(t)) =>
              assertEquals(t, "'+1234", "+ prefix should be escaped")
            case other => fail(s"Expected Text, got: $other")

          readSheet.cells.get(ref"A3").map(_.value) match
            case Some(CellValue.Text(t)) =>
              assertEquals(t, "'-dangerous", "- prefix should be escaped")
            case other => fail(s"Expected Text, got: $other")

          readSheet.cells.get(ref"A4").map(_.value) match
            case Some(CellValue.Text(t)) =>
              assertEquals(t, "'@import", "@ prefix should be escaped")
            case other => fail(s"Expected Text, got: $other")

          // Normal text unchanged
          readSheet.cells.get(ref"A5").map(_.value) match
            case Some(CellValue.Text(t)) =>
              assertEquals(t, "Normal text", "Normal text should be unchanged")
            case other => fail(s"Expected Text, got: $other")

    finally Files.deleteIfExists(tempFile)
  }

  test("WriterConfig.default does not escape text") {
    val sheet = Sheet("Test")
      .put("A1" -> "=SUM(A2:A10)")
      .put("A2" -> "Normal text")

    val wb = Workbook(sheet)
    val tempFile = Files.createTempFile("test-no-escape-", ".xlsx")

    try
      // Write with default config (no escaping)
      XlsxWriter.writeWith(wb, tempFile, WriterConfig.default) match
        case Left(err) => fail(s"Write failed: $err")
        case Right(()) => ()

      // Read back and verify no escaping
      XlsxReader.read(tempFile) match
        case Left(err) => fail(s"Read failed: $err")
        case Right(readWb) =>
          val readSheet = readWb.sheets.head

          // Text should NOT be escaped with default config
          readSheet.cells.get(ref"A1").map(_.value) match
            case Some(CellValue.Text(t)) =>
              assertEquals(t, "=SUM(A2:A10)", "Should not be escaped with default config")
            case other => fail(s"Expected Text, got: $other")

    finally Files.deleteIfExists(tempFile)
  }

  test("Formula cells are NOT escaped (only text cells)") {
    val sheet = Sheet("Test")
      .put("A1" -> CellValue.Formula("=A2+A3"))
      .put("A2" -> 10)
      .put("A3" -> 20)

    val wb = Workbook(sheet)
    val tempFile = Files.createTempFile("test-formula-", ".xlsx")

    try
      // Write with secure config
      XlsxWriter.writeWith(wb, tempFile, WriterConfig.secure) match
        case Left(err) => fail(s"Write failed: $err")
        case Right(()) => ()

      // Read back and verify formula is preserved
      XlsxReader.read(tempFile) match
        case Left(err) => fail(s"Read failed: $err")
        case Right(readWb) =>
          val readSheet = readWb.sheets.head

          // Formula cells should NOT be escaped (they're actual formulas)
          readSheet.cells.get(ref"A1").map(_.value) match
            case Some(CellValue.Formula(expr, _)) =>
              assertEquals(expr, "=A2+A3", "Formula should not be escaped")
            case other => fail(s"Expected Formula, got: $other")

    finally Files.deleteIfExists(tempFile)
  }

  test("WriterConfig.secure does not double-escape already escaped text") {
    val sheet = Sheet("Test")
      .put("A1" -> "'=already escaped")

    val wb = Workbook(sheet)
    val tempFile = Files.createTempFile("test-double-escape-", ".xlsx")

    try
      XlsxWriter.writeWith(wb, tempFile, WriterConfig.secure) match
        case Left(err) => fail(s"Write failed: $err")
        case Right(()) => ()

      XlsxReader.read(tempFile) match
        case Left(err) => fail(s"Read failed: $err")
        case Right(readWb) =>
          val readSheet = readWb.sheets.head

          // Should not double-escape (idempotent)
          readSheet.cells.get(ref"A1").map(_.value) match
            case Some(CellValue.Text(t)) =>
              assertEquals(t, "'=already escaped", "Should not double-escape")
            case other => fail(s"Expected Text, got: $other")

    finally Files.deleteIfExists(tempFile)
  }

  test("FormulaInjectionPolicy enum values") {
    // Verify enum values exist and are distinct
    assert(FormulaInjectionPolicy.Escape != FormulaInjectionPolicy.None)

    // Verify WriterConfig.secure uses Escape
    assertEquals(WriterConfig.secure.formulaInjectionPolicy, FormulaInjectionPolicy.Escape)

    // Verify WriterConfig.default uses None (trust input)
    assertEquals(WriterConfig.default.formulaInjectionPolicy, FormulaInjectionPolicy.None)
  }
