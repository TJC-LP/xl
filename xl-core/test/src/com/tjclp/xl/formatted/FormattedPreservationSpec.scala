package com.tjclp.xl.formatted

import com.tjclp.xl.*
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.formatted.Formatted
import com.tjclp.xl.sheet.Sheet
import com.tjclp.xl.style.CellStyle
import com.tjclp.xl.style.numfmt.NumFmt
import munit.FunSuite

class FormattedPreservationSpec extends FunSuite:

  test("money literal preserves Currency format in batch put") {
    val sheet = Sheet("Test").getOrElse(fail("Should create sheet"))

    // Use batch put with money literal
    val updated = sheet.put(ref"A1" -> money"$$1,234.56")

    // Verify cell has the value
    val cell = updated(ref"A1")
    cell.value match
      case CellValue.Number(n) =>
        assertEquals(n, BigDecimal("1234.56"))
      case other =>
        fail(s"Expected Number, got: $other")

    // Verify cell has Currency format applied via style
    cell.styleId match
      case Some(styleId) =>
        val style = updated.styleRegistry.get(styleId).getOrElse {
          fail(s"Style $styleId not found in registry")
        }
        assertEquals(
          style.numFmt,
          NumFmt.Currency,
          "money literal should apply Currency format"
        )
      case None =>
        fail("Cell should have a style with Currency format")
  }

  test("date literal preserves Date format in batch put") {
    val sheet = Sheet("Test").getOrElse(fail("Should create sheet"))

    val updated = sheet.put(ref"A1" -> date"2025-11-10")

    val cell = updated(ref"A1")
    cell.value match
      case CellValue.DateTime(_) => () // Correct type
      case other => fail(s"Expected DateTime, got: $other")

    cell.styleId match
      case Some(styleId) =>
        val style = updated.styleRegistry.get(styleId).getOrElse {
          fail(s"Style $styleId not found")
        }
        assertEquals(style.numFmt, NumFmt.Date, "date literal should apply Date format")
      case None =>
        fail("Cell should have style with Date format")
  }

  test("percent literal preserves Percent format in batch put") {
    val sheet = Sheet("Test").getOrElse(fail("Should create sheet"))

    val updated = sheet.put(ref"A1" -> percent"15.5%")

    val cell = updated(ref"A1")
    cell.value match
      case CellValue.Number(n) =>
        // Percent is stored as decimal (0.155)
        assert(
          (n - BigDecimal("0.155")).abs < BigDecimal("0.0001"),
          s"Expected 0.155, got $n"
        )
      case other =>
        fail(s"Expected Number, got: $other")

    cell.styleId match
      case Some(styleId) =>
        val style = updated.styleRegistry.get(styleId).getOrElse {
          fail(s"Style $styleId not found")
        }
        assertEquals(style.numFmt, NumFmt.Percent, "percent literal should apply Percent format")
      case None =>
        fail("Cell should have style with Percent format")
  }

  test("mixed formatted and plain values in single batch put") {
    val sheet = Sheet("Test").getOrElse(fail("Should create sheet"))

    val updated = sheet.put(
      ref"A1" -> "Product",               // Plain String
      ref"B1" -> money"$$1,234.56",       // Formatted Currency
      ref"C1" -> percent"15%",            // Formatted Percent
      ref"D1" -> BigDecimal("999.99"),    // Plain BigDecimal (auto Decimal format)
      ref"E1" -> date"2025-11-10"         // Formatted Date
    )

    // Verify A1 (plain String) has no specific format
    val cellA1 = updated(ref"A1")
    assertEquals(cellA1.value, CellValue.Text("Product"))

    // Verify B1 (money) has Currency format
    val cellB1 = updated(ref"B1")
    cellB1.styleId.foreach { styleId =>
      val style = updated.styleRegistry.get(styleId).get
      assertEquals(style.numFmt, NumFmt.Currency)
    }

    // Verify C1 (percent) has Percent format
    val cellC1 = updated(ref"C1")
    cellC1.styleId.foreach { styleId =>
      val style = updated.styleRegistry.get(styleId).get
      assertEquals(style.numFmt, NumFmt.Percent)
    }

    // Verify D1 (BigDecimal) has auto-inferred Decimal format
    val cellD1 = updated(ref"D1")
    cellD1.styleId.foreach { styleId =>
      val style = updated.styleRegistry.get(styleId).get
      assertEquals(style.numFmt, NumFmt.Decimal)
    }

    // Verify E1 (date) has Date format
    val cellE1 = updated(ref"E1")
    cellE1.styleId.foreach { styleId =>
      val style = updated.styleRegistry.get(styleId).get
      assertEquals(style.numFmt, NumFmt.Date)
    }
  }

  test("accounting literal preserves Currency format") {
    val sheet = Sheet("Test").getOrElse(fail("Should create sheet"))

    // Test accounting format (negative shown as parentheses)
    val updated = sheet.put(ref"A1" -> accounting"$$(999.99)")

    // Verify value is negative
    val cell = updated(ref"A1")
    cell.value match
      case CellValue.Number(n) =>
        assertEquals(n, BigDecimal("-999.99"))
      case other =>
        fail(s"Expected Number, got: $other")

    // Verify Currency format applied (accounting uses Currency format)
    cell.styleId.foreach { styleId =>
      val style = updated.styleRegistry.get(styleId).get
      assertEquals(style.numFmt, NumFmt.Currency)
    }
  }
