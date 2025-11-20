package com.tjclp.xl.formatted

import com.tjclp.xl.*
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formatted.Formatted
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.style.CellStyle
import com.tjclp.xl.style.numfmt.NumFmt
import com.tjclp.xl.unsafe.*
import munit.FunSuite

@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class FormattedPreservationSpec extends FunSuite:

  test("money literal preserves Currency format in batch put") {
    val sheet = Sheet("Test").getOrElse(fail("Should create sheet"))

    // Use batch put with money literal
    val updated = sheet.put(ref"A1" -> money"$$1,234.56").unsafe

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

    val updated = sheet.put(ref"A1" -> date"2025-11-10").unsafe

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

    val updated = sheet.put(ref"A1" -> percent"15.5%").unsafe

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
      .unsafe

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
    val updated = sheet.put(ref"A1" -> accounting"$$(999.99)").unsafe

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

  test("Formatted variables preserve format in batch put") {
    val sheet = Sheet("Test").getOrElse(fail("Should create sheet"))

    // Store formatted values in variables (not inline literals)
    val revenue = money"$$10,000.00"
    val growth = percent"15.5%"
    val reportDate = date"2025-11-14"

    // Pass variables to batch put
    val updated = sheet.put(
      ref"A1" -> "Revenue",
      ref"B1" -> revenue,     // Variable, not literal
      ref"A2" -> "Growth",
      ref"B2" -> growth,      // Variable, not literal
      ref"A3" -> "Date",
      ref"B3" -> reportDate   // Variable, not literal
    )
      .unsafe

    // Verify Currency format is preserved for money variable
    val cellB1 = updated(ref"B1")
    cellB1.styleId match
      case Some(styleId) =>
        val style = updated.styleRegistry.get(styleId).getOrElse {
          fail(s"Style $styleId not found")
        }
        assertEquals(style.numFmt, NumFmt.Currency, "money variable should preserve Currency format")
      case None =>
        fail("Cell B1 should have a style with Currency format")

    // Verify Percent format is preserved for percent variable
    val cellB2 = updated(ref"B2")
    cellB2.styleId match
      case Some(styleId) =>
        val style = updated.styleRegistry.get(styleId).getOrElse {
          fail(s"Style $styleId not found")
        }
        assertEquals(style.numFmt, NumFmt.Percent, "percent variable should preserve Percent format")
      case None =>
        fail("Cell B2 should have a style with Percent format")

    // Verify Date format is preserved for date variable
    val cellB3 = updated(ref"B3")
    cellB3.styleId match
      case Some(styleId) =>
        val style = updated.styleRegistry.get(styleId).getOrElse {
          fail(s"Style $styleId not found")
        }
        assertEquals(style.numFmt, NumFmt.Date, "date variable should preserve Date format")
      case None =>
        fail("Cell B3 should have a style with Date format")
  }
