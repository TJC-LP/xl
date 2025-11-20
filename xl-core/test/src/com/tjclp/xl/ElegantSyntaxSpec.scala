package com.tjclp.xl

import com.tjclp.xl.*
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formatted.Formatted
import com.tjclp.xl.unsafe.*
import munit.FunSuite
import java.time.LocalDateTime
import com.tjclp.xl.style.numfmt.NumFmt

/** Tests for elegant CellValue syntax (given conversions, batch put, formatted literals) */
class ElegantSyntaxSpec extends FunSuite:

  def emptySheet: Sheet = Sheet(SheetName.unsafe("Test"))

  // ========== Given Conversions Tests ==========

  test("Given conversion: String → CellValue.Text") {
    import conversions.given

    val sheet = emptySheet
    val ref = ARef.from1(1, 1)

    val updated: Sheet = sheet.put(ref, "Hello")  // String auto-converts
    assertEquals(updated(ref).value, CellValue.Text("Hello"))
  }

  test("Given conversion: Int → CellValue.Number") {
    import conversions.given

    val sheet = emptySheet
    val ref = ARef.from1(1, 1)

    val updated = sheet.put(ref, 42)  // Int auto-converts
    assertEquals(updated(ref).value, CellValue.Number(BigDecimal(42)))
  }

  test("Given conversion: Double → CellValue.Number") {
    import conversions.given

    val sheet = emptySheet
    val ref = ARef.from1(1, 1)

    val updated = sheet.put(ref, 3.14159)  // Double auto-converts
    assertEquals(updated(ref).value, CellValue.Number(BigDecimal(3.14159)))
  }

  test("Given conversion: Boolean → CellValue.Bool") {
    import conversions.given

    val sheet = emptySheet
    val ref = ARef.from1(1, 1)

    val updated = sheet.put(ref, true)  // Boolean auto-converts
    assertEquals(updated(ref).value, CellValue.Bool(true))
  }

  test("Given conversion: chained puts with mixed types") {
    import conversions.given
    import com.tjclp.xl.macros.ref

    val sheet = emptySheet
      .put(ref"A1", "Product")
      .put(ref"B1", "Price")
      .put(ref"C1", "In Stock")
      .put(ref"A2", "Laptop")
      .put(ref"B2", 999.99)
      .put(ref"C2", true)

    assertEquals(sheet(ref"A1").value, CellValue.Text("Product"))
    assertEquals(sheet(ref"B2").value, CellValue.Number(BigDecimal(999.99)))
    assertEquals(sheet(ref"C2").value, CellValue.Bool(true))
  }

  // ========== Batch Put Macro Tests ==========

  test("Batch put: multiple cells at once") {
    import com.tjclp.xl.macros.ref

    val sheet = emptySheet
      .put(
        ref"A1" -> "Name",
        ref"B1" -> "Age",
        ref"C1" -> "Active"
      )
      .unsafe

    assertEquals(sheet(ref"A1").value, CellValue.Text("Name"))
    assertEquals(sheet(ref"B1").value, CellValue.Text("Age"))
    assertEquals(sheet(ref"C1").value, CellValue.Text("Active"))
  }

  test("Batch put: mixed types") {
    import com.tjclp.xl.macros.ref

    val sheet = emptySheet
      .put(
        ref"A1" -> "Product",
        ref"B1" -> 42,
        ref"C1" -> 3.14,
        ref"D1" -> true,
        ref"E1" -> BigDecimal(1000)
      )
      .unsafe

    assertEquals(sheet(ref"A1").value, CellValue.Text("Product"))
    assertEquals(sheet(ref"B1").value, CellValue.Number(BigDecimal(42)))
    assertEquals(sheet(ref"C1").value, CellValue.Number(BigDecimal(3.14)))
    assertEquals(sheet(ref"D1").value, CellValue.Bool(true))
    assertEquals(sheet(ref"E1").value, CellValue.Number(BigDecimal(1000)))
  }

  test("Batch put: creates table structure") {
    import com.tjclp.xl.macros.ref

    val sheet = emptySheet
      .put(
        // Headers
        ref"A1" -> "Item",
        ref"B1" -> "Qty",
        ref"C1" -> "Price",
        // Row 1
        ref"A2" -> "Laptop",
        ref"B2" -> 5,
        ref"C2" -> 999.99,
        // Row 2
        ref"A3" -> "Mouse",
        ref"B3" -> 25,
        ref"C3" -> 19.99
      )
      .unsafe

    assertEquals(sheet.cellCount, 9)
    assertEquals(sheet(ref"A2").value, CellValue.Text("Laptop"))
    assertEquals(sheet(ref"B2").value, CellValue.Number(BigDecimal(5)))
    assertEquals(sheet(ref"C3").value, CellValue.Number(BigDecimal(19.99)))
  }

  // ========== Formatted Literals Tests ==========

  test("money literal: parses $1,234.56") {
    import com.tjclp.xl.macros.money

    val formatted = money"$$1,234.56"  // $$ escapes the dollar sign

    assertEquals(formatted.value, CellValue.Number(BigDecimal("1234.56")))
    assertEquals(formatted.numFmt, NumFmt.Currency)
  }

  test("money literal: handles simple formats") {
    import com.tjclp.xl.macros.money

    val f1 = money"$$100"
    val f2 = money"1000.50"
    val f3 = money"$$1,000,000.99"

    assertEquals(f1.value, CellValue.Number(BigDecimal(100)))
    assertEquals(f2.value, CellValue.Number(BigDecimal("1000.50")))
    assertEquals(f3.value, CellValue.Number(BigDecimal("1000000.99")))
  }

  test("percent literal: parses 45.5%") {
    import com.tjclp.xl.macros.percent

    val formatted = percent"45.5%"

    assertEquals(formatted.value, CellValue.Number(BigDecimal("0.455")))
    assertEquals(formatted.numFmt, NumFmt.Percent)
  }

  test("percent literal: handles various formats") {
    import com.tjclp.xl.macros.percent

    val f1 = percent"100%"
    val f2 = percent"0.5%"
    val f3 = percent"75%"

    assertEquals(f1.value, CellValue.Number(BigDecimal("1.00")))
    assertEquals(f2.value, CellValue.Number(BigDecimal("0.005")))
    assertEquals(f3.value, CellValue.Number(BigDecimal("0.75")))
  }

  test("date literal: parses ISO dates") {
    import com.tjclp.xl.macros.date

    val formatted = date"2025-11-10"

    formatted.value match
      case CellValue.DateTime(dt) =>
        assertEquals(dt.toLocalDate.toString, "2025-11-10")
      case _ => fail("Expected DateTime")

    assertEquals(formatted.numFmt, NumFmt.Date)
  }

  test("accounting literal: parses positive and negative") {
    import com.tjclp.xl.macros.accounting

    val positive = accounting"$$123.45"
    val negative = accounting"($$123.45)"

    assertEquals(positive.value, CellValue.Number(BigDecimal("123.45")))
    assertEquals(negative.value, CellValue.Number(BigDecimal("-123.45")))
  }

  test("accounting literal: handles commas") {
    import com.tjclp.xl.macros.accounting

    val formatted = accounting"($$1,234.56)"

    assertEquals(formatted.value, CellValue.Number(BigDecimal("-1234.56")))
    assertEquals(formatted.numFmt, NumFmt.Currency)
  }

  // ========== Integration Tests ==========

  test("Formatted.putFormatted extension works") {
    import com.tjclp.xl.macros.ref
    import com.tjclp.xl.macros.money
    import Formatted.putFormatted

    val formatted = money"$$1,234.56"
    val sheet = emptySheet.putFormatted(ref"A1", formatted)

    assertEquals(sheet(ref"A1").value, CellValue.Number(BigDecimal("1234.56")))
  }

  test("Combined: batch put + formatted literals") {
    import com.tjclp.xl.macros.ref
    import com.tjclp.xl.macros.{money, percent}

    val sheet = emptySheet
      .put(
        ref"A1" -> "Revenue",
        ref"B1" -> money"$$10,000.00",    // Preserves Currency format
        ref"A2" -> "Growth",
        ref"B2" -> percent"15.5%"         // Preserves Percent format
      )
      .unsafe

    assertEquals(sheet(ref"B1").value, CellValue.Number(BigDecimal("10000.00")))
    assertEquals(sheet(ref"B2").value, CellValue.Number(BigDecimal("0.155")))

    // Verify formats preserved
    assert(
      sheet.getCellStyle(ref"B1").exists(_.numFmt == NumFmt.Currency),
      "money literal should preserve Currency format"
    )
    assert(
      sheet.getCellStyle(ref"B2").exists(_.numFmt == NumFmt.Percent),
      "percent literal should preserve Percent format"
    )
  }

  test("Real-world example: financial report") {
    import com.tjclp.xl.macros.ref
    import com.tjclp.xl.macros.{money, percent}

    val sheet = emptySheet
      .put(
        // Headers
        ref"A1" -> "Quarter",
        ref"B1" -> "Revenue",
        ref"C1" -> "Growth",
        // Q1
        ref"A2" -> "Q1 2025",
        ref"B2" -> money"$$125,000.00",   // Preserves Currency format
        ref"C2" -> percent"12.5%",        // Preserves Percent format
        // Q2
        ref"A3" -> "Q2 2025",
        ref"B3" -> money"$$150,000.00",
        ref"C3" -> percent"20.0%"
      )
      .unsafe

    assertEquals(sheet.cellCount, 9)
    assertEquals(sheet(ref"B2").value, CellValue.Number(BigDecimal("125000.00")))
    assertEquals(sheet(ref"C2").value, CellValue.Number(BigDecimal("0.125")))

    // Verify formats preserved
    assert(sheet.getCellStyle(ref"B2").exists(_.numFmt == NumFmt.Currency))
    assert(sheet.getCellStyle(ref"C2").exists(_.numFmt == NumFmt.Percent))
  }
