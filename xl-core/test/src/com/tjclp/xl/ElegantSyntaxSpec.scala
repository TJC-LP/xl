package com.tjclp.xl

import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.formatted.Formatted
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
    import com.tjclp.xl.macros.{cell, range}

    val sheet = emptySheet
      .put(cell"A1", "Product")
      .put(cell"B1", "Price")
      .put(cell"C1", "In Stock")
      .put(cell"A2", "Laptop")
      .put(cell"B2", 999.99)
      .put(cell"C2", true)

    assertEquals(sheet(cell"A1").value, CellValue.Text("Product"))
    assertEquals(sheet(cell"B2").value, CellValue.Number(BigDecimal(999.99)))
    assertEquals(sheet(cell"C2").value, CellValue.Bool(true))
  }

  // ========== Batch Put Macro Tests ==========

  test("Batch put: multiple cells at once") {
    import com.tjclp.xl.macros.{cell, range}
    import com.tjclp.xl.putMacro.put

    val sheet = emptySheet.put(
      cell"A1" -> "Name",
      cell"B1" -> "Age",
      cell"C1" -> "Active"
    )

    assertEquals(sheet(cell"A1").value, CellValue.Text("Name"))
    assertEquals(sheet(cell"B1").value, CellValue.Text("Age"))
    assertEquals(sheet(cell"C1").value, CellValue.Text("Active"))
  }

  test("Batch put: mixed types") {
    import com.tjclp.xl.macros.{cell, range}
    import com.tjclp.xl.putMacro.put

    val sheet = emptySheet.put(
      cell"A1" -> "Product",
      cell"B1" -> 42,
      cell"C1" -> 3.14,
      cell"D1" -> true,
      cell"E1" -> BigDecimal(1000)
    )

    assertEquals(sheet(cell"A1").value, CellValue.Text("Product"))
    assertEquals(sheet(cell"B1").value, CellValue.Number(BigDecimal(42)))
    assertEquals(sheet(cell"C1").value, CellValue.Number(BigDecimal(3.14)))
    assertEquals(sheet(cell"D1").value, CellValue.Bool(true))
    assertEquals(sheet(cell"E1").value, CellValue.Number(BigDecimal(1000)))
  }

  test("Batch put: creates table structure") {
    import com.tjclp.xl.macros.{cell, range}
    import com.tjclp.xl.putMacro.put

    val sheet = emptySheet.put(
      // Headers
      cell"A1" -> "Item",
      cell"B1" -> "Qty",
      cell"C1" -> "Price",
      // Row 1
      cell"A2" -> "Laptop",
      cell"B2" -> 5,
      cell"C2" -> 999.99,
      // Row 2
      cell"A3" -> "Mouse",
      cell"B3" -> 25,
      cell"C3" -> 19.99
    )

    assertEquals(sheet.cellCount, 9)
    assertEquals(sheet(cell"A2").value, CellValue.Text("Laptop"))
    assertEquals(sheet(cell"B2").value, CellValue.Number(BigDecimal(5)))
    assertEquals(sheet(cell"C3").value, CellValue.Number(BigDecimal(19.99)))
  }

  // ========== Formatted Literals Tests ==========

  test("money literal: parses $1,234.56") {
    import com.tjclp.xl.money

    val formatted = money"$$1,234.56"  // $$ escapes the dollar sign

    assertEquals(formatted.value, CellValue.Number(BigDecimal("1234.56")))
    assertEquals(formatted.numFmt, NumFmt.Currency)
  }

  test("money literal: handles simple formats") {
    import com.tjclp.xl.money

    val f1 = money"$$100"
    val f2 = money"1000.50"
    val f3 = money"$$1,000,000.99"

    assertEquals(f1.value, CellValue.Number(BigDecimal(100)))
    assertEquals(f2.value, CellValue.Number(BigDecimal("1000.50")))
    assertEquals(f3.value, CellValue.Number(BigDecimal("1000000.99")))
  }

  test("percent literal: parses 45.5%") {
    import com.tjclp.xl.percent

    val formatted = percent"45.5%"

    assertEquals(formatted.value, CellValue.Number(BigDecimal("0.455")))
    assertEquals(formatted.numFmt, NumFmt.Percent)
  }

  test("percent literal: handles various formats") {
    import com.tjclp.xl.percent

    val f1 = percent"100%"
    val f2 = percent"0.5%"
    val f3 = percent"75%"

    assertEquals(f1.value, CellValue.Number(BigDecimal("1.00")))
    assertEquals(f2.value, CellValue.Number(BigDecimal("0.005")))
    assertEquals(f3.value, CellValue.Number(BigDecimal("0.75")))
  }

  test("date literal: parses ISO dates") {
    import com.tjclp.xl.date

    val formatted = date"2025-11-10"

    formatted.value match
      case CellValue.DateTime(dt) =>
        assertEquals(dt.toLocalDate.toString, "2025-11-10")
      case _ => fail("Expected DateTime")

    assertEquals(formatted.numFmt, NumFmt.Date)
  }

  test("accounting literal: parses positive and negative") {
    import com.tjclp.xl.accounting

    val positive = accounting"$$123.45"
    val negative = accounting"($$123.45)"

    assertEquals(positive.value, CellValue.Number(BigDecimal("123.45")))
    assertEquals(negative.value, CellValue.Number(BigDecimal("-123.45")))
  }

  test("accounting literal: handles commas") {
    import com.tjclp.xl.accounting

    val formatted = accounting"($$1,234.56)"

    assertEquals(formatted.value, CellValue.Number(BigDecimal("-1234.56")))
    assertEquals(formatted.numFmt, NumFmt.Currency)
  }

  // ========== Integration Tests ==========

  test("Formatted.putFormatted extension works") {
    import com.tjclp.xl.macros.cell
    import com.tjclp.xl.money
    import Formatted.putFormatted

    val formatted = money"$$1,234.56"
    val sheet = emptySheet.putFormatted(cell"A1", formatted)

    assertEquals(sheet(cell"A1").value, CellValue.Number(BigDecimal("1234.56")))
  }

  test("Combined: batch put + formatted literals") {
    import com.tjclp.xl.macros.cell
    import com.tjclp.xl.{money, percent}
    import com.tjclp.xl.putMacro.put
    import Formatted.given  // Auto-conversion Formatted → CellValue

    val sheet = emptySheet.put(
      cell"A1" -> "Revenue",
      cell"B1" -> money"$$10,000.00".value,    // Extract value from Formatted
      cell"A2" -> "Growth",
      cell"B2" -> percent"15.5%".value
    )

    assertEquals(sheet(cell"B1").value, CellValue.Number(BigDecimal("10000.00")))
    assertEquals(sheet(cell"B2").value, CellValue.Number(BigDecimal("0.155")))
  }

  test("Real-world example: financial report") {
    import com.tjclp.xl.macros.cell
    import com.tjclp.xl.{money, percent}
    import com.tjclp.xl.putMacro.put

    val sheet = emptySheet.put(
      // Headers
      cell"A1" -> "Quarter",
      cell"B1" -> "Revenue",
      cell"C1" -> "Growth",
      // Q1
      cell"A2" -> "Q1 2025",
      cell"B2" -> money"$$125,000.00".value,
      cell"C2" -> percent"12.5%".value,
      // Q2
      cell"A3" -> "Q2 2025",
      cell"B3" -> money"$$150,000.00".value,
      cell"C3" -> percent"20.0%".value
    )

    assertEquals(sheet.cellCount, 9)
    assertEquals(sheet(cell"B2").value, CellValue.Number(BigDecimal("125000.00")))
    assertEquals(sheet(cell"C2").value, CellValue.Number(BigDecimal("0.125")))
  }
