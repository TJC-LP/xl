package com.tjclp.xl

import com.tjclp.xl.*
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.XLResult
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.unsafe.*
import munit.FunSuite

/** Tests for string-based map syntax with compile-time validation */
class StringMapSyntaxSpec extends FunSuite:

  def emptySheet: Sheet = Sheet(SheetName.unsafe("Test"))

  // ========== String-Based Batch Put Tests ==========

  test("put with string refs: compile-time validated literals return Sheet") {
    val sheet = emptySheet
      .put(
        "A1" -> "Name",
        "B1" -> "Age",
        "C1" -> "Active"
      )

    // Should return Sheet directly (not XLResult) for literals
    assertEquals(sheet(ARef.parse("A1").toOption.get).value, CellValue.Text("Name"))
    assertEquals(sheet(ARef.parse("B1").toOption.get).value, CellValue.Text("Age"))
    assertEquals(sheet(ARef.parse("C1").toOption.get).value, CellValue.Text("Active"))
  }

  test("put with string refs: mixed types") {
    val sheet = emptySheet
      .put(
        "A1" -> "Product",
        "B1" -> 42,
        "C1" -> 3.14,
        "D1" -> true,
        "E1" -> BigDecimal(1000)
      )

    assertEquals(sheet(ARef.parse("A1").toOption.get).value, CellValue.Text("Product"))
    assertEquals(sheet(ARef.parse("B1").toOption.get).value, CellValue.Number(BigDecimal(42)))
    assertEquals(sheet(ARef.parse("C1").toOption.get).value, CellValue.Number(BigDecimal(3.14)))
    assertEquals(sheet(ARef.parse("D1").toOption.get).value, CellValue.Bool(true))
    assertEquals(sheet(ARef.parse("E1").toOption.get).value, CellValue.Number(BigDecimal(1000)))
  }

  test("put with string refs: creates table structure") {
    val sheet = emptySheet
      .put(
        // Headers
        "A1" -> "Item",
        "B1" -> "Qty",
        "C1" -> "Price",
        // Row 1
        "A2" -> "Laptop",
        "B2" -> 5,
        "C2" -> 999.99,
        // Row 2
        "A3" -> "Mouse",
        "B3" -> 25,
        "C3" -> 19.99
      )

    assertEquals(sheet.cellCount, 9)
    assertEquals(sheet(ARef.parse("A2").toOption.get).value, CellValue.Text("Laptop"))
    assertEquals(sheet(ARef.parse("B2").toOption.get).value, CellValue.Number(BigDecimal(5)))
    assertEquals(sheet(ARef.parse("C3").toOption.get).value, CellValue.Number(BigDecimal(19.99)))
  }

  test("put with string refs: runtime strings return XLResult") {
    val col = "A"
    val row = "1"
    val ref = s"$col$row"

    val result: XLResult[Sheet] = emptySheet.put(ref -> "Dynamic")

    assert(result.isRight, "Runtime string should parse successfully")
    val sheet = result.toOption.get
    assertEquals(sheet(ARef.parse("A1").toOption.get).value, CellValue.Text("Dynamic"))
  }

  test("put with string refs: invalid runtime ref returns Left") {
    val badRef = "INVALID"

    val result: XLResult[Sheet] = emptySheet.put(badRef -> "Value")

    assert(result.isLeft, "Invalid ref should return Left")
  }

  // ========== String-Based Batch Style Tests ==========

  test("style with string refs: single cell literal returns Sheet") {
    val bold = CellStyle.default.bold
    val sheet = emptySheet
      .put("A1" -> "Header")
      .style("A1" -> bold)

    val style = sheet.getCellStyle(ARef.parse("A1").toOption.get)
    assert(style.exists(_.font.bold), "Cell should be bold")
  }

  test("style with string refs: range literal returns Sheet") {
    val bold = CellStyle.default.bold
    val sheet = emptySheet
      .put(
        "A1" -> "H1",
        "B1" -> "H2",
        "C1" -> "H3"
      )
      .style("A1:C1" -> bold)

    // All cells in range should be styled
    val a1Style = sheet.getCellStyle(ARef.parse("A1").toOption.get)
    val b1Style = sheet.getCellStyle(ARef.parse("B1").toOption.get)
    val c1Style = sheet.getCellStyle(ARef.parse("C1").toOption.get)

    assert(a1Style.exists(_.font.bold), "A1 should be bold")
    assert(b1Style.exists(_.font.bold), "B1 should be bold")
    assert(c1Style.exists(_.font.bold), "C1 should be bold")
  }

  test("style with string refs: multiple styles at once") {
    val bold = CellStyle.default.bold
    val currency = CellStyle.default.currency

    val sheet = emptySheet
      .put(
        "A1" -> "Revenue",
        "B1" -> 1000
      )
      .style(
        "A1" -> bold,
        "B1" -> currency
      )

    val a1Style = sheet.getCellStyle(ARef.parse("A1").toOption.get)
    val b1Style = sheet.getCellStyle(ARef.parse("B1").toOption.get)

    assert(a1Style.exists(_.font.bold), "A1 should be bold")
    assert(
      b1Style.exists(_.numFmt == com.tjclp.xl.styles.numfmt.NumFmt.Currency),
      "B1 should have currency format"
    )
  }

  test("style with string refs: mixed cell and range refs") {
    val bold = CellStyle.default.bold
    val percent = CellStyle.default.percent

    val sheet = emptySheet
      .put(
        "A1" -> "Headers",
        "A2" -> "Data1",
        "A3" -> "Data2",
        "B1" -> 0.5,
        "B2" -> 0.75,
        "B3" -> 0.25
      )
      .style(
        "A1:A3" -> bold,     // Range
        "B1"    -> percent   // Single cell
      )

    val a1Style = sheet.getCellStyle(ARef.parse("A1").toOption.get)
    val a2Style = sheet.getCellStyle(ARef.parse("A2").toOption.get)
    val b1Style = sheet.getCellStyle(ARef.parse("B1").toOption.get)

    assert(a1Style.exists(_.font.bold), "A1 should be bold")
    assert(a2Style.exists(_.font.bold), "A2 should be bold")
    assert(
      b1Style.exists(_.numFmt == com.tjclp.xl.styles.numfmt.NumFmt.Percent),
      "B1 should have percent format"
    )
  }

  test("style with string refs: runtime strings return XLResult") {
    val col = "A"
    val ref = s"${col}1"
    val bold = CellStyle.default.bold

    val result: XLResult[Sheet] = emptySheet
      .put("A1" -> "Test")
      .style(ref -> bold)

    assert(result.isRight, "Runtime string should parse successfully")
  }

  test("style with string refs: invalid runtime ref returns Left") {
    val badRef = "INVALID"
    val bold = CellStyle.default.bold

    val result: XLResult[Sheet] = emptySheet.style(badRef -> bold)

    assert(result.isLeft, "Invalid ref should return Left")
  }

  // ========== Integration Tests ==========

  test("combined put and style with string refs") {
    val bold = CellStyle.default.bold
    val currency = CellStyle.default.currency
    val percent = CellStyle.default.percent

    val sheet = emptySheet
      .put(
        "A1" -> "Revenue",
        "B1" -> 1250000,
        "A2" -> "Margin",
        "B2" -> 0.125
      )
      .style(
        "A1:A2" -> bold,
        "B1"    -> currency,
        "B2"    -> percent
      )

    assertEquals(sheet.cellCount, 4)
    assertEquals(sheet(ARef.parse("A1").toOption.get).value, CellValue.Text("Revenue"))
    assertEquals(sheet(ARef.parse("B1").toOption.get).value, CellValue.Number(BigDecimal(1250000)))

    // Verify styles
    val a1Style = sheet.getCellStyle(ARef.parse("A1").toOption.get)
    val b1Style = sheet.getCellStyle(ARef.parse("B1").toOption.get)
    val b2Style = sheet.getCellStyle(ARef.parse("B2").toOption.get)

    assert(a1Style.exists(_.font.bold), "A1 should be bold")
    assert(
      b1Style.exists(_.numFmt == com.tjclp.xl.styles.numfmt.NumFmt.Currency),
      "B1 should have currency format"
    )
    assert(
      b2Style.exists(_.numFmt == com.tjclp.xl.styles.numfmt.NumFmt.Percent),
      "B2 should have percent format"
    )
  }

  test("real-world example: financial report with string refs") {
    import com.tjclp.xl.macros.{fx, money}

    val bold = CellStyle.default.bold
    val currency = CellStyle.default.currency
    val percent = CellStyle.default.percent

    val sheet = emptySheet
      .put(
        "A1" -> "Revenue",
        "B1" -> money"$$1,250,000",
        "A2" -> "Expenses",
        "B2" -> money"$$875,000",
        "A3" -> "Net Income",
        "B3" -> fx"=B1-B2",
        "A4" -> "Margin",
        "B4" -> fx"=B3/B1"
      )
      .style(
        "A1:A4" -> bold,
        "B4"    -> percent
      )

    assertEquals(sheet.cellCount, 8)

    // Verify formula cells
    val b3Value = sheet(ARef.parse("B3").toOption.get).value
    assert(b3Value.isInstanceOf[CellValue.Formula], "B3 should be a formula")

    // Verify styles
    val a1Style = sheet.getCellStyle(ARef.parse("A1").toOption.get)
    val b4Style = sheet.getCellStyle(ARef.parse("B4").toOption.get)

    assert(a1Style.exists(_.font.bold), "A1 should be bold")
    assert(
      b4Style.exists(_.numFmt == com.tjclp.xl.styles.numfmt.NumFmt.Percent),
      "B4 should have percent format"
    )
  }

  // ========== Workbook Convenience Constructor Tests ==========

  test("Workbook(sheet) creates single-sheet workbook") {
    val sheet = emptySheet.put("A1" -> "Test")
    val workbook = Workbook(sheet)

    assertEquals(workbook.sheetCount, 1)
    assertEquals(workbook.sheets.head.name, sheet.name)
  }

  test("Workbook(sheet1, sheet2, ...) creates multi-sheet workbook") {
    val sheet1 = Sheet(SheetName.unsafe("Sales")).put("A1" -> "Revenue")
    val sheet2 = Sheet(SheetName.unsafe("Marketing")).put("A1" -> "Budget")
    val sheet3 = Sheet(SheetName.unsafe("Finance")).put("A1" -> "Assets")

    val workbook = Workbook(sheet1, sheet2, sheet3)

    assertEquals(workbook.sheetCount, 3)
    assertEquals(workbook.sheetNames.map(_.value), Seq("Sales", "Marketing", "Finance"))
  }
