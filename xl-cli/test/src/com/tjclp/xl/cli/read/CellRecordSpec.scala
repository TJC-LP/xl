package com.tjclp.xl.cli.read

import java.time.LocalDateTime

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{CellRange, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.cli.output.Escape
import com.tjclp.xl.macros.ref
import com.tjclp.xl.richtext.{RichText, TextRun}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * W2.4: `CellRecord` is the one projection of a cell. Its legacy JSON is the `view --format json`
 * cell byte for byte (the literals below are the shapes the goldens and the GH-357 tests pin), its
 * texts are the ones `search`, `cell` and the table renderers print.
 */
class CellRecordSpec extends FunSuite:

  private val data = SheetName.unsafe("Data")

  private def record(value: CellValue, numFmt: NumFmt = NumFmt.General): CellRecord =
    CellRecord.of(
      data,
      ref"A1",
      value,
      Some(CellStyle.default.withNumFmt(numFmt)),
      hidden = false,
      mergedInto = None
    )

  // ---------------------------------------------------------------------------------------------
  // toJson(legacyKeys = true): the view --format json cell
  // ---------------------------------------------------------------------------------------------

  test("legacy JSON: number keeps every digit and the formatted text") {
    assertEquals(
      record(CellValue.Number(BigDecimal(42))).toJson(legacyKeys = true),
      """{"ref": "A1", "type": "number", "value": 42, "formatted": "42"}"""
    )
    assertEquals(
      record(CellValue.Number(BigDecimal("12345678901234567"))).toJson(legacyKeys = true),
      """{"ref": "A1", "type": "number", "value": 12345678901234567, "formatted": "12345678901234567"}"""
    )
    assertEquals(
      record(CellValue.Number(BigDecimal("12.5")), NumFmt.Decimal).toJson(legacyKeys = true),
      """{"ref": "A1", "type": "number", "value": 12.5, "formatted": "12.50"}"""
    )
  }

  test("legacy JSON: text, boolean, error, rich text and empty keep the historical shapes") {
    assertEquals(
      record(CellValue.Text("Revenue")).toJson(legacyKeys = true),
      """{"ref": "A1", "type": "text", "value": "Revenue", "formatted": "Revenue"}"""
    )
    assertEquals(
      record(CellValue.Bool(true)).toJson(legacyKeys = true),
      """{"ref": "A1", "type": "boolean", "value": true, "formatted": "TRUE"}"""
    )
    assertEquals(
      record(CellValue.Error(CellError.Div0)).toJson(legacyKeys = true),
      """{"ref": "A1", "type": "error", "value": "#DIV/0!", "formatted": "#DIV/0!"}"""
    )
    assertEquals(
      record(CellValue.Empty).toJson(legacyKeys = true),
      """{"ref": "A1", "type": "empty", "value": null, "formatted": ""}"""
    )
    val rich = CellValue.RichText(RichText(TextRun("Bold").bold, TextRun(" plain")))
    assertEquals(
      record(rich).toJson(legacyKeys = true),
      """{"ref": "A1", "type": "richtext", "value": "Bold plain", "formatted": "Bold plain"}"""
    )
  }

  test("legacy JSON: a text cell is escaped once, through the one JSON escaper") {
    val text = "say \"hi\"\n\ttab \\ back"
    assertEquals(
      record(CellValue.Text(text)).toJson(legacyKeys = true),
      s"""{"ref": "A1", "type": "text", "value": ${Escape.json(text)}, "formatted": ${Escape
          .json(text)}}"""
    )
  }

  test("legacy JSON: formula cells carry the expression, then the cached value") {
    assertEquals(
      record(CellValue.Formula("=A1*2", Some(CellValue.Number(BigDecimal(84)))))
        .toJson(legacyKeys = true),
      """{"ref": "A1", "type": "formula", "formula": "=A1*2", "value": 84, "formatted": "84"}"""
    )
    assertEquals(
      record(CellValue.Formula("=A1+1")).toJson(legacyKeys = true),
      """{"ref": "A1", "type": "formula", "formula": "=A1+1", "value": null, "formatted": ""}"""
    )
    // an expression stored without its `=` is normalised in the field
    assertEquals(
      record(CellValue.Formula("1+1", Some(CellValue.Number(BigDecimal(2)))))
        .toJson(legacyKeys = true),
      """{"ref": "A1", "type": "formula", "formula": "=1+1", "value": 2, "formatted": "2"}"""
    )
    // a text-typed cached value is a JSON string
    assertEquals(
      record(CellValue.Formula("=B1", Some(CellValue.Text("Revenue")))).toJson(legacyKeys = true),
      """{"ref": "A1", "type": "formula", "formula": "=B1", "value": "Revenue", "formatted": "Revenue"}"""
    )
  }

  test("legacy JSON: record kinds gain the additive formulaKind field") {
    val array = record(
      CellValue.Formula(
        "SUM(A1:A2*10)",
        Some(CellValue.Number(BigDecimal(30))),
        FormulaKind.ArrayFormula(CellRange(ref"A1", ref"A1"))
      )
    )
    assert(
      array.toJson(legacyKeys = true).contains(""""formulaKind": "array""""),
      array.toJson(true)
    )
    assert(
      record(CellValue.Formula("=1+1", None))
        .toJson(legacyKeys = true)
        .contains("formulaKind") == false
    )
  }

  // ---------------------------------------------------------------------------------------------
  // The typed shape and the textualisers
  // ---------------------------------------------------------------------------------------------

  test("typed JSON: kind, sheet, formula object, hidden and mergedInto") {
    val merged = CellRecord
      .of(
        data,
        ref"A1",
        CellValue.Formula("=A2", Some(CellValue.Number(BigDecimal(7)))),
        None,
        hidden = true,
        mergedInto = Some(CellRange(ref"A1", ref"B2"))
      )
      .toJson(legacyKeys = false)
    val parsed = ujson.read(merged)
    assertEquals(parsed("kind").str, "formula")
    assertEquals(parsed("sheet").str, "Data")
    assertEquals(parsed("formula")("expression").str, "=A2")
    assertEquals(parsed("formula")("cached").bool, true)
    assertEquals(parsed("hidden").bool, true)
    assertEquals(parsed("mergedInto").str, "A1:B2")
    assertEquals(parsed("value").num, 7.0)
  }

  test("raw is the search text: digits, ISO date-times, TRUE/FALSE, error tokens, plain text") {
    assertEquals(record(CellValue.Number(BigDecimal("1000.50"))).raw, "1000.5")
    assertEquals(record(CellValue.Number(BigDecimal("1000.50")), NumFmt.Currency).raw, "1000.5")
    assertEquals(
      record(CellValue.DateTime(LocalDateTime.of(2024, 1, 15, 9, 30))).raw,
      "2024-01-15T09:30"
    )
    assertEquals(record(CellValue.Bool(false)).raw, "FALSE")
    assertEquals(record(CellValue.Error(CellError.Ref)).raw, "#REF!")
    assertEquals(record(CellValue.Empty).raw, "")
    // an uncached formula has no value: the search text is its expression
    assertEquals(record(CellValue.Formula("A1+1")).searchText, "=A1+1")
    assertEquals(
      record(CellValue.Formula("A1+1", Some(CellValue.Number(BigDecimal(3))))).searchText,
      "3"
    )
  }

  test("text(showFormulas) is the table cell: formatted value, or the formula-bar spelling") {
    val cached = record(CellValue.Formula("SUM(B1:B3)", Some(CellValue.Number(BigDecimal("42.5")))))
    assertEquals(cached.text(showFormulas = false), "42.5")
    assertEquals(cached.text(showFormulas = true), "=SUM(B1:B3)")
    assertEquals(record(CellValue.Formula("A1+1")).text(showFormulas = false), "=A1+1")
    assertEquals(record(CellValue.Number(BigDecimal("0.25")), NumFmt.Percent).text(false), "25%")
  }

  test("isEmpty: empty, blank text and a formula cached to those; never an uncached formula") {
    assert(record(CellValue.Empty).isEmpty)
    assert(record(CellValue.Text("   ")).isEmpty)
    assert(record(CellValue.Formula("=A1", Some(CellValue.Empty))).isEmpty)
    assert(record(CellValue.Formula("=A1", Some(CellValue.Text(" ")))).isEmpty)
    assert(!record(CellValue.Formula("=A1")).isEmpty)
    assert(!record(CellValue.Number(BigDecimal(0))).isEmpty)
  }

  test("cellValue round-trips the stored value, formula cache included") {
    val values = Vector(
      CellValue.Text("x"),
      CellValue.Number(BigDecimal(1)),
      CellValue.Empty,
      CellValue.Formula("A1", None),
      CellValue.Formula("A1", Some(CellValue.Text("y")))
    )
    values.foreach(v => assertEquals(record(v).cellValue, v))
  }

  test("evaluated: the computed value replaces the cache; a failure is the error token") {
    val base = record(CellValue.Formula("=A1+1"))
    val ok = CellRecord.evaluated(base, Right(CellValue.Number(BigDecimal(43))))
    assertEquals(
      ok.toJson(legacyKeys = true),
      """{"ref": "A1", "type": "formula", "formula": "=A1+1", "value": 43, "formatted": "43"}"""
    )
    val failed = CellRecord.evaluated(base, Left("#NAME?"))
    assertEquals(
      failed.toJson(legacyKeys = true),
      """{"ref": "A1", "type": "formula", "formula": "=A1+1", "value": "#NAME?", "formatted": "#NAME?"}"""
    )
  }
