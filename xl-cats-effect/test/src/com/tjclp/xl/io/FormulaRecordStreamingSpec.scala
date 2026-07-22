package com.tjclp.xl.io

import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.io.streaming.SaxSingleCellReader

/**
 * GH-430 streaming parity: the SAX row stream and the single-cell reader model `t="array"` /
 * `t="dataTable"` records exactly like the DOM reader — `readStream -> write` must not bake data
 * tables to constants.
 */
class FormulaRecordStreamingSpec extends CatsEffectSuite:

  private val recordsXml =
    """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      |  <sheetData>
      |    <row r="1">
      |      <c r="A1"><v>1</v></c>
      |      <c r="C1" t="n"><f t="array" ref="C1:C1">SUM(A1:A3*10)</f><v>60</v></c>
      |    </row>
      |    <row r="2">
      |      <c r="A2"><v>2</v></c>
      |      <c r="F2"><f t="dataTable" ref="F2:G2" dt2D="1" dtr="0" r1="A1" r2="A2"/><v>42</v></c>
      |      <c r="G2" t="e"><f t="dataTable" ref="F2:G2" dt2D="1" dtr="0" r1="A1" r2="A2" ca="1"/><v>#NUM!</v></c>
      |    </row>
      |  </sheetData>
      |</worksheet>""".stripMargin

  private val c1 = range("C1:C1")
  private val f2g2 = range("F2:G2")
  private val arrayValue =
    CellValue.Formula(
      "SUM(A1:A3*10)",
      Some(CellValue.Number(BigDecimal(60))),
      FormulaKind.ArrayFormula(c1)
    )
  private val tableValue =
    CellValue.Formula(
      "TABLE(A1,A2)",
      Some(CellValue.Number(BigDecimal(42))),
      FormulaKind.DataTable(
        f2g2,
        dt2D = true,
        dtr = false,
        r1 = Some(ref("A1")),
        r2 = Some(ref("A2"))
      )
    )
  private val tableErrValue =
    CellValue.Formula(
      "TABLE(A1,A2)",
      Some(CellValue.Error(CellError.Num)),
      FormulaKind
        .DataTable(
          f2g2,
          dt2D = true,
          dtr = false,
          r1 = Some(ref("A1")),
          r2 = Some(ref("A2")),
          ca = true
        )
    )

  test("row stream models array + dataTable records (incl. cached #NUM!) like the DOM reader") {
    rows(recordsXml).map { parsed =>
      assertEquals(valueAt(parsed, 1, 2), arrayValue)
      assertEquals(valueAt(parsed, 2, 5), tableValue)
      assertEquals(valueAt(parsed, 2, 6), tableErrValue)
    }
  }

  test("row stream degrades junk record attrs exactly like the DOM reader") {
    val xml =
      """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1">
        |      <c r="A1"><f t="array">A9+1</f><v>1</v></c>
        |      <c r="B1"><f t="dataTable" dt2D="1" dtr="0" r1="A1" r2="A2"/><v>4</v></c>
        |      <c r="C1"><f t="array" ref="C1:C1"/><v>3</v></c>
        |    </row>
        |  </sheetData>
        |</worksheet>""".stripMargin
    rows(xml).map { parsed =>
      // no ref: Normal formula keeping text; no parsable ref on empty dataTable: cached constant;
      // empty array text: cached constant
      assertEquals(
        valueAt(parsed, 1, 0),
        CellValue.Formula("A9+1", Some(CellValue.Number(BigDecimal(1))))
      )
      assertEquals(valueAt(parsed, 1, 1), CellValue.Number(BigDecimal(4)))
      assertEquals(valueAt(parsed, 1, 2), CellValue.Number(BigDecimal(3)))
    }
  }

  test("single-cell reader models the dataTable record (kind parity)") {
    val result = extract(recordsXml, "F2").getOrElse(fail("missing F2"))
    assertEquals(result.value, tableValue)
    assertEquals(result.formulaText, Some("TABLE(A1,A2)"))
  }

  test("single-cell reader models the array record (kind parity)") {
    val result = extract(recordsXml, "C1").getOrElse(fail("missing C1"))
    assertEquals(result.value, arrayValue)
    assertEquals(result.formulaText, Some("SUM(A1:A3*10)"))
  }

  test("single-cell reader models the ca'd cached-#NUM! interior") {
    val result = extract(recordsXml, "G2").getOrElse(fail("missing G2"))
    assertEquals(result.value, tableErrValue)
  }

  private def rows(xml: String): IO[Vector[RowData]] =
    SaxStreamingReader
      .parseWorksheetStream[IO](input(xml), None, None, None)
      .compile
      .toVector

  private def extract(xml: String, address: String): Option[SaxSingleCellReader.CellResult] =
    SaxSingleCellReader.extractCell(input(xml), ref(address), None)

  private def input(xml: String): ByteArrayInputStream =
    new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8))

  private def ref(address: String): ARef =
    ARef.parse(address).fold(error => fail(error), identity)

  private def range(address: String): CellRange =
    CellRange.parse(address).fold(error => fail(error), identity)

  private def valueAt(rows: Vector[RowData], row: Int, col: Int): CellValue =
    rows
      .find(_.rowIndex == row)
      .flatMap(_.cells.get(col))
      .getOrElse(fail(s"missing row=$row col=$col"))
