package com.tjclp.xl.io

import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.streaming.SaxSingleCellReader

/** GH-370 parity and safety checks for the streaming worksheet readers. */
class SharedFormulaStreamingSpec extends CatsEffectSuite:

  test("row stream expands shared dependents") {
    rows(masterFirstXml).map { parsed =>
      assertEquals(valueAt(parsed, 1, 1), formula("A1+$A$1", 2))
      assertEquals(valueAt(parsed, 2, 1), formula("A2+$A$1", 3))
      assertEquals(valueAt(parsed, 3, 1), formula("A3+$A$1", 4))
    }
  }

  test("bounded row stream captures a shared master outside both row and column bounds") {
    val xml =
      """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1"><c r="B1"><f t="shared" si="0" ref="B1:C2">A1</f><v>1</v></c></row>
        |    <row r="2"><c r="C2"><f t="shared" si="0"/><v>2</v></c></row>
        |  </sheetData>
        |</worksheet>""".stripMargin

    rows(xml, Some((2, 2)), Some((2, 2))).map { parsed =>
      assertEquals(parsed.map(_.rowIndex), Vector(2))
      assertEquals(valueAt(parsed, 2, 2), formula("B2", 2))
    }
  }

  test("unrelated malformed shared cells outside a bounded selection do not poison it") {
    val xml =
      """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1">
        |      <c r="A1"><f t="shared" si="99"/><v>99</v></c>
        |      <c r="B1"><f t="shared" si="0" ref="B1:B2">A1</f><v>1</v></c>
        |    </row>
        |    <row r="2"><c r="B2"><f t="shared" si="0"/><v>2</v></c></row>
        |  </sheetData>
        |</worksheet>""".stripMargin

    rows(xml, Some((2, 2)), Some((1, 1))).map { parsed =>
      assertEquals(valueAt(parsed, 2, 1), formula("A2", 2))
    }
  }

  test("row stream keeps a dependent-before-master formula-shaped") {
    val xml =
      """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1"><c r="B1"><f t="shared" si="0"/><v>1</v></c></row>
        |    <row r="2"><c r="B2"><f t="shared" si="0" ref="B1:B2">A2</f><v>2</v></c></row>
        |  </sheetData>
        |</worksheet>""".stripMargin

    rows(xml).map { parsed =>
      assertEquals(valueAt(parsed, 1, 1), formula("#REF!", 1))
      assertEquals(valueAt(parsed, 2, 1), formula("A2", 2))
    }
  }

  test("row stream keeps a known out-of-range dependent formula-shaped") {
    val xml =
      """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1"><c r="B1"><f t="shared" si="5" ref="B1:B2">A1</f><v>1</v></c></row>
        |    <row r="2"><c r="A2"><f t="shared" si="5"/><v>3</v></c></row>
        |  </sheetData>
        |</worksheet>""".stripMargin

    rows(xml).map { parsed =>
      assertEquals(valueAt(parsed, 2, 0), formula("#REF!", 3))
    }
  }

  test("row stream keeps a past-end out-of-range dependent formula-shaped") {
    val xml =
      """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1"><c r="B1"><f t="shared" si="5" ref="B1:B2">A1</f><v>1</v></c></row>
        |    <row r="2"><c r="B2"><f t="shared" si="5"/><v>2</v></c></row>
        |    <row r="3">
        |      <c r="A3"><v>0</v></c>
        |      <c r="C3"><f t="shared" si="5"/><v>3</v></c>
        |    </row>
        |  </sheetData>
        |</worksheet>""".stripMargin

    rows(xml).map { parsed =>
      assertEquals(valueAt(parsed, 2, 1), formula("A2", 2))
      assertEquals(valueAt(parsed, 3, 2), formula("#REF!", 3))
    }
  }

  test("single-cell reader expands a master that precedes the target") {
    val result = extract(masterFirstXml, "B3").getOrElse(fail("missing B3"))
    assertEquals(result.value, formula("A3+$A$1", 4))
    assertEquals(result.formulaText, Some("A3+$A$1"))
  }

  test("single-cell reader defers a target whose master appears later") {
    val xml =
      """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1"><c r="B1"><f t="shared" si="7">JUNK99</f><v>1</v></c></row>
        |    <row r="2"><c r="B2"><f t="shared" si="07" ref="B1:B2">A2</f><v>2</v></c></row>
        |  </sheetData>
        |</worksheet>""".stripMargin

    val result = extract(xml, "B1").getOrElse(fail("missing B1"))
    assertEquals(result.value, formula("A1", 1))
    assertEquals(result.formulaText, Some("A1"))
  }

  test("single-cell reader keeps an orphan dependent formula-shaped") {
    val xml =
      """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1">
        |      <c r="A1"><f t="shared">UNRELATED()</f><v>5</v></c>
        |      <c r="B1"><f t="shared" si="404"/><v>9</v></c>
        |    </row>
        |  </sheetData>
        |</worksheet>""".stripMargin

    val result = extract(xml, "B1").getOrElse(fail("missing B1"))
    assertEquals(result.value, formula("#REF!", 9))
    assertEquals(result.formulaText, Some("#REF!"))
  }

  test("single-cell reader rejects a matching group outside its declared range") {
    val xml =
      """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        |  <sheetData>
        |    <row r="1"><c r="B1"><f t="shared" si="8" ref="B1:B2">A1</f><v>1</v></c></row>
        |    <row r="2"><c r="C2"><f t="shared" si="8">STALE1</f><v>9</v></c></row>
        |  </sheetData>
        |</worksheet>""".stripMargin

    val result = extract(xml, "C2").getOrElse(fail("missing C2"))
    assertEquals(result.value, formula("#REF!", 9))
    assertEquals(result.formulaText, Some("#REF!"))
  }

  private val masterFirstXml =
    """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      |  <sheetData>
      |    <row r="1">
      |      <c r="A1"><v>1</v></c>
      |      <c r="B1"><f t="shared" si="01" ref="B1:B3">A1+$A$1</f><v>2</v></c>
      |    </row>
      |    <row r="2"><c r="B2"><f t="shared" si="1"/><v>3</v></c></row>
      |    <row r="3"><c r="B3"><f t="shared" si="1"/><v>4</v></c></row>
      |  </sheetData>
      |</worksheet>""".stripMargin

  private def rows(
    xml: String,
    rowBounds: Option[(Int, Int)] = None,
    colBounds: Option[(Int, Int)] = None
  ): IO[Vector[RowData]] =
    SaxStreamingReader
      .parseWorksheetStream[IO](input(xml), None, rowBounds, colBounds)
      .compile
      .toVector

  private def extract(xml: String, address: String): Option[SaxSingleCellReader.CellResult] =
    SaxSingleCellReader.extractCell(input(xml), cellRef(address), None)

  private def input(xml: String): ByteArrayInputStream =
    new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8))

  private def cellRef(address: String): ARef =
    ARef.parse(address).fold(error => fail(error), identity)

  private def valueAt(rows: Vector[RowData], row: Int, col: Int): CellValue =
    rows
      .find(_.rowIndex == row)
      .flatMap(_.cells.get(col))
      .getOrElse(fail(s"missing row=$row col=$col"))

  private def formula(expression: String, cached: Int): CellValue =
    CellValue.Formula(expression, Some(CellValue.Number(BigDecimal(cached))))
