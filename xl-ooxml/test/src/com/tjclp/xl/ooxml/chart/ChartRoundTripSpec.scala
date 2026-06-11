package com.tjclp.xl.ooxml.chart

import munit.ScalaCheckSuite
import org.scalacheck.Prop
import org.scalacheck.Prop.forAll

import com.tjclp.xl.Generators
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.charts.Chart
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.XmlUtil
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-222 KEYSTONE self-coherence law: `ChartReader.parse(emit(chart)) == Some(chart)` for arbitrary
 * generated charts — proves the emitted dialect is a subset of the read fence mechanically (every
 * element/attribute the canonical emitter produces must pass the strict whitelist, with
 * caches/idx/order/axes dropped losslessly for the model).
 *
 * Exercised under BOTH cache regimes: bare `c:f` (referenced sheet absent from the workbook) and
 * fully-resolved caches over a sheet mixing numbers, text, booleans, blanks, dates, and
 * cached-formula values.
 *
 * Determinism: pinned ScalaCheck seed, the OoxmlGenerativeRoundTripSpec convention.
 */
class ChartRoundTripSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters: org.scalacheck.Test.Parameters =
    super.scalaCheckTestParameters
      .withMinSuccessfulTests(200)
      .withInitialSeed(org.scalacheck.rng.Seed(20260611L))

  private def emit(chart: Chart, workbook: Workbook): String =
    XmlUtil.compact(OoxmlChart(chart, ChartCaches.resolve(workbook, chart)).toXml)

  private def lawProp(chart: Chart, workbook: Workbook): Prop =
    val xml = emit(chart, workbook)
    val parsed = ChartReader.parse(xml)
    Prop(parsed == Some(chart)) :|
      s"parse(emit(chart)) != Some(chart)\nchart:  $chart\nparsed: $parsed\nxml: $xml"

  /** Grid sheet covering the generator space with every cache-relevant value class. */
  private def materialized(name: SheetName): Sheet =
    (0 to 11).foldLeft(Sheet(name)) { (sheet, row) =>
      (0 to 7).foldLeft(sheet) { (s, col) =>
        val cellRef = com.tjclp.xl.addressing.ARef.from0(col, row)
        col match
          case 0 | 1 | 2 => s.put(Cell(cellRef, CellValue.Number(BigDecimal(row * 10 + col))))
          case 3 => s.put(Cell(cellRef, CellValue.Number(BigDecimal("1.5") * row - 7)))
          case 4 | 5 => s.put(Cell(cellRef, CellValue.Text(s"v$row$col")))
          case 6 => s // blank column: pts skipped, ptCount still counts
          case _ =>
            row % 3 match
              case 0 => s.put(Cell(cellRef, CellValue.Bool(row % 2 == 0)))
              case 1 =>
                s.put(
                  Cell(
                    cellRef,
                    CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(row))))
                  )
                )
              case _ =>
                s.put(
                  Cell(
                    cellRef,
                    CellValue.DateTime(java.time.LocalDateTime.of(2026, 6, 1 + row, 12, 0))
                  )
                )
      }
    }

  private val plain = SheetName.unsafe("Data")
  private val quoted = SheetName.unsafe("Q1 'Report") // needsQuoting: space + quote + cell-shaped

  property("LAW: parse(emit(chart)) == chart — bare c:f (referenced sheet absent)") {
    forAll(Generators.genChart(plain)) { chart =>
      lawProp(chart, Workbook(Vector(Sheet(SheetName.unsafe("Other")))))
    }
  }

  property("LAW: parse(emit(chart)) == chart — resolved caches over a mixed-value sheet") {
    forAll(Generators.genChart(plain)) { chart =>
      lawProp(chart, Workbook(Vector(materialized(plain))))
    }
  }

  property("LAW: parse(emit(chart)) == chart — quoting-positive sheet name, resolved caches") {
    forAll(Generators.genChart(quoted)) { chart =>
      lawProp(chart, Workbook(Vector(materialized(quoted))))
    }
  }

  test("emission is write-twice stable and backend-agnostic at the model level") {
    val chart = Chart(
      com.tjclp.xl.charts.ChartType.Bar(),
      Vector(
        com.tjclp.xl.charts.Series(
          com.tjclp.xl.charts.DataRef(plain, ref"B2:B5"),
          Some(com.tjclp.xl.charts.DataRef(plain, ref"A2:A5")),
          Some(com.tjclp.xl.charts.SeriesName.FromCell(plain, ref"B1"))
        )
      ),
      Some("T"),
      Some(com.tjclp.xl.charts.Legend())
    )
    val wb = Workbook(Vector(materialized(plain)))
    assertEquals(emit(chart, wb), emit(chart, wb))
  }
