package com.tjclp.xl.formula

import java.time.{LocalDate, LocalDateTime}

import munit.FunSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cf.CfPaint
import com.tjclp.xl.formula.eval.{CfEvaluation, CfUnevaluated}
import com.tjclp.xl.ooxml.{TestFixtures, XlsxReader}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-497: the evaluator over the real conditional-formatting corpus. `condformat.xlsx` (openpyxl)
 * carries one rule of every typed kind plus an icon set; `condformat-lo.xlsx` (LibreOffice) carries
 * the same rules with default-value noise, all outside the typed model.
 */
class CfEvaluatorFixtureSpec extends FunSuite:

  private val clock = Clock.fixed(LocalDate.of(2026, 9, 25), LocalDateTime.of(2026, 9, 25, 12, 0))
  private val window = CellRange.parse("A1:J9").fold(fail(_), identity)
  private val pink = Some(Fill.Solid(Color.Rgb(0xffffc7ce)))
  private val yellow = Some(Fill.Solid(Color.Rgb(0xffffeb9c)))

  private def aref(a1: String): ARef = ARef.parse(a1).fold(fail(_), identity)

  private def evaluate(fixture: String): (Workbook, CfEvaluation) =
    val wb = XlsxReader.read(TestFixtures.copyToTemp(fixture)).fold(e => fail(e.message), identity)
    val sheet = wb.sheets.find(_.name == SheetName.unsafe("CondFmt")).getOrElse(fail("CondFmt"))
    (wb, sheet.evaluateConditionalFormats(window, Some(wb), clock))

  private def at(ev: CfEvaluation, a1: String): Option[CfPaint] = ev.overlay.at(aref(a1))

  private def painted(ev: CfEvaluation, col: Char): Set[String] =
    ev.overlay.cells.keySet.map(_.toA1).filter(_.headOption.contains(col))

  test(
    "condformat.xlsx: cellIs > 100 paints pink and dark-red bold; its fill beats the stop rule"
  ) {
    val (_, ev) = evaluate("condformat.xlsx")
    assertEquals(painted(ev, 'B'), Set("B3", "B5", "B7", "B9"))
    List("B3", "B5", "B7", "B9").foreach { a1 =>
      val dxf = at(ev, a1).map(_.dxf).getOrElse(fail(a1))
      assertEquals(dxf.fill, pink, s"$a1: priority 1's fill wins over the yellow expression")
      assertEquals(dxf.font.flatMap(_.bold), Some(true))
      assertEquals(dxf.font.flatMap(_.color), Some(Color.Rgb(0xff9c0006)))
    }
  }

  test("condformat.xlsx: the 3-point colour scale spans min, the 50th percentile and max") {
    val (_, ev) = evaluate("condformat.xlsx")
    def hex(a1: String) = at(ev, a1).flatMap(_.dxf.fill).collect { case Fill.Solid(c) =>
      c.toResolvedHex(ThemePalette.office)
    }
    assertEquals(hex("C6"), Some("#F8696B"), "6 is the minimum")
    assertEquals(hex("C7"), Some("#63BE7B"), "1998 is the maximum")
    // 174 sits 168/182 of the way from the minimum (6) to the median (188)
    assertEquals(hex("C4"), Some("#FEE182"))
    assertEquals(painted(ev, 'C').size, 8)
  }

  test("condformat.xlsx: the data bar runs 10% at the minimum to 90% at the maximum") {
    val (_, ev) = evaluate("condformat.xlsx")
    assertEquals(at(ev, "D6").flatMap(_.bar).map(_.fraction), Some(0.1), "3 is the minimum")
    assertEqualsDouble(at(ev, "D4").flatMap(_.bar).map(_.fraction).getOrElse(-1.0), 0.9, 1e-9)
    assertEquals(at(ev, "D2").flatMap(_.bar).map(_.color), Some(Color.Rgb(0xff638ec6)))
  }

  test("condformat.xlsx: containsText, top 3 and the multi-range between") {
    val (_, ev) = evaluate("condformat.xlsx")
    assertEquals(painted(ev, 'E'), Set("E2", "E4", "E6", "E8"))
    assertEquals(painted(ev, 'F'), Set("F3", "F7", "F9"))
    List("F3", "F7", "F9").foreach(a1 => assertEquals(at(ev, a1).flatMap(_.dxf.fill), yellow))
    assertEquals(painted(ev, 'H') ++ painted(ev, 'J'), Set("H3", "H4", "H5", "J2", "J3"))
  }

  test("condformat.xlsx: the icon set is unpainted and reported, alone") {
    val (_, ev) = evaluate("condformat.xlsx")
    assertEquals(painted(ev, 'G'), Set.empty[String])
    assertEquals(
      ev.unevaluated.map(u => (u.priority, u.kind, u.reason)),
      Vector((Some(7), "iconSet", CfUnevaluated.Reason.NotModeled))
    )
  }

  test("condformat-lo.xlsx: every rule is outside the typed model, so nothing paints") {
    val (_, ev) = evaluate("condformat-lo.xlsx")
    assertEquals(ev.overlay, CfOverlay.empty)
    assertEquals(ev.unevaluated.size, 8)
    assert(ev.unevaluated.forall(_.reason == CfUnevaluated.Reason.NotModeled), ev.unevaluated)
    val priorities = ev.unevaluated.flatMap(_.priority)
    assertEquals(priorities, priorities.sorted, "reported in precedence order")
    assertEquals(ev.unevaluated.map(_.kind).count(_ == "iconSet"), 1, ev.unevaluated.map(_.kind))
  }
