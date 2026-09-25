package com.tjclp.xl.formula

import java.time.{LocalDate, LocalDateTime}

import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.forAll

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.cf.{CfBar, CfPaint}
import com.tjclp.xl.formula.eval.{CfEvaluation, CfUnevaluated}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-497: Excel's conditional-formatting semantics for the renderers — rule kinds, precedence and
 * stopIfTrue, relative anchoring, statistics over the whole block, and a total evaluation whose
 * every unpainted rule is reported once.
 */
class CfEvaluatorSpec extends ScalaCheckSuite:

  private val pink = Color.Rgb(0xffffc7ce)
  private val yellow = Color.Rgb(0xffffeb9c)
  private val red = Color.Rgb(0xffff0000)
  private val green = Color.Rgb(0xff00ff00)
  private val blue = Color.Rgb(0xff638ec6)
  private val clock = Clock.fixed(LocalDate.of(2026, 9, 25), LocalDateTime.of(2026, 9, 25, 12, 0))

  private def range(a1: String): CellRange = CellRange.parse(a1).fold(fail(_), identity)
  private def aref(a1: String): ARef = ARef.parse(a1).fold(fail(_), identity)
  private def num(n: Double): CellValue = CellValue.Number(BigDecimal(n))
  private def text(s: String): CellValue = CellValue.Text(s)

  /** Column A from A1 down. */
  private def column(values: CellValue*): Sheet =
    values.zipWithIndex.foldLeft(Sheet("S")) { case (s, (v, i)) => s.put(ARef.from0(0, i), v) }

  private def block(a1: String, rules: CfRule*): ConditionalFormat =
    ConditionalFormat.Rules(a1.split(' ').toVector.map(range), rules.toVector)

  private def withCf(sheet: Sheet, blocks: ConditionalFormat*): Sheet =
    sheet.copy(conditionalFormats = blocks.toVector)

  private def run(sheet: Sheet, window: String = "A1:J20", wb: Option[Workbook] = None) =
    sheet.evaluateConditionalFormats(range(window), wb, clock)

  private def fillAt(ev: CfEvaluation, a1: String): Option[Fill] =
    ev.overlay.at(aref(a1)).flatMap(_.dxf.fill)

  private def painted(ev: CfEvaluation): Set[String] = ev.overlay.cells.keySet.map(_.toA1)

  private def cellIs(op: CfOperator, f: String, p: Int = 1, stop: Boolean = false): CfRule =
    CfRule.CellIs(op, f, None, Some(Dxf.fill(pink)), p, stop)

  private def expr(f: String, p: Int = 1, stop: Boolean = false): CfRule =
    CfRule.Expression(f, Some(Dxf.fill(pink)), p, stop)

  private def scale(p: Int, mid: Option[CfPoint] = None): CfRule =
    CfRule.ColorScale(CfPoint(Cfvo.Min, red), mid, CfPoint(Cfvo.Max, green), p)

  private def bar(p: Int, min: Cfvo = Cfvo.Min, max: Cfvo = Cfvo.Max): CfRule =
    CfRule.DataBar(min, max, blue, showValue = true, p)

  // ========== Array evaluation ==========

  test("a formula rule evaluates as an array formula, as Excel evaluates it (not a plain cell)") {
    // Excel evaluates conditional-format formulas as arrays: the duplicate test and the list
    // membership test read every cell of their ranges, not the rule cell's row of them
    val values = column(num(3), num(5), num(3), num(9))
    val dupes = run(withCf(values, block("A1:A4", expr("SUM(($A$1:$A$4=A1)*1)>1"))))
    assertEquals(painted(dupes), Set("A1", "A3"))
    val list = withCf(
      values.put(aref("C1"), num(9)).put(aref("C2"), num(5)),
      block("A1:A4", expr("OR(A1=$C$1:$C$2)"))
    )
    assertEquals(painted(run(list)), Set("A2", "A4"))
  }

  // ========== Precedence ==========

  test("non-conflicting properties from every true rule combine") {
    val bold = CfRule.Expression("TRUE", Some(Dxf.font(DxfFont(bold = Some(true)))), 1)
    val fill = CfRule.Expression("TRUE", Some(Dxf.fill(pink)), 2)
    val ev = run(withCf(column(num(1)), block("A1", bold, fill)))
    assertEquals(
      ev.overlay.at(aref("A1")).map(_.dxf),
      Some(Dxf(font = Some(DxfFont(bold = Some(true))), fill = Some(Fill.Solid(pink))))
    )
  }

  test("the lower priority number wins a conflict, whatever the document order") {
    val low = CfRule.Expression("TRUE", Some(Dxf.fill(yellow)), 2)
    val high = CfRule.Expression("TRUE", Some(Dxf.fill(pink)), 1)
    val ev = run(withCf(column(num(1)), block("A1", low), block("A1", high)))
    assertEquals(fillAt(ev, "A1"), Some(Fill.Solid(pink)))
  }

  test("duplicate priorities fall back to document order") {
    val first = CfRule.Expression("TRUE", Some(Dxf.fill(yellow)), 1)
    val second = CfRule.Expression("TRUE", Some(Dxf.fill(pink)), 1)
    val ev = run(withCf(column(num(1)), block("A1", first), block("A1", second)))
    assertEquals(fillAt(ev, "A1"), Some(Fill.Solid(yellow)))
  }

  test("stopIfTrue on a true rule stops every lower rule: dxf, colour scale and data bar") {
    val sheet = column(num(1), num(5), num(10))
    val stop = CfRule.Expression("A1>=5", Some(Dxf.font(DxfFont(bold = Some(true)))), 1, true)
    val ev = run(withCf(sheet, block("A1:A3", stop, expr("TRUE", 2), scale(3), bar(4))))
    List("A2", "A3").foreach { a1 =>
      assertEquals(
        ev.overlay.at(aref(a1)),
        Some(CfPaint(Dxf.font(DxfFont(bold = Some(true))), None)),
        s"$a1 stops after the bold rule"
      )
    }
    // A1: the stop rule is false, so every lower rule applies (the dxf fill outranks the scale)
    val a1 = ev.overlay.at(aref("A1")).getOrElse(fail("A1 unpainted"))
    assertEquals(a1.dxf.fill, Some(Fill.Solid(pink)))
    assert(a1.bar.isDefined, "the bar still draws")
  }

  test("a true rule with no format and stopIfTrue suppresses the lower rules (Excel's guard)") {
    val guard = CfRule.Expression("LEN(TRIM(A1))=0", None, 1, stopIfTrue = true)
    val sheet = column(num(3), CellValue.Empty, num(9))
    val ev = run(withCf(sheet, block("A1:A3", guard, cellIs(CfOperator.LessThan, "5", 2))))
    assertEquals(painted(ev), Set("A1"), "the blank A2 is guarded, the 3 in A1 is painted")
  }

  test("a cell-value fill outranks a colour scale at a lower priority, and loses at a higher one") {
    val sheet = column(num(1), num(2))
    val above = run(withCf(sheet, block("A1:A2", cellIs(CfOperator.GreaterThan, "0", 1), scale(2))))
    assertEquals(fillAt(above, "A1"), Some(Fill.Solid(pink)))
    val below = run(withCf(sheet, block("A1:A2", scale(1), cellIs(CfOperator.GreaterThan, "0", 2))))
    assertEquals(fillAt(below, "A1"), Some(Fill.Solid(Color.Rgb(0xffff0000))))
  }

  test("two data bars keep only the higher-precedence one") {
    val other = CfRule.DataBar(Cfvo.Min, Cfvo.Max, red, showValue = false, 2)
    val ev = run(withCf(column(num(1), num(2)), block("A1:A2", other, bar(1))))
    assertEquals(ev.overlay.at(aref("A2")).flatMap(_.bar).map(_.color), Some(blue))
  }

  // ========== Cell value ==========

  test("cell value: all eight operators on numbers") {
    val sheet = column(num(1), num(5), num(9))
    def hits(op: CfOperator, f1: String, f2: Option[String] = None): Set[String] =
      painted(
        run(withCf(sheet, block("A1:A3", CfRule.CellIs(op, f1, f2, Some(Dxf.fill(pink)), 1))))
      )
    assertEquals(hits(CfOperator.LessThan, "5"), Set("A1"))
    assertEquals(hits(CfOperator.LessThanOrEqual, "5"), Set("A1", "A2"))
    assertEquals(hits(CfOperator.Equal, "5"), Set("A2"))
    assertEquals(hits(CfOperator.NotEqual, "5"), Set("A1", "A3"))
    assertEquals(hits(CfOperator.GreaterThanOrEqual, "5"), Set("A2", "A3"))
    assertEquals(hits(CfOperator.GreaterThan, "5"), Set("A3"))
    assertEquals(hits(CfOperator.Between, "1", Some("5")), Set("A1", "A2"), "inclusive")
    assertEquals(hits(CfOperator.Between, "9", Some("5")), Set("A2", "A3"), "operand order")
    assertEquals(hits(CfOperator.NotBetween, "2", Some("8")), Set("A1", "A3"))
  }

  test("cell value: a blank compares as zero, text above numbers, text case-insensitively") {
    val sheet = column(CellValue.Empty, text("yes"), num(3))
    def hits(op: CfOperator, f: String): Set[String] =
      painted(run(withCf(sheet, block("A1:A3", cellIs(op, f)))))
    assertEquals(hits(CfOperator.LessThan, "5"), Set("A1", "A3"), "a blank is less than 5")
    assertEquals(hits(CfOperator.Equal, "0"), Set("A1"), "a blank equals 0")
    assertEquals(hits(CfOperator.GreaterThan, "5"), Set("A2"), "text sorts above every number")
    assertEquals(hits(CfOperator.Equal, "\"YES\""), Set("A2"), "text compares case-insensitively")
  }

  test("cell value: an error cell never matches, not even 'not equal'") {
    val sheet = column(CellValue.Error(CellError.Div0), num(1))
    val ev = run(withCf(sheet, block("A1:A2", cellIs(CfOperator.NotEqual, "5"))))
    assertEquals(painted(ev), Set("A2"))
    assertEquals(ev.unevaluated, Vector.empty, "an error value is no match, not a failure")
  }

  test("cell value: a relative operand shifts per cell, an absolute one stays put") {
    val sheet = column(num(1), num(5), num(9))
      .put(aref("C1"), num(2))
      .put(aref("C2"), num(6))
      .put(aref("C3"), num(1))
      .put(aref("D1"), num(4))
    val relative = run(withCf(sheet, block("A1:A3", cellIs(CfOperator.LessThan, "C1"))))
    assertEquals(painted(relative), Set("A1", "A2"), "A2 compares to C2")
    val absolute = run(withCf(sheet, block("A1:A3", cellIs(CfOperator.GreaterThan, "$D$1"))))
    assertEquals(painted(absolute), Set("A2", "A3"))
  }

  test("cell value: a cached formula's value is used, an uncached one is computed") {
    val sheet = column(CellValue.Formula("1+1", Some(num(200))), CellValue.Formula("100*2", None))
    val ev = run(withCf(sheet, block("A1:A2", cellIs(CfOperator.GreaterThan, "100"))))
    assertEquals(painted(ev), Set("A1", "A2"))
  }

  test("cell value: an operand that does not parse paints nothing and is reported once") {
    val ev =
      run(withCf(column(num(1), num(2)), block("A1:A2", cellIs(CfOperator.GreaterThan, "1+"))))
    assertEquals(painted(ev), Set.empty[String])
    ev.unevaluated match
      case Vector(CfUnevaluated(Some(1), "cellIs", _, CfUnevaluated.Reason.FormulaFailed(f, m))) =>
        assertEquals(f, "1+")
        assert(m.contains("Parse error"), m)
      case other => fail(s"expected one parse failure, got $other")
  }

  test("cell value: between without a second operand is reported, never painted") {
    val rule = CfRule.CellIs(CfOperator.Between, "1", None, Some(Dxf.fill(pink)), 1)
    val ev = run(withCf(column(num(1)), block("A1", rule)))
    assertEquals(painted(ev), Set.empty[String])
    assertEquals(ev.unevaluated.map(_.reason.getClass.getSimpleName), Vector("InvalidRule"))
  }

  test("a multi-range block anchors its formulas at the top-left of its bounding box") {
    // ranges C1:C3 and A1:A3: the formula is written for A1 (MS-XLS refBound), so C1 reads C1
    val sheet = column(num(0), num(0), num(0))
      .put(aref("C1"), num(1))
      .put(aref("C2"), num(0))
      .put(aref("C3"), num(1))
    val ev = run(withCf(sheet, block("C1:C3 A1:A3", expr("A1>0"))))
    assertEquals(painted(ev), Set("C1", "C3"))
  }

  // ========== Expression ==========

  test("expression truth: TRUE, a non-zero number; not FALSE, 0, text or an error") {
    List(
      "TRUE" -> true,
      "1" -> true,
      "-2.5" -> true,
      "DATE(2026,1,1)" -> true,
      "FALSE" -> false,
      "0" -> false,
      "\"yes\"" -> false,
      "1/0" -> false
    ).foreach { (f, expected) =>
      val ev = run(withCf(column(num(1)), block("A1", expr(f))))
      assertEquals(painted(ev).nonEmpty, expected, f)
      assertEquals(ev.unevaluated, Vector.empty, s"$f is never a failure")
    }
  }

  test("expression: an array result is judged by its top-left element") {
    val first = column(num(1)).put(aref("B1"), num(-1))
    val second = column(num(-1)).put(aref("B1"), num(1))
    assertEquals(painted(run(withCf(first, block("A1", expr("A1:B1>0"))))), Set("A1"))
    assertEquals(painted(run(withCf(second, block("A1", expr("A1:B1>0"))))), Set.empty[String])
  }

  test("expression: ROW() is the target cell, so banding works") {
    val sheet = column(num(1), num(1), num(1), num(1))
    val ev = run(withCf(sheet, block("A1:A4", expr("MOD(ROW(),2)=0"))))
    assertEquals(painted(ev), Set("A2", "A4"))
  }

  test("expression: a cross-sheet reference reads the workbook; without one it is reported once") {
    val data = Sheet("Data").put(aref("A1"), num(10))
    val sheet = withCf(column(num(1), num(2), num(3)), block("A1:A3", expr("Data!$A$1>5")))
    val wb = Workbook(Vector(sheet, data))
    assertEquals(painted(run(sheet, wb = Some(wb))), Set("A1", "A2", "A3"))
    val alone = run(sheet)
    assertEquals(painted(alone), Set.empty[String])
    assertEquals(alone.unevaluated.size, 1, s"one report per rule, not per cell: $alone")
    assertEquals(alone.unevaluated.map(_.kind), Vector("expression"))
  }

  test("a formula rule's report names the first cell it failed for, row-major") {
    val sheet = withCf(column(num(1), num(2), num(3)), block("A1:A3", expr("Data!$A$1>5")))
    run(sheet, "A2:A3").unevaluated match
      case Vector(
            report @ CfUnevaluated(
              Some(1),
              "expression",
              _,
              CfUnevaluated.Reason.FailedAt(cell, _, false)
            )
          ) =>
        assertEquals(cell, ref"A2")
        assert(report.message.contains(" not rendered: it could not be evaluated at A2"), report)
      case other => fail(s"expected one report for the expression, got $other")
  }

  test("a rule that fails for some cells is painted where it evaluates, and says so") {
    // A2 and A4 read a missing sheet, so `A<n> < 4` cannot be evaluated there; A1 and A3 can
    val sheet = withCf(
      column(num(1), CellValue.Formula("Missing!A1"), num(3), CellValue.Formula("Missing!A2")),
      block("A1:A4", cellIs(CfOperator.LessThan, "4"))
    )
    val ev = run(sheet, "A1:A4")
    assertEquals(painted(ev), Set("A1", "A3"))
    ev.unevaluated match
      case Vector(
            report @ CfUnevaluated(_, "cellIs", _, CfUnevaluated.Reason.FailedAt(cell, _, true))
          ) =>
        assertEquals(cell, ref"A2")
        assert(report.message.contains(" partly rendered: it could not be evaluated at A2"), report)
        assert(report.message.endsWith("it is painted on the cells where it evaluates"), report)
        assert(!report.message.contains("formula '4'"), s"the operand did not fail: $report")
      case other => fail(s"expected one partly-rendered report, got $other")
  }

  test("expression: an unknown function paints nothing and is reported, never thrown") {
    val ev = run(withCf(column(num(1), num(2)), block("A1:A2", expr("NOSUCHFN(A1)>0"))))
    assertEquals(painted(ev), Set.empty[String])
    assertEquals(ev.unevaluated.map(_.kind), Vector("expression"))
    assert(ev.unevaluated.forall(_.message.contains("NOSUCHFN")), ev.unevaluated.map(_.message))
  }

  test("expression: TODAY() reads the explicit clock") {
    val sheet = column(
      CellValue.DateTime(LocalDateTime.of(2026, 9, 24, 0, 0)),
      CellValue.DateTime(LocalDateTime.of(2026, 9, 26, 0, 0))
    )
    val ev = run(withCf(sheet, block("A1:A2", expr("A1<TODAY()"))))
    assertEquals(painted(ev), Set("A1"))
  }

  // ========== Text ==========

  test("text rules evaluate as the SEARCH/LEFT/RIGHT formula Excel stores") {
    val sheet = column(text("ToDo: ship"), text("done"), CellValue.Empty, num(123), text("abc"))
    def hits(op: CfTextOp, t: String): Set[String] =
      painted(run(withCf(sheet, block("A1:A5", CfRule.Text(op, t, Some(Dxf.fill(pink)), 1)))))
    assertEquals(hits(CfTextOp.Contains, "todo"), Set("A1"), "case-insensitive")
    assertEquals(hits(CfTextOp.Contains, "a*c"), Set("A5"), "SEARCH wildcards")
    assertEquals(hits(CfTextOp.Contains, "2"), Set("A4"), "a number is searched as its text")
    assertEquals(hits(CfTextOp.NotContains, "o"), Set("A3", "A4", "A5"), "a blank does not contain")
    assertEquals(hits(CfTextOp.BeginsWith, "do"), Set("A2"))
    assertEquals(hits(CfTextOp.EndsWith, "ship"), Set("A1"))
  }

  // ========== Top 10 ==========

  private def top(rank: Int, percent: Boolean = false, bottom: Boolean = false): CfRule =
    CfRule.Top10(rank, percent, bottom, Some(Dxf.fill(pink)), 1)

  test("top N paints every value reaching the N-th largest, ties included; bottom N mirrors it") {
    val sheet = column(num(5), num(5), num(4), num(4), num(3))
    def hits(rule: CfRule) = painted(run(withCf(sheet, block("A1:A5", rule))))
    assertEquals(hits(top(2)), Set("A1", "A2"))
    assertEquals(hits(top(3)), Set("A1", "A2", "A3", "A4"))
    assertEquals(hits(top(1, bottom = true)), Set("A5"))
    assertEquals(hits(top(0)), Set.empty[String])
    assertEquals(hits(top(99)), Set("A1", "A2", "A3", "A4", "A5"))
  }

  test("top N percent: floor(n * rank / 100) values, and at least one") {
    val ten = column((1 to 10).map(i => num(i.toDouble))*)
    assertEquals(
      painted(run(withCf(ten, block("A1:A10", top(20, percent = true))))),
      Set("A9", "A10")
    )
    val five = column((1 to 5).map(i => num(i.toDouble))*)
    assertEquals(painted(run(withCf(five, block("A1:A5", top(10, percent = true))))), Set("A5"))
  }

  test("top N counts numbers and dates only, over the whole block, each cell once") {
    val sheet = column(
      num(1),
      text("999"),
      CellValue.Bool(true),
      CellValue.Error(CellError.NA),
      CellValue.Empty,
      num(50),
      CellValue.DateTime(LocalDateTime.of(2026, 1, 1, 0, 0)) // serial 46023
    )
    // A7 (a date) is the largest; the window A1:A6 sees A6 second-largest
    val ev = run(withCf(sheet, block("A1:A7 A5:A7", top(2))), window = "A1:A6")
    assertEquals(painted(ev), Set("A6"), "text, logicals, blanks and errors never paint")
  }

  test("a whole-column block reads the populated cells only") {
    val sheet = column(num(3), num(9), num(6))
    assertEquals(painted(run(withCf(sheet, block("A:A", top(1))), window = "A1:A5")), Set("A2"))
  }

  test("numeric rules read dates by the workbook's date system") {
    // 2026-01-01 is serial 46023 in the 1900 system and 44561 in the 1904 one: 45000 sits between
    val sheet = withCf(
      column(CellValue.DateTime(LocalDateTime.of(2026, 1, 1, 0, 0)), num(45000)),
      block("A1:A2", top(1))
    )
    val wb = Workbook(Vector(sheet))
    val mac = wb.copy(metadata = wb.metadata.copy(date1904 = true))
    assertEquals(painted(run(sheet, wb = Some(wb))), Set("A1"))
    assertEquals(painted(run(sheet, wb = Some(mac))), Set("A2"))
  }

  // ========== Uncached formulas under numeric rules ==========

  /** A1 = 0, then An = A(n-1)+1 down to row `n`: cached with each value, or uncached (openpyxl). */
  private def chain(n: Int, cached: Boolean): Sheet =
    (2 to n).foldLeft(column(num(0))) { (s, i) =>
      val value = Option.when(cached)(num((i - 1).toDouble))
      s.put(ARef.from0(0, i - 1), CellValue.Formula(s"A${i - 1}+1", value))
    }

  test("an uncached chain deeper than the evaluator's recursion guard paints as if cached") {
    List(scale(1), top(10), bar(1)).foreach { rule =>
      List("A1:A300", "A1:A5", "A250:A260").foreach { window =>
        val uncached = run(withCf(chain(300, cached = false), block("A:A", rule)), window)
        val cached = run(withCf(chain(300, cached = true), block("A:A", rule)), window)
        assertEquals(uncached.unevaluated, Vector.empty, s"$rule over $window")
        assertEquals(uncached.overlay, cached.overlay, s"$rule over $window")
      }
    }
    val ev = run(withCf(chain(300, cached = false), block("A:A", scale(1))), "A1:A300")
    assertEquals(scaleFill(ev, "A1"), Some("#FF0000"))
    assertEquals(scaleFill(ev, "A150"), Some("#807F00"), "149 of 0..299")
    assertEquals(scaleFill(ev, "A300"), Some("#00FF00"))
  }

  test("a block cell xl cannot compute stops the block's numeric rules, each reported once") {
    // A3 calls a function xl lacks: painting from the other cells would draw the wrong colours
    val sheet = column(
      num(1),
      CellValue.Formula("A1+1", None),
      CellValue.Formula("NOSUCHFN(1)", None),
      num(10)
    )
    val rules = block("A1:A4", scale(1), top(1), bar(3), expr("ROW()=4", 4))
    val ev = run(withCf(sheet, rules), "A1:A4")
    assertEquals(
      ev.overlay.cells.map((ref, paint) => ref.toA1 -> paint),
      Map("A4" -> CfPaint(Dxf.fill(pink), None)),
      "only the formula rule paints"
    )
    assertEquals(
      ev.unevaluated.map(u => (u.priority, u.kind)),
      Vector(
        (Some(1), "colorScale"),
        (Some(1), "top10"),
        (Some(3), "dataBar")
      )
    )
    ev.unevaluated.foreach {
      case CfUnevaluated(_, _, _, CfUnevaluated.Reason.FormulaFailed(f, m)) =>
        assertEquals(f, "NOSUCHFN(1)")
        assert(m.contains("NOSUCHFN") && m.contains("A3"), m)
      case other => fail(s"expected the cell's failure, got $other")
    }
  }

  test("an uncached chain running upward past the recursion guard is reported, not guessed") {
    // An = A(n+1)+1 up from A300 = 0: read top-down, each cell recurses to the bottom
    val sheet = (1 to 299).foldLeft(Sheet("S").put(aref("A300"), num(0))) { (s, i) =>
      s.put(ARef.from0(0, i - 1), CellValue.Formula(s"A${i + 1}+1", None))
    }
    val ev = run(withCf(sheet, block("A:A", scale(1))), "A1:A5")
    assertEquals(painted(ev), Set.empty[String])
    ev.unevaluated match
      case Vector(
            CfUnevaluated(Some(1), "colorScale", _, CfUnevaluated.Reason.FormulaFailed(f, m))
          ) =>
        assertEquals(f, "A2+1", "the first cell that fails, row-major")
        assert(m.contains("A1") && m.contains("depth"), m)
      case other => fail(s"expected one report for the scale, got $other")
  }

  test("a formula rule reading an uncached chain past the recursion guard is reported") {
    // each target is evaluated on its own: A295 recurses 294 cells deep; cached, the rule paints
    val rule = CfRule.CellIs(CfOperator.GreaterThan, "290", None, Some(Dxf.fill(pink)), 1)
    val ev = run(withCf(chain(300, cached = false), block("A1:A300", rule)), "A295:A300")
    assertEquals(painted(ev), Set.empty[String])
    ev.unevaluated match
      case Vector(
            CfUnevaluated(Some(1), "cellIs", _, CfUnevaluated.Reason.FailedAt(cell, m, false))
          ) =>
        assertEquals(cell, ref"A295")
        assert(m.contains("depth"), m)
      case other => fail(s"expected one report for the rule, got $other")
    val cached = run(withCf(chain(300, cached = true), block("A1:A300", rule)), "A295:A300")
    assertEquals(painted(cached), (295 to 300).map(i => s"A$i").toSet)
  }

  test("an uncached cell computing an Excel error is not a number, and never a failure") {
    val sheet = column(num(1), CellValue.Formula("1/0", None), num(3))
    val ev = run(withCf(sheet, block("A1:A3", top(1, bottom = true))), "A1:A3")
    assertEquals(painted(ev), Set("A1"))
    assertEquals(ev.unevaluated, Vector.empty)
  }

  test("a block that cannot compute a cell never stops another block's statistics") {
    // B1 reads the far end of the chain in one recursion, past the guard; A:A reads it row-major
    val sheet = chain(3000, cached = false)
      .put(aref("B1"), CellValue.Formula("A3000", None))
      .put(aref("B2"), num(5))
    val x = block("B1:B2", scale(1))
    val y = block("A:A", scale(2))
    val xAlone = run(withCf(sheet, x), "A1:B5")
    val yAlone = run(withCf(sheet, y), "A1:B5")
    assertEquals(yAlone.unevaluated, Vector.empty)
    assertEquals(scaleFill(yAlone, "A1"), Some("#FF0000"))
    xAlone.unevaluated match
      case Vector(
            CfUnevaluated(Some(1), "colorScale", _, CfUnevaluated.Reason.FormulaFailed(f, m))
          ) =>
        assertEquals(f, "A3000")
        assert(m.contains("B1"), m)
      case other => fail(s"expected one report for B1:B2's scale, got $other")
    List(Vector(x, y), Vector(y, x)).foreach { order =>
      val ev = run(withCf(sheet, order*), "A1:B5")
      assertEquals(ev.overlay, yAlone.overlay, s"order $order")
      assertEquals(ev.unevaluated, xAlone.unevaluated, s"order $order")
    }
  }

  test("a block's uncached RAND() draws the same whichever block is read first") {
    // seeded draws: 0.73 first, 0.24 second; A1 tops 0.5 only on the first draw
    val sheet = Sheet("S")
      .put(aref("A1"), CellValue.Formula("RAND()", None))
      .put(aref("A2"), num(0.5))
      .put(aref("C1"), CellValue.Formula("RAND()", None))
      .put(aref("C2"), num(0.5))
    val p = block("A1:A2", top(1))
    val q = block("C1:C2", top(1))
    val alone = painted(run(withCf(sheet, p), "A1:C2")) ++ painted(run(withCf(sheet, q), "A1:C2"))
    assertEquals(alone, Set("A1", "C1"))
    assertEquals(painted(run(withCf(sheet, p, q), "A1:C2")), alone)
    assertEquals(painted(run(withCf(sheet, q, p), "A1:C2")), alone)
  }

  // ========== Colour scale ==========

  private def scaleFill(ev: CfEvaluation, a1: String): Option[String] =
    fillAt(ev, a1).collect { case Fill.Solid(c) => c.toResolvedHex(ThemePalette.office) }

  test("colour scale: the ends take the end colours, between them it interpolates half-up") {
    val sheet = column(num(0), num(5), num(10))
    val rule = CfRule.ColorScale(
      CfPoint(Cfvo.Min, Color.Rgb(0xfff8696b)),
      None,
      CfPoint(Cfvo.Max, Color.Rgb(0xffffeb84)),
      1
    )
    val ev = run(withCf(sheet, block("A1:A3", rule)))
    assertEquals(scaleFill(ev, "A1"), Some("#F8696B"))
    assertEquals(scaleFill(ev, "A2"), Some("#FCAA78"))
    assertEquals(scaleFill(ev, "A3"), Some("#FFEB84"))
  }

  test("colour scale: a 3-point scale picks the segment around its percentile middle") {
    val sheet = column(num(0), num(1), num(2), num(10))
    val mid = CfPoint(Cfvo.Percentile(BigDecimal(50)), Color.Rgb(0xffffffff))
    val ev = run(withCf(sheet, block("A1:A4", scale(1, Some(mid)))))
    // PERCENTILE.INC({0,1,2,10}, 0.5) = 1.5: A2 (1) is two thirds of the way from red to white
    assertEquals(scaleFill(ev, "A2"), Some("#FFAAAA"))
    // A3 (2) is 0.5/8.5 of the way from white to green
    assertEquals(scaleFill(ev, "A3"), Some("#F0FFF0"))
  }

  test("colour scale: num, percent and percentile value objects") {
    val sheet = column((1 to 5).map(i => num(i.toDouble))*)
    def at(cfvo: Cfvo) = run(
      withCf(
        sheet,
        block("A1:A5", CfRule.ColorScale(CfPoint(cfvo, red), None, CfPoint(Cfvo.Max, green), 1))
      )
    )
    // a lower bound of 2 paints A1 and A2 with the first colour
    List(Cfvo.Num(BigDecimal(2)), Cfvo.Percent(BigDecimal(25)), Cfvo.Percentile(BigDecimal(25)))
      .foreach { cfvo =>
        val ev = at(cfvo)
        assertEquals(scaleFill(ev, "A2"), Some("#FF0000"), cfvo.toString)
        assertNotEquals(scaleFill(ev, "A3"), Some("#FF0000"), cfvo.toString)
      }
  }

  test("colour scale: a formula value object is evaluated once, at the anchor") {
    val sheet = column(num(1), num(2), num(3)).put(aref("E1"), num(2))
    val rule =
      CfRule.ColorScale(CfPoint(Cfvo.Formula("$E$1"), red), None, CfPoint(Cfvo.Max, green), 1)
    val ev = run(withCf(sheet, block("A1:A3", rule)))
    assertEquals(scaleFill(ev, "A2"), Some("#FF0000"))
    assertEquals(scaleFill(ev, "A3"), Some("#00FF00"))
  }

  test("colour scale: a text formula value object paints nothing and is reported") {
    val rule =
      CfRule.ColorScale(CfPoint(Cfvo.Formula("\"x\""), red), None, CfPoint(Cfvo.Max, green), 1)
    val ev = run(withCf(column(num(1), num(2)), block("A1:A2", rule)))
    assertEquals(painted(ev), Set.empty[String])
    assertEquals(ev.unevaluated.map(_.kind), Vector("colorScale"))
  }

  test("colour scale: a percentile outside 0..100 paints nothing and is reported") {
    val rule = CfRule.ColorScale(
      CfPoint(Cfvo.Percentile(BigDecimal(150)), red),
      None,
      CfPoint(Cfvo.Max, green),
      1
    )
    val ev = run(withCf(column(num(1), num(2)), block("A1:A2", rule)))
    assertEquals(painted(ev), Set.empty[String])
    assert(ev.unevaluated.exists(_.message.contains("150")), ev.unevaluated.map(_.message))
  }

  test("colour scale: all-equal values take the first colour; text and blanks are unpainted") {
    val ev =
      run(withCf(column(num(4), num(4), text("x"), CellValue.Empty), block("A1:A4", scale(1))))
    assertEquals(scaleFill(ev, "A1"), Some("#FF0000"))
    assertEquals(scaleFill(ev, "A2"), Some("#FF0000"))
    assertEquals(painted(ev), Set("A1", "A2"))
  }

  test("colour scale: an empty population paints nothing and reports nothing") {
    val ev = run(withCf(column(text("a"), text("b")), block("A1:A2", scale(1))))
    assertEquals(ev, CfEvaluation(CfOverlay.empty, Vector.empty))
  }

  test("colour scale: theme colours resolve with the workbook's theme, to RGB") {
    val rule = CfRule.ColorScale(
      CfPoint(Cfvo.Min, Color.Theme(ThemeSlot.Accent1, 0.0)),
      None,
      CfPoint(Cfvo.Max, green),
      1
    )
    val sheet = withCf(column(num(1), num(2)), block("A1:A2", rule))
    val purple = ThemePalette.office.copy(accent1 = 0xff7030a0)
    val wb = Workbook(Vector(sheet))
    val themed = wb.copy(metadata = wb.metadata.copy(theme = purple))
    assertEquals(
      fillAt(run(sheet, wb = Some(themed)), "A1"),
      Some(Fill.Solid(Color.Rgb(0xff7030a0)))
    )
    assertEquals(fillAt(run(sheet), "A1"), Some(Fill.Solid(Color.Rgb(0xff4f81bd))), "Office")
  }

  // ========== Data bar ==========

  test("data bar: 10% at the minimum, 90% at the maximum, linear between") {
    val sheet = column(num(0), num(5), num(10), text("x"))
    val ev = run(withCf(sheet, block("A1:A4", CfRule.DataBar(Cfvo.Min, Cfvo.Max, blue, false, 1))))
    def fraction(a1: String) = ev.overlay.at(aref(a1)).flatMap(_.bar).map(_.fraction)
    assertEquals(fraction("A1"), Some(0.1))
    assertEqualsDouble(fraction("A2").getOrElse(-1.0), 0.5, 1e-9)
    assertEqualsDouble(fraction("A3").getOrElse(-1.0), 0.9, 1e-9)
    assertEquals(fraction("A4"), None, "text gets no bar")
    assertEquals(ev.overlay.at(aref("A1")).flatMap(_.bar).map(_.showValue), Some(false))
  }

  // ========== Scope, reports, determinism ==========

  test("an unmodeled rule is reported by its kind; an unreadable block as 'block'") {
    val iconSet =
      CfRule.Preserved("""<cfRule type="iconSet" priority="7"><iconSet/></cfRule>""", Some(7))
    val sheet = withCf(
      column(num(1)),
      block("A1", iconSet),
      ConditionalFormat.Preserved(
        """<conditionalFormatting sqref="??"><cfRule priority="3"/></conditionalFormatting>"""
      )
    )
    val ev = run(sheet)
    assertEquals(painted(ev), Set.empty[String])
    assertEquals(
      ev.unevaluated.map(u => (u.priority, u.kind, u.reason)),
      Vector(
        (Some(3), "block", CfUnevaluated.Reason.NotModeled),
        (Some(7), "iconSet", CfUnevaluated.Reason.NotModeled)
      )
    )
    assertEquals(
      ev.unevaluated(1).message,
      "conditional format iconSet (priority 7) on A1 not rendered: " +
        "xl does not evaluate this rule (its kind or its formatting) yet"
    )
  }

  test("a Preserved rule's kind is read off its cfRule tag, past any XML declaration") {
    val payload = """<?xml version="1.0" encoding="UTF-8"?><cfRule xmlns="urn:x" priority="4" """ +
      """type="dataBar"><dataBar/><extLst/></cfRule>"""
    val ev = run(withCf(column(num(1)), block("A1", CfRule.Preserved(payload, Some(4)))))
    assertEquals(ev.unevaluated.map(_.kind), Vector("dataBar"))
  }

  test("a block outside the window paints nothing and reports nothing") {
    val iconSet = CfRule.Preserved("""<cfRule type="iconSet" priority="2"/>""", Some(2))
    val sheet = withCf(column(num(1)), block("Z50:Z60", expr("TRUE"), iconSet))
    assertEquals(run(sheet, window = "A1:C3"), CfEvaluation(CfOverlay.empty, Vector.empty))
  }

  test("only window cells are painted; statistics still read the whole block") {
    val sheet = column((1 to 10).map(i => num(i.toDouble))*)
    val ev = run(withCf(sheet, block("A1:A10", top(3))), window = "A1:A8")
    assertEquals(painted(ev), Set("A8"), "8 is in the top 3 of 1..10 (9 and 10 lie outside)")
  }

  test("the --eval shape: a formula replaced by its live value paints by that value") {
    val stale = column(CellValue.Formula("B1*2", Some(num(0)))).put(aref("B1"), num(75))
    val rule = block("A1", cellIs(CfOperator.GreaterThan, "100"))
    assertEquals(painted(run(withCf(stale, rule))), Set.empty[String], "cached 0")
    val live = stale.put(aref("A1"), num(150))
    assertEquals(painted(run(withCf(live, rule))), Set("A1"))
  }

  test("evaluation is deterministic and its reports are sorted by priority") {
    val sheet = withCf(
      column(num(1), num(2)),
      block("A1:A2", expr("NOPE(", 5), CfRule.Preserved("<cfRule type=\"iconSet\"/>", Some(2))),
      block("A1:A2", expr("RAND()>0.5", 1), expr("BAD(", 3))
    )
    val once = run(sheet)
    assertEquals(run(sheet), once)
    assertEquals(once.unevaluated.flatMap(_.priority), Vector(2, 3, 5))
  }

  test("a report quotes at most 200 characters of a formula") {
    val long = "1+" * 300 + "("
    val ev = run(withCf(column(num(1)), block("A1", expr(long))))
    val message = ev.unevaluated.map(_.message).mkString
    assert(message.contains(s"'${"1+" * 100}…'"), message)
    assert(!message.contains(long) && message.length < 800, message)
  }

  test("conditionalFormatOverlay is the evaluation's overlay") {
    val sheet = withCf(column(num(1), num(9)), block("A1:A2", cellIs(CfOperator.GreaterThan, "5")))
    assertEquals(sheet.conditionalFormatOverlay(range("A1:A2")), run(sheet, "A1:A2").overlay)
  }

  // Bounded by the evaluator's own totality: a formula that makes it throw (the Weaver
  // SUMPRODUCT(--(r>0),ABS(r1+r2)) ClassCastException) escapes here too, and the CLI degrades it
  // to a CF_NOT_RENDERED warning (InMemorySource.conditionalPaint).
  property("totality: any expression text yields an evaluation, never an exception") {
    val genFormula = Gen.oneOf(
      FormulaTextGens.genRefText.map(r => s"$r>0"),
      FormulaTextGens.genRangeText.map(r => s"SUM($r)"),
      Gen.asciiPrintableStr
    )
    forAll(genFormula) { f =>
      val sheet = withCf(column(num(1), text("x"), CellValue.Empty), block("A1:B3", expr(f)))
      val ev = run(sheet, "A1:B3")
      assert(ev.overlay.cells.keySet.forall(r => range("A1:B3").contains(r)))
    }
  }
