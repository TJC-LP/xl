package com.tjclp.xl.cf

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.styles.color.Color
import munit.FunSuite

/**
 * GH-136: typed conditional-formatting model — constructors, Sheet authoring API, priority
 * allocation (auto above every existing priority, incl. Preserved ones), and structural shifts.
 */
class ConditionalFormatSpec extends FunSuite:

  private val red = Color.Rgb(0xffff0000)
  private val green = Color.Rgb(0xff00ff00)
  private val dxf = Dxf.fill(red)

  // ===== constructors =====

  test("constructors strip a leading '=' from formula text") {
    CfRule.cellIs(CfOperator.GreaterThan, "=100", dxf) match
      case CfRule.CellIs(op, f1, f2, d, _, stop) =>
        assertEquals(op, CfOperator.GreaterThan)
        assertEquals(f1, "100")
        assertEquals(f2, None)
        assertEquals(d, Some(dxf))
        assertEquals(stop, false)
      case other => fail(s"expected CellIs, got $other")
    CfRule.expression("=$B1>$C1", dxf) match
      case CfRule.Expression(f, _, _, _) => assertEquals(f, "$B1>$C1")
      case other => fail(s"expected Expression, got $other")
    CfRule.between("=1", "=9", dxf) match
      case CfRule.CellIs(CfOperator.Between, f1, f2, _, _, _) =>
        assertEquals((f1, f2), ("1", Some("9")))
      case other => fail(s"expected Between CellIs, got $other")
  }

  test("constructors default to the auto-priority sentinel") {
    assertEquals(CfRule.priorityOf(CfRule.cellIs(CfOperator.LessThan, "0", dxf)), Some(0))
    assertEquals(CfRule.priorityOf(CfRule.dataBar(green)), Some(0))
    assertEquals(
      CfRule.priorityOf(CfRule.colorScale2(CfPoint(Cfvo.Min, red), CfPoint(Cfvo.Max, green))),
      Some(0)
    )
    assertEquals(CfRule.priorityOf(CfRule.Preserved("<cfRule/>", None)), None)
  }

  // ===== Sheet authoring =====

  test("conditionalFormat appends ONE Rules block and assigns priorities 1..n") {
    val s = Sheet("CF").conditionalFormat(
      ref"A1:A5",
      CfRule.cellIs(CfOperator.GreaterThan, "100", dxf),
      CfRule.containsText("x", dxf)
    )
    assertEquals(s.conditionalFormats.size, 1)
    s.conditionalFormats(0) match
      case ConditionalFormat.Rules(ranges, rules, pivot) =>
        assertEquals(ranges, Vector(ref"A1:A5": CellRange))
        assertEquals(rules.map(CfRule.priorityOf), Vector(Some(1), Some(2)))
        assertEquals(pivot, false)
      case other => fail(s"expected Rules, got $other")
  }

  test("second block allocates above the first; explicit positives pass through") {
    val s = Sheet("CF")
      .conditionalFormat(ref"A1:A5", CfRule.cellIs(CfOperator.GreaterThan, "1", dxf))
      .conditionalFormat(
        Vector(ref"B1:B5": CellRange),
        Vector(
          CfRule.expression("B1>0", dxf),
          CfRule.CellIs(CfOperator.LessThan, "0", None, Some(dxf), priority = 7),
          CfRule.dataBar(green)
        )
      )
    val priorities = s.typedConditionalFormats.flatMap(_.rules.map(CfRule.priorityOf))
    // block 1: auto -> 1; block 2: auto -> 2, explicit 7 untouched, auto -> 3
    assertEquals(priorities, Vector(Some(1), Some(2), Some(7), Some(3)))
  }

  test("allocation goes above rule-Preserved parsed priorities") {
    val preservedRule = CfRule.Preserved("<cfRule type=\"iconSet\" priority=\"9\"/>", Some(9))
    val s = Sheet("CF")
      .copy(conditionalFormats =
        Vector(ConditionalFormat.Rules(Vector(ref"A1:A2": CellRange), Vector(preservedRule)))
      )
      .conditionalFormat(ref"B1:B2", CfRule.cellIs(CfOperator.Equal, "1", dxf))
    assertEquals(
      s.typedConditionalFormats.flatMap(_.rules).flatMap(CfRule.priorityOf),
      Vector(9, 10)
    )
  }

  test("allocation goes above block-Preserved scanned priorities") {
    val block = ConditionalFormat.Preserved(
      """<conditionalFormatting sqref="A1" bogus="1"><cfRule type="iconSet" priority="41"/></conditionalFormatting>"""
    )
    val s = Sheet("CF")
      .copy(conditionalFormats = Vector(block))
      .conditionalFormat(ref"B1:B2", CfRule.cellIs(CfOperator.Equal, "1", dxf))
    assertEquals(s.typedConditionalFormats.flatMap(_.rules).flatMap(CfRule.priorityOf), Vector(42))
  }

  test("scanPriorities is total on junk and overflow") {
    assertEquals(ConditionalFormat.scanPriorities("no priorities here"), Vector.empty)
    assertEquals(ConditionalFormat.scanPriorities("""priority="3" priority="7""""), Vector(3, 7))
    // overflow guarded by toIntOption
    assertEquals(ConditionalFormat.scanPriorities("""priority="99999999999999""""), Vector.empty)
  }

  test("auto-priority allocation saturates at Int.MaxValue (never wraps negative)") {
    val s = Sheet("CF")
      .conditionalFormat(
        Vector(ref"A1:A5": CellRange),
        Vector(CfRule.CellIs(CfOperator.LessThan, "0", None, Some(dxf), priority = Int.MaxValue))
      )
      .conditionalFormat(
        ref"B1:B5",
        CfRule.cellIs(CfOperator.GreaterThan, "1", dxf),
        CfRule.dataBar(green)
      )
    // a wrapped Int.MinValue is schema-invalid; saturation collides at Int.MaxValue instead
    // (colliding priorities are Excel-tolerated, negative ones are not)
    assertEquals(
      s.typedConditionalFormats.flatMap(_.rules).flatMap(CfRule.priorityOf),
      Vector(Int.MaxValue, Int.MaxValue, Int.MaxValue)
    )
  }

  test("removeConditionalFormat mirrors removeDrawing (identity out of range)") {
    val s = Sheet("CF")
      .conditionalFormat(ref"A1:A5", CfRule.cellIs(CfOperator.GreaterThan, "1", dxf))
      .conditionalFormat(ref"B1:B5", CfRule.expression("B1>0", dxf))
    assertEquals(s.removeConditionalFormat(0).conditionalFormats.size, 1)
    assertEquals(s.removeConditionalFormat(5), s)
    s.removeConditionalFormat(0).conditionalFormats(0) match
      case ConditionalFormat.Rules(ranges, _, _) =>
        assertEquals(ranges, Vector(ref"B1:B5": CellRange))
      case other => fail(s"expected Rules, got $other")
  }

  test("typedConditionalFormats excludes block-Preserved fragments") {
    val s = Sheet("CF")
      .copy(conditionalFormats = Vector(ConditionalFormat.Preserved("<conditionalFormatting/>")))
      .conditionalFormat(ref"A1:A2", CfRule.cellIs(CfOperator.Equal, "1", dxf))
    assertEquals(s.conditionalFormats.size, 2)
    assertEquals(s.typedConditionalFormats.size, 1)
  }

  // ===== structural shifts (the mergedRanges clamp/split/drop algebra) =====

  private def blockRanges(s: Sheet): Vector[Vector[String]] =
    s.typedConditionalFormats.map(_.ranges.map(_.toA1))

  test("insertRows shifts cf ranges below the cut") {
    val s = Sheet("CF").conditionalFormat(ref"A2:A4", CfRule.cellIs(CfOperator.Equal, "1", dxf))
    assertEquals(blockRanges(s.insertRows(at = 0, count = 2)), Vector(Vector("A4:A6")))
  }

  test("deleteRows clamps a cf range spanning the cut") {
    val s = Sheet("CF").conditionalFormat(ref"A1:A4", CfRule.cellIs(CfOperator.Equal, "1", dxf))
    assertEquals(blockRanges(s.deleteRows(at = 1, count = 2)), Vector(Vector("A1:A2")))
  }

  test("deleteColumns drops a fully-deleted range and removes an empty block") {
    val s = Sheet("CF")
      .conditionalFormat(ref"B1:B5", CfRule.cellIs(CfOperator.Equal, "1", dxf))
      .conditionalFormat(
        Vector(ref"B2:B3": CellRange, ref"D2:D3": CellRange),
        Vector(CfRule.expression("B2>0", dxf))
      )
    val r = s.deleteColumns(at = 1, count = 1) // delete column B
    // first block: only range fully deleted -> whole block removed
    // second block: B2:B3 dropped, D2:D3 -> C2:C3
    assertEquals(blockRanges(r), Vector(Vector("C2:C3")))
  }

  test("structural shifts leave Preserved blocks untouched and shift Preserved rules' envelopes") {
    val preservedBlock = ConditionalFormat.Preserved("<conditionalFormatting sqref=\"Z9\"/>")
    val s = Sheet("CF")
      .copy(conditionalFormats = Vector(preservedBlock))
      .conditionalFormat(
        Vector(ref"A2:A4": CellRange),
        Vector(CfRule.Preserved("<cfRule type=\"iconSet\" priority=\"1\"/>", Some(1)))
      )
    val r = s.insertRows(at = 0, count = 1)
    assertEquals(r.conditionalFormats(0), preservedBlock)
    r.conditionalFormats(1) match
      case ConditionalFormat.Rules(ranges, rules, _) =>
        assertEquals(ranges.map(_.toA1), Vector("A3:A5"))
        assertEquals(
          rules,
          Vector(CfRule.Preserved("<cfRule type=\"iconSet\" priority=\"1\"/>", Some(1)))
        )
      case other => fail(s"expected Rules, got $other")
  }

  // ===== Dxf builders =====

  test("Dxf builders populate only the requested slots") {
    assertEquals(Dxf.fill(red), Dxf(fill = Some(com.tjclp.xl.styles.fill.Fill.Solid(red))))
    val f = DxfFont(bold = Some(true), color = Some(green))
    assertEquals(Dxf.font(f), Dxf(font = Some(f)))
    assertEquals(
      Dxf.fillAndFont(red, f),
      Dxf(font = Some(f), fill = Some(com.tjclp.xl.styles.fill.Fill.Solid(red)))
    )
  }
