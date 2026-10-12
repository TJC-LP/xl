package com.tjclp.xl.ooxml

import munit.FunSuite

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{ArrayMode, FormulaKind}
import com.tjclp.xl.macros.ref

/** GH-714: the codec's dynamic-array arms — `cm` resolution, `cm` rendering, the bare 1x1 ref. */
class FormulaKindCodecSpec extends FunSuite:

  private val sortRange = CellRange(ref"B1", ref"B5")
  private val legacy: FormulaKind.ArrayFormula =
    FormulaKind.dynamicArray(sortRange).copy(mode = ArrayMode.Legacy)
  private val idx = CellMetadataIndex(
    Map(1 -> ArrayMode.Dynamic(), 3 -> ArrayMode.Dynamic(collapsed = true))
  )

  test("a cm that resolves upgrades an array record to a dynamic array") {
    assertEquals(
      FormulaKindCodec.withCellMetadata(Some(legacy), Some("1"), idx),
      Some(legacy.copy(mode = ArrayMode.Dynamic()))
    )
    assertEquals(
      FormulaKindCodec.withCellMetadata(Some(legacy), Some(" 3 "), idx),
      Some(legacy.copy(mode = ArrayMode.Dynamic(collapsed = true)))
    )
  }

  test("a cm that does not resolve leaves the record legacy") {
    List(None, Some("0"), Some("2"), Some("junk"), Some("-1"), Some("")).foreach { cm =>
      assertEquals(FormulaKindCodec.withCellMetadata(Some(legacy), cm, idx), Some(legacy), cm)
    }
    assertEquals(
      FormulaKindCodec.withCellMetadata(Some(legacy), Some("1"), CellMetadataIndex.empty),
      Some(legacy)
    )
  }

  test("a cm on a record that is not an array changes nothing") {
    val dt = FormulaKind.DataTable(sortRange, dt2D = false, dtr = false, Some(ref"A1"), None)
    List[Option[FormulaKind]](
      None,
      Some(FormulaKind.Normal()),
      Some(FormulaKind.Normal(ca = true)),
      Some(dt)
    )
      .foreach { kind =>
        assertEquals(FormulaKindCodec.withCellMetadata(kind, Some("1"), idx), kind)
      }
  }

  test("cellAttrs: cm for a dynamic record the plan allocated, nothing otherwise") {
    val cmOf: ArrayMode.Dynamic => Option[Int] = {
      case ArrayMode.Dynamic(false) => Some(1)
      case ArrayMode.Dynamic(true) => Some(2)
    }
    assertEquals(
      FormulaKindCodec.cellAttrs(FormulaKind.dynamicArray(sortRange), cmOf),
      List("cm" -> "1")
    )
    assertEquals(
      FormulaKindCodec.cellAttrs(legacy.copy(mode = ArrayMode.Dynamic(collapsed = true)), cmOf),
      List("cm" -> "2")
    )
    assertEquals(FormulaKindCodec.cellAttrs(legacy, cmOf), Nil)
    assertEquals(FormulaKindCodec.cellAttrs(FormulaKind.Normal(), cmOf), Nil)
    // no metadata part to point at: written in the legacy shape, never with a dangling cm
    assertEquals(
      FormulaKindCodec
        .cellAttrs(FormulaKind.dynamicArray(sortRange), FormulaKindCodec.noCellMetadata),
      Nil
    )
  }

  test("a 1x1 dynamic array renders a bare ref; the CSE pin keeps the range form") {
    val one = CellRange(ref"C5", ref"C5")
    assertEquals(
      FormulaKindCodec.toAttrs(FormulaKind.dynamicArray(one)),
      List("t" -> "array", "ref" -> "C5")
    )
    assertEquals(
      FormulaKindCodec.toAttrs(FormulaKind.ArrayFormula(CellRange(ref"E10", ref"E10"))),
      List("t" -> "array", "ref" -> "E10:E10")
    )
    assertEquals(
      FormulaKindCodec.toAttrs(FormulaKind.dynamicArray(sortRange)),
      List("t" -> "array", "ref" -> "B1:B5")
    )
  }

  test("isDynamicArray holds only for a dynamic record") {
    assert(FormulaKind.dynamicArray(sortRange).isDynamicArray)
    assert(!legacy.isDynamicArray)
    assert(!FormulaKind.Normal().isDynamicArray)
    assertEquals(FormulaKind.dynamicArray(sortRange).ref.start, ARef.from0(1, 0))
  }
