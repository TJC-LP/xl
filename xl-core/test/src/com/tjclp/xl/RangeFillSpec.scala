package com.tjclp.xl

import munit.ScalaCheckSuite
import org.scalacheck.Prop.*
import Generators.given
import Generators.genCellRange
import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.dsl.syntax.*
import com.tjclp.xl.macros.ref
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.syntax.*

/**
 * Properties for range fill semantics (`range := value`, Excel Ctrl+Enter) and total ARef
 * navigation, added in 0.11.0. Before 0.11.0, `:=` on a RefType range silently returned
 * Patch.empty and CellRange had no `:=` at all.
 */
class RangeFillSpec extends ScalaCheckSuite:

  def emptySheet: Sheet = Sheet(SheetName.unsafe("Test"))

  // genCellRange builds from genSmallARef (cols 0-25, rows 0-99), so size is bounded.

  property("range := v puts v in every cell of the range") {
    forAll(genCellRange) { (range: CellRange) =>
      val sheet = emptySheet.put(range := 42)
      assertEquals(sheet.cells.size, range.size)
      assert(range.cells.forall { r =>
        sheet.cells.get(r).map(_.value).contains(CellValue.Number(BigDecimal(42)))
      })
    }
  }

  property("fill on a 1x1 range is equivalent to a single-cell Put") {
    forAll { (aref: ARef) =>
      val single = CellRange(aref, aref)
      val viaFill = emptySheet.put(single := "x")
      val viaPut = emptySheet.put(aref := "x")
      assertEquals(viaFill.cells, viaPut.cells)
    }
  }

  property("fill composes with the Patch monoid: later fill wins on overlap") {
    forAll(genCellRange) { (range: CellRange) =>
      val sheet = emptySheet.put((range := 1) ++ (range := 2))
      assert(range.cells.forall { r =>
        sheet.cells.get(r).map(_.value).contains(CellValue.Number(BigDecimal(2)))
      })
    }
  }

  test("RefType range := fills (runtime-interpolated refs)") {
    val end = "3"
    val patch = (for refType <- ref"A1:B$end" yield refType := true).getOrElse(Patch.empty)
    val sheet = emptySheet.put(patch)
    assertEquals(sheet.cells.size, 6)
    assert(sheet.cells.values.forall(_.value == CellValue.Bool(true)))
  }

  test("compile-time range literal supports := directly") {
    val sheet = emptySheet.put(ref"A1:C2" := 0)
    assertEquals(sheet.cells.size, 6)
  }

  test("range := accepts all conversion overloads") {
    val r = ref"A1:A2"
    val patches = List(
      r := CellValue.Text("cv"),
      r := "s",
      r := 1,
      r := 1L,
      r := 1.5,
      r := BigDecimal(2),
      r := false,
      r := java.time.LocalDateTime.of(2026, 6, 9, 0, 0)
    )
    patches.foreach(p => assertEquals(emptySheet.put(p).cells.size, 2))
  }

  // ========== ARef navigation ==========

  property("down/up and right/left are inverses (within bounds)") {
    forAll(Generators.genSmallARef) { (aref: ARef) =>
      forAll(org.scalacheck.Gen.choose(0, 50)) { (n: Int) =>
        assertEquals(aref.down(n).up(n), aref)
        assertEquals(aref.right(n).left(n), aref)
      }
    }
  }

  test("navigation defaults move by one") {
    assertEquals(ref"B2".down(), ref"B3")
    assertEquals(ref"B2".up(), ref"B1")
    assertEquals(ref"B2".right(), ref"C2")
    assertEquals(ref"B2".left(), ref"A2")
    assertEquals(ref"A1".down(3).right(2), ref"C4")
  }

  test("large range fill: one cell per address, by design (cost is range.size)"):
    // Pins the documented contract for big fills: ref"A1:D2500" materializes 10,000 cells.
    // Full-column fills (A:A = 1,048,576 cells) follow the same rule — prefer bounded ranges.
    val sheet = emptySheet.put(ref"A1:D2500" := 0)
    assertEquals(sheet.cells.size, 10000)
