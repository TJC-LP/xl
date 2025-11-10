package com.tjclp.xl

import munit.ScalaCheckSuite
import org.scalacheck.Prop.*
import cats.syntax.all.*
import Generators.given
import Patch.{given, *}
import com.tjclp.xl.style.{CellStyle, StyleId, Font, Fill, Color}

/** Property tests for Patch monoid laws and semantics */
class PatchSpec extends ScalaCheckSuite:

  // Helper to create a simple test sheet
  def emptySheet: Sheet = Sheet(SheetName.unsafe("Test"))

  // ========== Monoid Laws ==========

  property("Patch Monoid: left identity (empty |+| p == p)") {
    forAll { (ref: ARef, value: CellValue) =>
      val p = Patch.Put(ref, value)
      val combined = Patch.empty |+| p

      // Should produce the same patch (semantically)
      combined match
        case Patch.Batch(Vector(Patch.Put(r, v))) =>
          assertEquals(r, ref)
          assertEquals(v, value)
        case Patch.Put(r, v) =>
          assertEquals(r, ref)
          assertEquals(v, value)
        case _ => fail(s"Expected Put or single-element Batch, got: $combined")

      true
    }
  }

  property("Patch Monoid: right identity (p |+| empty == p)") {
    forAll { (ref: ARef, value: CellValue) =>
      val p = Patch.Put(ref, value)
      val combined = p |+| Patch.empty

      // Should produce the same patch (semantically)
      combined match
        case Patch.Batch(Vector(Patch.Put(r, v))) =>
          assertEquals(r, ref)
          assertEquals(v, value)
        case Patch.Put(r, v) =>
          assertEquals(r, ref)
          assertEquals(v, value)
        case _ => fail(s"Expected Put or single-element Batch, got: $combined")

      true
    }
  }

  property("Patch Monoid: associativity ((p1 |+| p2) |+| p3 == p1 |+| (p2 |+| p3))") {
    forAll { (r1: ARef, r2: ARef, r3: ARef, v1: CellValue, v2: CellValue, v3: CellValue) =>
      val p1 = Patch.Put(r1, v1)
      val p2 = Patch.Put(r2, v2)
      val p3 = Patch.Put(r3, v3)

      val left = (p1 |+| p2) |+| p3
      val right = p1 |+| (p2 |+| p3)

      // Apply both to the same sheet and verify results are identical
      val sheet = emptySheet
      val leftResult = Patch.applyPatch(sheet, left)
      val rightResult = Patch.applyPatch(sheet, right)

      (leftResult, rightResult) match
        case (Right(s1), Right(s2)) =>
          assertEquals(s1.cells, s2.cells)
        case _ => fail("Both should succeed")

      true
    }
  }

  // ========== Patch Application Tests ==========

  test("Put patch adds a cell") {
    val sheet = emptySheet
    val ref = ARef.from1(1, 1) // A1
    val value = CellValue.Text("Hello")
    val patch = Patch.Put(ref, value)

    val result = Patch.applyPatch(sheet, patch)
    assert(result.isRight)

    val updated = result.getOrElse(fail("Should succeed"))
    assertEquals(updated(ref).value, value)
  }

  test("SetStyle patch sets cell style") {
    val sheet = emptySheet
    val ref = ARef.from1(1, 1)
    val styleId = StyleId(42)

    // First put a cell, then style it
    val patch = (Patch.Put(ref, CellValue.Text("Test")): Patch) |+| (Patch.SetStyle(ref, styleId): Patch)
    val result = Patch.applyPatch(sheet, patch)

    assert(result.isRight)
    val updated = result.getOrElse(fail("Should succeed"))
    assertEquals(updated(ref).styleId, Some(styleId))
  }

  test("ClearStyle patch removes cell style") {
    val sheet = emptySheet
      .put(ARef.from1(1, 1), CellValue.Text("Test"))
      .put(Cell(ARef.from1(1, 1), CellValue.Text("Test"), Some(StyleId(42))))

    val ref = ARef.from1(1, 1)
    val patch = Patch.ClearStyle(ref)
    val result = Patch.applyPatch(sheet, patch)

    assert(result.isRight)
    val updated = result.getOrElse(fail("Should succeed"))
    assertEquals(updated(ref).styleId, None)
  }

  test("SetCellStyle patch applies CellStyle object") {
    val boldStyle = CellStyle.default.withFont(Font("Arial", 14.0, bold = true))
    val sheet = emptySheet.put(cell"A1", CellValue.Text("Header"))

    val patch = Patch.SetCellStyle(cell"A1", boldStyle)
    val result = sheet.applyPatch(patch)

    assert(result.isRight)
    result.foreach { updated =>
      // Style should be registered
      assertEquals(updated.styleRegistry.size, 2) // default + bold

      // Cell should have styleId
      assert(updated(cell"A1").styleId.isDefined)

      // Can retrieve original style
      assertEquals(updated.getCellStyle(cell"A1"), Some(boldStyle))
    }
  }

  test("SetCellStyle deduplicates identical styles") {
    val redStyle = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xFFFF0000)))
    val sheet = emptySheet
      .put(cell"A1", CellValue.Text("Red1"))
      .put(cell"A2", CellValue.Text("Red2"))

    val patch =
      (Patch.SetCellStyle(cell"A1", redStyle): Patch) |+|
        (Patch.SetCellStyle(cell"A2", redStyle): Patch)

    val result = sheet.applyPatch(patch)
    result.foreach { updated =>
      // Should only have 2 styles (default + red)
      assertEquals(updated.styleRegistry.size, 2)

      // Both cells should reference same style index
      assertEquals(
        updated(cell"A1").styleId,
        updated(cell"A2").styleId
      )
    }
  }

  test("Merge patch adds merged range") {
    val sheet = emptySheet
    val range = CellRange(ARef.from1(1, 1), ARef.from1(2, 2))
    val patch = Patch.Merge(range)

    val result = Patch.applyPatch(sheet, patch)
    assert(result.isRight)

    val updated = result.getOrElse(fail("Should succeed"))
    assert(updated.mergedRanges.contains(range))
  }

  test("Unmerge patch removes merged range") {
    val range = CellRange(ARef.from1(1, 1), ARef.from1(2, 2))
    val sheet = emptySheet.merge(range)
    val patch = Patch.Unmerge(range)

    val result = Patch.applyPatch(sheet, patch)
    assert(result.isRight)

    val updated = result.getOrElse(fail("Should succeed"))
    assert(!updated.mergedRanges.contains(range))
  }

  test("Remove patch removes a cell") {
    val ref = ARef.from1(1, 1)
    val sheet = emptySheet.put(ref, CellValue.Text("Test"))
    val patch = Patch.Remove(ref)

    val result = Patch.applyPatch(sheet, patch)
    assert(result.isRight)

    val updated = result.getOrElse(fail("Should succeed"))
    assert(!updated.contains(ref))
  }

  test("SetColumnProperties patch sets column properties") {
    val sheet = emptySheet
    val col = Column.from0(0) // Column A
    val props = ColumnProperties(width = Some(20.0), hidden = false)
    val patch = Patch.SetColumnProperties(col, props)

    val result = Patch.applyPatch(sheet, patch)
    assert(result.isRight)

    val updated = result.getOrElse(fail("Should succeed"))
    assertEquals(updated.getColumnProperties(col), props)
  }

  test("SetRowProperties patch sets row properties") {
    val sheet = emptySheet
    val row = Row.from0(0) // Row 1
    val props = RowProperties(height = Some(30.0), hidden = true)
    val patch = Patch.SetRowProperties(row, props)

    val result = Patch.applyPatch(sheet, patch)
    assert(result.isRight)

    val updated = result.getOrElse(fail("Should succeed"))
    assertEquals(updated.getRowProperties(row), props)
  }

  test("Batch patch applies multiple patches in order") {
    val sheet = emptySheet
    val r1 = ARef.from1(1, 1)
    val r2 = ARef.from1(2, 1)

    val patches = Vector(
      Patch.Put(r1, CellValue.Text("A1")),
      Patch.Put(r2, CellValue.Number(42)),
      Patch.SetStyle(r1, StyleId(1))
    )

    val batch = Patch.Batch(patches)
    val result = Patch.applyPatch(sheet, batch)

    assert(result.isRight)
    val updated = result.getOrElse(fail("Should succeed"))

    assertEquals(updated(r1).value, CellValue.Text("A1"))
    assertEquals(updated(r1).styleId, Some(StyleId(1)))
    assertEquals(updated(r2).value, CellValue.Number(42))
  }

  // ========== Idempotence Tests ==========

  property("Put patch is idempotent (applying twice yields same result)") {
    forAll { (ref: ARef, value: CellValue) =>
      val sheet = emptySheet
      val patch = Patch.Put(ref, value)

      val once = Patch.applyPatch(sheet, patch).getOrElse(fail("Should succeed"))
      val twice = Patch.applyPatch(once, patch).getOrElse(fail("Should succeed"))

      assertEquals(once.cells, twice.cells)
      true
    }
  }

  property("SetStyle patch is idempotent") {
    forAll { (ref: ARef, styleId: Int) =>
      val sheet = emptySheet.put(ref, CellValue.Empty)
      val patch = Patch.SetStyle(ref, StyleId(styleId))

      val once = Patch.applyPatch(sheet, patch).getOrElse(fail("Should succeed"))
      val twice = Patch.applyPatch(once, patch).getOrElse(fail("Should succeed"))

      assertEquals(once.cells, twice.cells)
      true
    }
  }

  property("Merge patch is idempotent") {
    forAll { (range: CellRange) =>
      val sheet = emptySheet
      val patch = Patch.Merge(range)

      val once = Patch.applyPatch(sheet, patch).getOrElse(fail("Should succeed"))
      val twice = Patch.applyPatch(once, patch).getOrElse(fail("Should succeed"))

      assertEquals(once.mergedRanges, twice.mergedRanges)
      true
    }
  }

  // ========== Override Semantics Tests ==========

  test("Later Put patch overrides earlier Put to same cell") {
    val sheet = emptySheet
    val ref = ARef.from1(1, 1)

    val patch = (Patch.Put(ref, CellValue.Text("First")): Patch) |+| (Patch.Put(ref, CellValue.Text("Second")): Patch)
    val result = Patch.applyPatch(sheet, patch)

    assert(result.isRight)
    val updated = result.getOrElse(fail("Should succeed"))
    assertEquals(updated(ref).value, CellValue.Text("Second"))
  }

  test("Later SetStyle overrides earlier SetStyle to same cell") {
    val sheet = emptySheet.put(ARef.from1(1, 1), CellValue.Empty)
    val ref = ARef.from1(1, 1)

    val patch = (Patch.SetStyle(ref, StyleId(1)): Patch) |+| (Patch.SetStyle(ref, StyleId(2)): Patch)
    val result = Patch.applyPatch(sheet, patch)

    assert(result.isRight)
    val updated = result.getOrElse(fail("Should succeed"))
    assertEquals(updated(ref).styleId, Some(StyleId(2)))
  }

  // ========== Extension Method Tests ==========

  test("Sheet.applyPatch extension method works") {
    val sheet = emptySheet
    val ref = ARef.from1(1, 1)
    val patch = Patch.Put(ref, CellValue.Text("Test"))

    val result = sheet.applyPatch(patch)
    assert(result.isRight)

    val updated = result.getOrElse(fail("Should succeed"))
    assertEquals(updated(ref).value, CellValue.Text("Test"))
  }

  test("Sheet.applyPatches extension method works with varargs") {
    val sheet = emptySheet
    val r1 = ARef.from1(1, 1)
    val r2 = ARef.from1(2, 1)

    val result = sheet.applyPatches(
      Patch.Put(r1, CellValue.Text("A1")),
      Patch.Put(r2, CellValue.Number(42))
    )

    assert(result.isRight)
    val updated = result.getOrElse(fail("Should succeed"))
    assertEquals(updated(r1).value, CellValue.Text("A1"))
    assertEquals(updated(r2).value, CellValue.Number(42))
  }
