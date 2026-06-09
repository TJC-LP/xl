package com.tjclp.xl

import munit.FunSuite
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.api.*
import com.tjclp.xl.dsl.syntax.*
import com.tjclp.xl.macros.ref
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.syntax.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.dsl.*

/**
 * Tests for `range.outlined(style[, color])` (GH-257): edge-correct outline around a range built
 * from per-cell [[Patch.MergeBorder]] patches that preserve existing cell styles.
 */
class OutlinedPatchSpec extends FunSuite:

  def emptySheet: Sheet = Sheet(SheetName.unsafe("Test"))

  private val thin = BorderSide(BorderStyle.Thin)

  /** Border of the cell at `ref`, or Border.none when the cell is unstyled. */
  private def borderAt(sheet: Sheet, cellRef: ARef): Border =
    sheet.getCellStyle(cellRef).map(_.border).getOrElse(Border.none)

  // ========== Geometry: 3x3 ==========

  test("outlined 3x3: corners get two sides, edges one, interior none") {
    val sheet = Patch.applyPatch(emptySheet, ref"B2:D4".outlined(BorderStyle.Thin))

    // Corners
    assertEquals(borderAt(sheet, ref"B2"), Border(left = thin, top = thin))
    assertEquals(borderAt(sheet, ref"D2"), Border(right = thin, top = thin))
    assertEquals(borderAt(sheet, ref"B4"), Border(left = thin, bottom = thin))
    assertEquals(borderAt(sheet, ref"D4"), Border(right = thin, bottom = thin))

    // Edge midpoints
    assertEquals(borderAt(sheet, ref"C2"), Border(top = thin))
    assertEquals(borderAt(sheet, ref"C4"), Border(bottom = thin))
    assertEquals(borderAt(sheet, ref"B3"), Border(left = thin))
    assertEquals(borderAt(sheet, ref"D3"), Border(right = thin))

    // Interior cell is untouched (no cell materialized at all)
    assertEquals(sheet.cells.get(ref"C3"), None)
  }

  test("outlined 3x3: exactly the 8 perimeter cells are patched") {
    val patch = ref"B2:D4".outlined(BorderStyle.Thin)
    patch match
      case Patch.Batch(patches) =>
        assertEquals(patches.size, 8)
        val refs = patches.collect { case Patch.MergeBorder(r, _) => r }
        assertEquals(refs.size, 8, "every patch should be a MergeBorder")
        assertEquals(refs.distinct.size, 8, "one patch per perimeter cell")
        assert(!refs.contains(ref"C3"), "interior cell must not be patched")
      case other => fail(s"Expected Batch, got: $other")
  }

  // ========== Geometry: degenerate shapes ==========

  test("outlined 1x3 row: every cell top+bottom, ends add left/right") {
    val sheet = Patch.applyPatch(emptySheet, ref"A1:C1".outlined(BorderStyle.Thin))

    assertEquals(borderAt(sheet, ref"A1"), Border(left = thin, top = thin, bottom = thin))
    assertEquals(borderAt(sheet, ref"B1"), Border(top = thin, bottom = thin))
    assertEquals(borderAt(sheet, ref"C1"), Border(right = thin, top = thin, bottom = thin))
  }

  test("outlined 3x1 column: every cell left+right, caps add top/bottom") {
    val sheet = Patch.applyPatch(emptySheet, ref"A1:A3".outlined(BorderStyle.Thin))

    assertEquals(borderAt(sheet, ref"A1"), Border(left = thin, right = thin, top = thin))
    assertEquals(borderAt(sheet, ref"A2"), Border(left = thin, right = thin))
    assertEquals(borderAt(sheet, ref"A3"), Border(left = thin, right = thin, bottom = thin))
  }

  test("outlined 1x1: all four sides (full box)") {
    val sheet = Patch.applyPatch(emptySheet, ref"B2:B2".outlined(BorderStyle.Medium))
    assertEquals(borderAt(sheet, ref"B2"), Border.all(BorderStyle.Medium))
  }

  test("outlined 2x2: every cell is a corner with exactly two sides") {
    val sheet = Patch.applyPatch(emptySheet, ref"A1:B2".outlined(BorderStyle.Thin))

    assertEquals(borderAt(sheet, ref"A1"), Border(left = thin, top = thin))
    assertEquals(borderAt(sheet, ref"B1"), Border(right = thin, top = thin))
    assertEquals(borderAt(sheet, ref"A2"), Border(left = thin, bottom = thin))
    assertEquals(borderAt(sheet, ref"B2"), Border(right = thin, bottom = thin))
  }

  // ========== Color overload ==========

  test("outlined with color applies the color to every emitted side") {
    val gray = Color.fromRgb(128, 128, 128)
    val side = BorderSide(BorderStyle.MediumDashed, gray)
    val sheet = Patch.applyPatch(emptySheet, ref"A1:B1".outlined(BorderStyle.MediumDashed, gray))

    assertEquals(borderAt(sheet, ref"A1"), Border(left = side, top = side, bottom = side))
    assertEquals(borderAt(sheet, ref"B1"), Border(right = side, top = side, bottom = side))
  }

  // ========== Existing style preservation ==========

  test("outlined preserves existing font, fill, and numFmt on edge cells") {
    val keyFigure = CellStyle.default.bold.bgYellow.currency
    val base = emptySheet
      .put(ref"B2", CellValue.Number(BigDecimal(42)))
      .withCellStyle(ref"B2", keyFigure)

    val sheet = Patch.applyPatch(base, ref"B2:D4".outlined(BorderStyle.Thin))
    val style = sheet.getCellStyle(ref"B2").getOrElse(fail("B2 should have a style"))

    assert(style.font.bold, "bold preserved")
    assertEquals(style.fill, keyFigure.fill, "fill preserved")
    assertEquals(style.numFmt, keyFigure.numFmt, "numFmt preserved")
    assertEquals(style.border, Border(left = thin, top = thin), "outline sides applied")
    assertEquals(
      sheet.cells.get(ref"B2").map(_.value),
      Some(CellValue.Number(BigDecimal(42))),
      "cell value untouched"
    )
  }

  test("outlined preserves existing border sides it does not set") {
    // B2 sits on the top-left corner of the outline: outline sets top+left only,
    // so a pre-existing bottom rule must survive.
    val ruled = CellStyle.default.borderBottom(BorderStyle.Double)
    val base = emptySheet.withCellStyle(ref"B2", ruled)

    val sheet = Patch.applyPatch(base, ref"B2:D4".outlined(BorderStyle.Thin))

    assertEquals(
      borderAt(sheet, ref"B2"),
      Border(left = thin, top = thin, bottom = BorderSide(BorderStyle.Double))
    )
  }

  test("outlined overrides the sides it does set") {
    val boxed = CellStyle.default.bordered // thin on all four sides
    val base = emptySheet.withCellStyle(ref"A1", boxed)

    val sheet = Patch.applyPatch(base, ref"A1:B2".outlined(BorderStyle.Thick))

    // A1 is the top-left corner: top+left become thick, right+bottom keep the existing thin
    assertEquals(
      borderAt(sheet, ref"A1"),
      Border(
        left = BorderSide(BorderStyle.Thick),
        right = thin,
        top = BorderSide(BorderStyle.Thick),
        bottom = thin
      )
    )
  }

  // ========== Composition and idempotence ==========

  test("outlined composes with other patches via ++") {
    val patch = (ref"A1" := "Total") ++ ref"A1:C1".outlined(BorderStyle.Thin)
    val sheet = Patch.applyPatch(emptySheet, patch)

    assertEquals(sheet.cells.get(ref"A1").map(_.value), Some(CellValue.Text("Total")))
    assertEquals(borderAt(sheet, ref"A1"), Border(left = thin, top = thin, bottom = thin))
  }

  test("outlined is idempotent (applying twice equals applying once)") {
    val patch = ref"B2:D4".outlined(BorderStyle.Thin)
    val once = Patch.applyPatch(emptySheet, patch)
    val twice = Patch.applyPatch(once, patch)

    assertEquals(twice.cells, once.cells)
    ref"B2:D4".cells.foreach { r =>
      assertEquals(twice.getCellStyle(r), once.getCellStyle(r))
    }
  }

  // ========== MergeBorder patch semantics ==========

  test("MergeBorder on an unstyled cell creates a border-only style") {
    val patch = Patch.MergeBorder(ref"A1", Border(top = thin))
    val sheet = Patch.applyPatch(emptySheet, patch)
    val style = sheet.getCellStyle(ref"A1").getOrElse(fail("A1 should have a style"))
    assertEquals(style, CellStyle.default.withBorder(Border(top = thin)))
  }

  test("MergeBorder with Border.none leaves the existing style intact") {
    val base = emptySheet.withCellStyle(ref"A1", CellStyle.default.bold.bordered)
    val sheet = Patch.applyPatch(base, Patch.MergeBorder(ref"A1", Border.none))
    assertEquals(sheet.getCellStyle(ref"A1"), base.getCellStyle(ref"A1"))
  }
