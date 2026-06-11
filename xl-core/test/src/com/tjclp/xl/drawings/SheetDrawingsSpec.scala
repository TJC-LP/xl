package com.tjclp.xl.drawings

import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.forAll

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.Comment
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.units.Emu

/** Sheet drawing API: addImage overloads, pictures, removeDrawing, shiftAxis remap (GH-221). */
class SheetDrawingsSpec extends ScalaCheckSuite:

  private val png = ImageData(TestImages.png2x3, ImageFormat.Png)
  private val wmf = ImageData(TestImages.wmfHeader, ImageFormat.Wmf)
  private val extent = Extent(Emu(95250L), Emu(95250L))
  private val sheet = Sheet(com.tjclp.xl.addressing.SheetName.unsafe("Pics"))

  test("addImage(image, anchor) appends in z-order") {
    val s = sheet
      .addImage(png, DrawingAnchor.at(ref"B2", extent))
      .addImage(png, DrawingAnchor.Absolute(Emu(0), Emu(0), extent))
    assertEquals(s.drawings.size, 2)
    assertEquals(s.pictures.size, 2)
    s.drawings(0) match
      case Drawing.Picture(DrawingAnchor.OneCell(from, e), img, _, _) =>
        assertEquals(from.cell, ref"B2")
        assertEquals(e, extent)
        assertEquals(img, png)
      case other => fail(s"unexpected drawing $other")
  }

  test("addImage(image, at) sizes naturally and fails for unsniffable formats") {
    val ok = sheet.addImage(png, ref"C3")
    assert(ok.isRight)
    ok.foreach { s =>
      s.drawings(0) match
        case Drawing.Picture(DrawingAnchor.OneCell(from, e), _, _, _) =>
          assertEquals(from.cell, ref"C3")
          assertEquals(e, Extent.fromPx(2, 3))
        case other => fail(s"unexpected drawing $other")
    }
    assert(sheet.addImage(wmf, ref"C3").isLeft, "wmf has no sniffable dimensions")
  }

  test("addImage(image, range, editAs) builds a two-cell anchor one past the range end") {
    val s = sheet.addImage(png, ref"B2:D5", EditAs.OneCell)
    s.drawings(0) match
      case Drawing.Picture(DrawingAnchor.TwoCell(from, to, editAs), _, _, _) =>
        assertEquals(from.cell, ref"B2")
        assertEquals(to.cell, ref"E6") // one past D5
        assertEquals(editAs, EditAs.OneCell)
      case other => fail(s"unexpected drawing $other")
  }

  test("DrawingAnchor.over clamps at the grid edge") {
    val last = ARef.from0(Column.MaxIndex0, Row.MaxIndex0)
    DrawingAnchor.over(CellRange(last, last)) match
      case DrawingAnchor.TwoCell(from, to, _) =>
        assertEquals(from.cell, last)
        assertEquals(to.cell, last) // clamped, degenerate edge documented
      case other => fail(s"unexpected anchor $other")
  }

  test("removeDrawing removes by index and is identity out of range") {
    val s = sheet
      .addImage(png, ref"A1", extent)
      .addImage(png, ref"B2", extent)
    assertEquals(s.removeDrawing(0).drawings.size, 1)
    assertEquals(s.removeDrawing(2), s)
    assertEquals(s.removeDrawing(-1), s)
    assertEquals(sheet.removeDrawing(0), sheet)
  }

  test("pictures excludes Preserved fragments but drawings keeps document order") {
    val s = sheet
      .addImage(png, ref"A1", extent)
      .copy(drawings = sheet.addImage(png, ref"A1", extent).drawings :+ Drawing.Preserved("<sp/>"))
    assertEquals(s.drawings.size, 2)
    assertEquals(s.pictures.size, 1)
  }

  test("insertRows shifts one-cell and two-cell anchors below the insertion point") {
    val s = sheet
      .addImage(png, ref"B5", extent)
      .addImage(png, ref"C2:D3", EditAs.TwoCell)
      .insertRows(at = 2, count = 3) // insert above row index 2 (row 3 in A1 terms)
    s.drawings(0) match
      case Drawing.Picture(DrawingAnchor.OneCell(from, _), _, _, _) =>
        assertEquals(from.cell, ref"B8") // B5 -> shifted down 3
      case other => fail(s"unexpected $other")
    s.drawings(1) match
      case Drawing.Picture(DrawingAnchor.TwoCell(from, to, _), _, _, _) =>
        assertEquals(from.cell, ref"C2") // above insertion, unchanged (row index 1 < 2)
        assertEquals(to.cell, ref"E7") // E4 (one past D3) -> shifted down 3
      case other => fail(s"unexpected $other")
  }

  test("deleteRows clamps a deleted anchor to the deletion point instead of dropping") {
    val s = sheet
      .addImage(png, ref"B5", extent) // row index 4
      .deleteRows(at = 3, count = 4) // delete row indices 3-6
    s.drawings(0) match
      case Drawing.Picture(DrawingAnchor.OneCell(from, _), _, _, _) =>
        assertEquals(from.cell, ref"B4") // clamped to `at` = index 3 = row 4
      case other => fail(s"unexpected $other")
  }

  test("deleteColumns degenerates a fully-deleted two-cell range to zero extent") {
    val s = sheet
      .addImage(png, ref"C2:D3", EditAs.TwoCell) // from C2 (col 2), to E4 (col 4)
      .deleteColumns(at = 2, count = 3) // delete cols C,D,E
    s.drawings(0) match
      case Drawing.Picture(DrawingAnchor.TwoCell(from, to, _), _, _, _) =>
        assertEquals(from.cell.col.index0, 2) // clamped to at
        assertEquals(to.cell.col.index0, 2) // clamped to at -> zero width, accepted
      case other => fail(s"unexpected $other")
  }

  test("absolute anchors and Preserved fragments are untouched by structural edits") {
    val abs = DrawingAnchor.Absolute(Emu(123), Emu(456), extent)
    val s0 = sheet
      .addImage(png, abs)
      .copy(drawings = sheet.addImage(png, abs).drawings :+ Drawing.Preserved("<sp/>"))
    val s = s0.insertRows(0, 10).deleteColumns(0, 2)
    assertEquals(s.drawings(0), Drawing.Picture(abs, png))
    assertEquals(s.drawings(1), Drawing.Preserved("<sp/>"))
  }

  // LAW: a OneCell picture anchor and a comment on the same cell move identically through
  // insert/delete on both axes — except when the anchor index is deleted, where the comment is
  // dropped but the picture clamps to the deletion point (Excel keeps pictures).
  property("shiftAxis moves picture anchors exactly like comments (insert and surviving delete)") {
    val genCase = for
      col <- Gen.choose(0, 15)
      row <- Gen.choose(0, 15)
      at <- Gen.choose(0, 12)
      count <- Gen.choose(1, 4)
      rowAxis <- Gen.oneOf(true, false)
      deleting <- Gen.oneOf(true, false)
    yield (ARef.from0(col, row), at, count, rowAxis, deleting)
    forAll(genCase) { case (cell, at, count, rowAxis, deleting) =>
      val base = sheet
        .addImage(png, cell, extent)
        .comment(cell, Comment.plainText("note", None))
      val shifted = (rowAxis, deleting) match
        case (true, false) => base.insertRows(at, count)
        case (true, true) => base.deleteRows(at, count)
        case (false, false) => base.insertColumns(at, count)
        case (false, true) => base.deleteColumns(at, count)
      val picCell = shifted.pictures.headOption.map(_.anchor) match
        case Some(DrawingAnchor.OneCell(from, _)) => from.cell
        case other => fail(s"picture vanished or changed form: $other")
      val axisIdx = if rowAxis then cell.row.index0 else cell.col.index0
      val deleted = deleting && axisIdx >= at && axisIdx < at + count
      if deleted then
        // comment dropped, picture clamped to `at` on the active axis
        val picAxis = if rowAxis then picCell.row.index0 else picCell.col.index0
        Prop(shifted.comments.isEmpty) :| "comment dropped" &&
        (Prop(picAxis == at) :| s"picture clamped to $at, got $picAxis")
      else
        Prop(shifted.comments.keySet == Set(picCell)) :|
          s"comment moved to ${shifted.comments.keySet}, picture to $picCell"
    }
  }
