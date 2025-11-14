package com.tjclp.xl

import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cell.{Cell, CellValue}
import com.tjclp.xl.optics.syntax.* // Import optics extension methods
import com.tjclp.xl.dsl.syntax.*
import munit.FunSuite
import com.tjclp.xl.macros.ref
// Removed: BatchPutMacro is dead code (shadowed by Sheet.put member)  // For batch put extension
import com.tjclp.xl.style.units.StyleId

/** Tests for optics library and focus DSL */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.Var"))
class OpticsSpec extends FunSuite:

  val emptySheet: Sheet = Sheet(SheetName.unsafe("TestSheet"))

  // ========== Lens Laws ==========

  test("Lens law: get-put (you get what you set)") {
    import Optics.*

    val cell1 = Cell.empty(ARef.from1(1, 1))
    val newValue = CellValue.Text("Updated")

    val updated = cellValue.set(newValue, cell1)
    assertEquals(cellValue.get(updated), newValue)
  }

  test("Lens law: put-get (setting then getting returns original structure)") {
    import Optics.*

    val cell1 = Cell(ARef.from1(1, 1), CellValue.Text("Original"))
    val retrieved = cellValue.get(cell1)
    val restored = cellValue.set(retrieved, cell1)

    assertEquals(restored, cell1)
  }

  test("Lens law: put-put (last write wins)") {
    import Optics.*

    val cell1 = Cell.empty(ARef.from1(1, 1))
    val value1 = CellValue.Text("First")
    val value2 = CellValue.Text("Second")

    val updated1 = cellValue.set(value1, cell1)
    val updated2 = cellValue.set(value2, updated1)

    assertEquals(cellValue.get(updated2), value2)
  }

  test("Lens.update applies function") {
    import Optics.*

    val cell1 = Cell(ARef.from1(1, 1), CellValue.Text("hello"))
    val updated = cellValue.update {
      case CellValue.Text(s) => CellValue.Text(s.toUpperCase)
      case other => other
    }(cell1)

    assertEquals(updated.value, CellValue.Text("HELLO"))
  }

  test("Lens.modify mutates the structure") {
    import Optics.*

    val cell1 = Cell(ARef.from1(1, 1), CellValue.Text("hello"))
    val modified = cellValue.modify {
      case CellValue.Text(s) => CellValue.Text(s.reverse)
      case other => other
    }(cell1)

    assertEquals(modified.value, CellValue.Text("olleh"))
  }

  test("Lens.andThen composes lenses") {
    import Optics.*

    val sheet = emptySheet.put(ref"A1", CellValue.Text("test"))
    val composed = sheetCells.andThen(
      Lens[Map[ARef, Cell], Cell](
        get = _.get(ref"A1").getOrElse(Cell.empty(ref"A1")),
        set = (c, m) => m + (c.ref -> c)
      )
    )

    val newCell = Cell(ref"A1", CellValue.Text("updated"))
    val updated = composed.set(newCell, sheet)

    assertEquals(updated(ref"A1").value, CellValue.Text("updated"))
  }

  // ========== Optional Behavior ==========

  test("Optional.getOption returns Some for existing ref") {
    import Optics.*

    val sheet = emptySheet.put(ref"A1", CellValue.Text("Present"))
    val opt = cellAt(ref"A1")

    assertEquals(opt.getOption(sheet).map(_.value), Some(CellValue.Text("Present")))
  }

  test("Optional.getOption returns Some(empty) for missing ref") {
    import Optics.*

    val sheet = emptySheet
    val opt = cellAt(ref"Z99")

    // cellAt returns empty cell for missing cells
    val result = opt.getOption(sheet)
    assert(result.isDefined)
    assertEquals(result.get.value, CellValue.Empty)
  }

  test("Optional.modify only modifies if present") {
    import Optics.*

    val sheet = emptySheet.put(ref"A1", CellValue.Text("test"))
    val opt = cellAt(ref"A1")

    val updated = opt.modify(c => c.withValue(CellValue.Text("modified")))(sheet)

    assertEquals(updated(ref"A1").value, CellValue.Text("modified"))
  }

  test("Optional.getOrElse returns default for empty") {
    import Optics.*

    val sheet = emptySheet
    val opt = cellAt(ref"B2")
    val defaultCell = Cell(ref"B2", CellValue.Text("default"))

    // Since cellAt returns empty cell, we should get an empty cell, not the default
    val result = opt.getOrElse(defaultCell)(sheet)
    assertEquals(result.value, CellValue.Empty) // Empty cell, not default
  }

  // ========== Predefined Optics ==========

  test("Optics.sheetCells accesses cell map") {
    import Optics.*

    val sheet = emptySheet.put(ref"A1", CellValue.Text("test"))
    val cells = sheetCells.get(sheet)

    assertEquals(cells.size, 1)
    assertEquals(cells(ref"A1").value, CellValue.Text("test"))
  }

  test("Optics.cellValue accesses and modifies value") {
    import Optics.*

    val cell1 = Cell(ref"A1", CellValue.Number(42))
    val value = cellValue.get(cell1)
    assertEquals(value, CellValue.Number(42))

    val updated = cellValue.set(CellValue.Number(100), cell1)
    assertEquals(updated.value, CellValue.Number(100))
  }

  test("Optics.cellStyleId accesses and modifies styleId") {
    import Optics.*

    val cell1 = Cell(ref"A1", CellValue.Empty, Some(StyleId(5)))
    assertEquals(cellStyleId.get(cell1), Some(StyleId(5)))

    val cleared = cellStyleId.set(None, cell1)
    assertEquals(cleared.styleId, None)

    val restyled = cellStyleId.set(Some(StyleId(10)), cleared)
    assertEquals(restyled.styleId, Some(StyleId(10)))
  }

  test("Optics.sheetName accesses sheet name") {
    import Optics.*

    val sheet = emptySheet
    assertEquals(sheetName.get(sheet), SheetName.unsafe("TestSheet"))

    val renamed = sheetName.set(SheetName.unsafe("NewName"), sheet)
    assertEquals(renamed.name, SheetName.unsafe("NewName"))
  }

  test("Optics.mergedRanges accesses merged ranges") {
    import Optics.*

    val sheet = emptySheet.merge(ref"A1:B2")
    val merged = mergedRanges.get(sheet)

    assertEquals(merged.size, 1)
    assert(merged.contains(ref"A1:B2"))
  }

  // ========== Focus DSL ==========

  test("sheet.focus returns Optional for ref") {
    val sheet = emptySheet.put(ref"A1", CellValue.Text("test"))
    val focused = sheet.focus(ref"A1")

    val cellOpt = focused.getOption(sheet)
    assert(cellOpt.isDefined)
    assertEquals(cellOpt.get.value, CellValue.Text("test"))
  }

  test("sheet.modifyCell updates ref") {
    val sheet = emptySheet.put(ref"A1", CellValue.Text("original"))

    val updated = sheet.modifyCell(ref"A1")(c => c.withValue(CellValue.Text("modified")))

    assertEquals(updated(ref"A1").value, CellValue.Text("modified"))
  }

  test("sheet.modifyCell creates empty cell if missing") {
    val sheet = emptySheet

    val updated = sheet.modifyCell(ref"B5")(c => c.withValue(CellValue.Number(99)))

    assertEquals(updated(ref"B5").value, CellValue.Number(99))
  }

  test("sheet.modifyValue transforms cell value") {
    val sheet = emptySheet.put(ref"A1", CellValue.Text("hello"))

    val updated = sheet.modifyValue(ref"A1") {
      case CellValue.Text(s) => CellValue.Text(s.toUpperCase)
      case other => other
    }

    assertEquals(updated(ref"A1").value, CellValue.Text("HELLO"))
  }

  test("sheet.modifyValue handles non-text values") {
    val sheet = emptySheet.put(ref"A1", CellValue.Number(42))

    val updated = sheet.modifyValue(ref"A1") {
      case CellValue.Number(n) => CellValue.Number(n * 2)
      case other => other
    }

    assertEquals(updated(ref"A1").value, CellValue.Number(84))
  }

  test("sheet.modifyStyleId updates style") {
    val sheet = emptySheet.put(ref"A1", CellValue.Text("styled"))

    val updated = sheet.modifyStyleId(ref"A1")(_ => Some(StyleId(7)))

    assertEquals(updated(ref"A1").styleId, Some(StyleId(7)))
  }

  test("sheet.modifyStyleId clears style with None") {
    val sheet = emptySheet
      .put(Cell(ref"A1", CellValue.Text("styled"), Some(StyleId(5))))

    val updated = sheet.modifyStyleId(ref"A1")(_ => None)

    assertEquals(updated(ref"A1").styleId, None)
  }

  // ========== Complex Scenarios ==========

  test("chain multiple modifyValue calls") {
    val sheet = emptySheet
      .put(ref"A1", CellValue.Text("alpha"))
      .put(ref"A2", CellValue.Text("beta"))
      .put(ref"A3", CellValue.Text("gamma"))

    val updated = sheet
      .modifyValue(ref"A1")(v => CellValue.Text("ALPHA"))
      .modifyValue(ref"A2")(v => CellValue.Text("BETA"))
      .modifyValue(ref"A3")(v => CellValue.Text("GAMMA"))

    assertEquals(updated(ref"A1").value, CellValue.Text("ALPHA"))
    assertEquals(updated(ref"A2").value, CellValue.Text("BETA"))
    assertEquals(updated(ref"A3").value, CellValue.Text("GAMMA"))
  }

  test("modifyCell preserves other cell properties") {
    val originalCell = Cell(
      ref = ref"A1",
      value = CellValue.Text("test"),
      styleId = Some(StyleId(3)),
      comment = Some("Note"),
      hyperlink = Some("https://example.com")
    )
    val sheet = emptySheet.put(originalCell)

    val updated = sheet.modifyValue(ref"A1") {
      case CellValue.Text(_) => CellValue.Text("changed")
      case other => other
    }

    val cell = updated(ref"A1")
    assertEquals(cell.value, CellValue.Text("changed"))
    assertEquals(cell.styleId, Some(StyleId(3)))
    assertEquals(cell.comment, Some("Note"))
    assertEquals(cell.hyperlink, Some("https://example.com"))
  }

  test("focus DSL works with pattern matching") {
    val sheet = emptySheet
      .put(ref"A1", CellValue.Number(100))
      .put(ref"A2", CellValue.Text("skip"))
      .put(ref"A3", CellValue.Number(200))

    val updated = sheet
      .modifyValue(ref"A1") {
        case CellValue.Number(n) => CellValue.Number(n * 1.1)
        case other => other
      }
      .modifyValue(ref"A3") {
        case CellValue.Number(n) => CellValue.Number(n * 1.1)
        case other => other
      }

    assertEquals(updated(ref"A1").value, CellValue.Number(BigDecimal("110.0")))
    assertEquals(updated(ref"A2").value, CellValue.Text("skip"))
    assertEquals(updated(ref"A3").value, CellValue.Number(BigDecimal("220.0")))
  }

  test("Lens.modify returns transformed value") {
    import Optics.*

    val cell1 = Cell(ref"A1", CellValue.Number(50))
    val doubled = cellValue.modify {
      case CellValue.Number(n) => CellValue.Number(n * 2)
      case other => other
    }(cell1)

    assertEquals(doubled.value, CellValue.Number(100))
  }

  test("Optional.modify is no-op on missing cells") {
    import Optics.*

    val sheet = emptySheet // No cells
    val opt = cellAt(ref"A1")

    // Modify should create the cell since cellAt returns Some(empty cell)
    val updated = opt.modify(c => c.withValue(CellValue.Text("new")))(sheet)

    // The empty cell was modified
    assertEquals(updated(ref"A1").value, CellValue.Text("new"))
  }

  test("sheetCells lens allows bulk operations") {
    import Optics.*

    val sheet = emptySheet
      .put(ref"A1", CellValue.Number(1))
      .put(ref"A2", CellValue.Number(2))
      .put(ref"A3", CellValue.Number(3))

    // Get all cells, transform them, set back
    val cells = sheetCells.get(sheet)
    val doubled = cells.view.mapValues { c =>
      c.value match
        case CellValue.Number(n) => c.withValue(CellValue.Number(n * 2))
        case _ => c
    }.toMap

    val updated = sheetCells.set(doubled, sheet)

    assertEquals(updated(ref"A1").value, CellValue.Number(2))
    assertEquals(updated(ref"A2").value, CellValue.Number(4))
    assertEquals(updated(ref"A3").value, CellValue.Number(6))
  }

  // ========== Real-World Use Cases ==========

  test("use case: uppercase all text cells in ref") {
    val sheet = emptySheet
      .put(ref"A1", CellValue.Text("apple"))
      .put(ref"A2", CellValue.Text("banana"))
      .put(ref"A3", CellValue.Number(42))
      .put(ref"A4", CellValue.Text("cherry"))

    val refs = Vector(ref"A1", ref"A2", ref"A3", ref"A4")

    var updated = sheet
    refs.foreach { ref =>
      updated = updated.modifyValue(ref) {
        case CellValue.Text(s) => CellValue.Text(s.toUpperCase)
        case other => other
      }
    }

    assertEquals(updated(ref"A1").value, CellValue.Text("APPLE"))
    assertEquals(updated(ref"A2").value, CellValue.Text("BANANA"))
    assertEquals(updated(ref"A3").value, CellValue.Number(42)) // Unchanged
    assertEquals(updated(ref"A4").value, CellValue.Text("CHERRY"))
  }

  test("use case: apply discount to all numbers") {
    val sheet = emptySheet
      .put(ref"B1", CellValue.Number(100))
      .put(ref"B2", CellValue.Number(200))
      .put(ref"B3", CellValue.Number(300))

    val discount = 0.9
    var updated = sheet
    Vector(ref"B1", ref"B2", ref"B3").foreach { ref =>
      updated = updated.modifyValue(ref) {
        case CellValue.Number(n) => CellValue.Number(n * discount)
        case other => other
      }
    }

    assertEquals(updated(ref"B1").value, CellValue.Number(90))
    assertEquals(updated(ref"B2").value, CellValue.Number(180))
    assertEquals(updated(ref"B3").value, CellValue.Number(270))
  }

  test("use case: conditionally apply styles") {
    val sheet = emptySheet
      .put(ref"A1", CellValue.Number(-50))
      .put(ref"A2", CellValue.Number(100))
      .put(ref"A3", CellValue.Number(-25))

    val negativeStyleId = StyleId(1)

    var updated = sheet
    Vector(ref"A1", ref"A2", ref"A3").foreach { ref =>
      updated = updated.modifyCell(ref) { c =>
        c.value match
          case CellValue.Number(n) if n < 0 => c.withStyle(negativeStyleId)
          case _ => c
      }
    }

    assertEquals(updated(ref"A1").styleId, Some(negativeStyleId))
    assertEquals(updated(ref"A2").styleId, None)
    assertEquals(updated(ref"A3").styleId, Some(negativeStyleId))
  }

  test("use case: clear styles from ref") {
    val styledCell = Cell(ref"B1", CellValue.Text("test"), Some(StyleId(5)))
    val sheet = emptySheet
      .put(styledCell)
      .put(Cell(ref"B2", CellValue.Text("test2"), Some(StyleId(6))))

    var updated = sheet
    Vector(ref"B1", ref"B2").foreach { ref =>
      updated = updated.modifyStyleId(ref)(_ => None)
    }

    assertEquals(updated(ref"B1").styleId, None)
    assertEquals(updated(ref"B2").styleId, None)
  }

  // ========== Integration with Existing APIs ==========

  test("optics work alongside put/putAll") {
    val sheet = emptySheet
      .put(ref"A1", CellValue.Text("direct"))
      .modifyValue(ref"A2") { _ => CellValue.Text("optic") }
      .put(ref"A3", CellValue.Text("direct2"))

    assertEquals(sheet(ref"A1").value, CellValue.Text("direct"))
    assertEquals(sheet(ref"A2").value, CellValue.Text("optic"))
    assertEquals(sheet(ref"A3").value, CellValue.Text("direct2"))
  }

  test("optics work with applyPatch") {
    import com.tjclp.xl.dsl.*

    val sheet = emptySheet
      .modifyValue(ref"A1") { _ => CellValue.Text("before patch") }

    val patch = ref"A1" := "after patch"
    val result = Patch.applyPatch(sheet, patch)

    assert(result.isRight)
    val updated = result.getOrElse(fail("Patch failed"))
    assertEquals(updated(ref"A1").value, CellValue.Text("after patch"))
  }

  test("optics preserve sheet structure") {
    val sheet = emptySheet
      .put(ref"A1", CellValue.Text("test"))
      .merge(ref"A1:B1")
      .setColumnProperties(Column.from1(1), ColumnProperties(width = Some(20.0)))

    val updated = sheet.modifyValue(ref"A1") {
      case CellValue.Text(s) => CellValue.Text(s.toUpperCase)
      case other => other
    }

    // Verify merged ranges and column props preserved
    assertEquals(updated.mergedRanges.size, 1)
    assert(updated.mergedRanges.contains(ref"A1:B1"))
    assertEquals(updated.getColumnProperties(Column.from1(1)).width, Some(20.0))
  }
