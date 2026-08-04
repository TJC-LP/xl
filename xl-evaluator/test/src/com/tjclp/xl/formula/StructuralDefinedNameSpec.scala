package com.tjclp.xl.formula

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.formula.eval.StructuralEditor
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{DefinedName, Workbook, WorkbookMetadata}
import munit.FunSuite

/**
 * GH-473: structural edits rewrite general defined-name refersTo text (workbook- AND sheet-scoped)
 * through the same shift plane as print areas / DV / tables (GH-429) — a reference targeting the
 * edited sheet tracks the edit, a fully-deleted target degrades to "#REF!", and constants /
 * other-sheet references ride byte-identical. A refersTo may be a top-level comma union
 * (multi-range) the formula parser rejects; each segment shifts independently.
 */
class StructuralDefinedNameSpec extends FunSuite:

  private val One = SheetName.unsafe("One")
  private val Two = SheetName.unsafe("Two")

  private def wbWith(names: DefinedName*): Workbook =
    Workbook(
      Vector(new Sheet(name = One), new Sheet(name = Two)),
      metadata = WorkbookMetadata(definedNames = names.toVector)
    )

  private def formulas(wb: Workbook): Vector[String] =
    wb.metadata.definedNames.map(_.formula)

  test("a sheet-scoped name tracks an insert on its target sheet (the GH-473 repro)") {
    val wb = wbWith(DefinedName("Case", "Two!$B$2", localSheetId = Some(1)))
    val r = StructuralEditor.insertRows(wb, Two, at = 0, count = 5)
    assertEquals(formulas(r), Vector("Two!$B$7"))
    // scope is untouched — only the refersTo moved
    assertEquals(r.metadata.definedNames.map(_.localSheetId), Vector(Some(1)))
  }

  test("a workbook-scoped name tracks the edit too") {
    val wb = wbWith(DefinedName("Total", "Two!$A$1:$B$2"))
    val r = StructuralEditor.insertRows(wb, Two, at = 0, count = 5)
    assertEquals(formulas(r), Vector("Two!$A$6:$B$7"))
  }

  test("a name targeting an UNTOUCHED sheet rides byte-identical") {
    val wb = wbWith(
      DefinedName("Other", "One!$B$2"),
      DefinedName("Rate", "0.08"),
      DefinedName("Label", "\"two, please\"")
    )
    val r = StructuralEditor.insertRows(wb, Two, at = 0, count = 5)
    assertEquals(formulas(r), Vector("One!$B$2", "0.08", "\"two, please\""))
  }

  test("deleting the rows a name points at degrades its refersTo to #REF!") {
    val wb = wbWith(DefinedName("Case", "Two!$B$2", localSheetId = Some(1)))
    val r = StructuralEditor.deleteRows(wb, Two, at = 1, count = 1)
    assertEquals(formulas(r), Vector("#REF!"))
  }

  test("a multi-range refersTo shifts each top-level comma segment independently") {
    val wb = wbWith(DefinedName("Bands", "Two!$A$1:$B$2,Two!$D$10"))
    val r = StructuralEditor.insertRows(wb, Two, at = 0, count = 5)
    assertEquals(formulas(r), Vector("Two!$A$6:$B$7,Two!$D$15"))
  }

  test("column edits shift names on the edited axis") {
    val wb = wbWith(DefinedName("Case", "Two!$B$2", localSheetId = Some(1)))
    val r = StructuralEditor.insertColumns(wb, Two, at = 0, count = 2)
    assertEquals(formulas(r), Vector("Two!$D$2"))
  }
