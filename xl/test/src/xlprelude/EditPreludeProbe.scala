// Deliberately OUTSIDE com.tjclp.xl (see ScriptingPreludeTest): the edit algebra must resolve
// through `import com.tjclp.xl.scripting.{*, given}` alone, on receivers exactly as scripts see them.
package xlprelude

import java.time.LocalDate

import munit.FunSuite

import com.tjclp.xl.scripting.{*, given}

/**
 * Gate test for W2.1 (ADR-017 §2.12) through the prelude: `Edit` and its targets, the nested
 * `Edit.FillDir` / `Edit.SortKeySpec` through the export forwarder, `wb.edit` / `wb.editIn` /
 * `sheet.edit` with the prelude's `given FormulaSupport`, `Edit.lower` / `Patch.toEdits`,
 * `EditSchema` and `EditScope` (the exported name of `ops.Scope`). xl-cli's `FillDirection` /
 * `SortDirection` / `SortMode` are NOT part of this surface (naming them here would not compile).
 * Compile success is most of the test.
 */
class EditPreludeProbe extends FunSuite:

  private val Data = SheetName.unsafe("Data")
  private val Other = SheetName.unsafe("Other")

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  // fx"=A1*2" stores its text verbatim; compare the reference shape without the leading '='
  private def formulaText(v: CellValue): String = v match
    case CellValue.Formula(expr, _, _) => expr.stripPrefix("=")
    case other => other.toString

  private val book: Workbook = Workbook(
    Sheet(Data).put(ref"A1", 10).put(ref"A2", 20).put(ref"B1", fx"=A1*2"),
    Sheet(Other).put(ref"A1", fx"=Data!A1+1")
  )

  test("Edit cases, targets, hints, overlays and nested enums resolve; wb.editIn applies them"):
    val edits = Vector(
      Edit.put(Loc(None, ref"C1"), LocalDate.of(2025, 1, 15)),
      Edit.PutValues(
        Area(None, ref"D1:D2"),
        Vector(num(1), num(2)),
        Some(FormatHint.Explicit(NumFmt.Percent))
      ),
      Edit.Style(Area(None, ref"A1:A2"), StyleOverlay(bold = Some(true)), StyleMode.Merge),
      Edit.Fill(Area(None, CellRange(ref"B1", ref"B1")), ref"B1:B2", Edit.FillDir.Down),
      Edit.Sort(
        Area(None, ref"A1:A2"),
        Vector(Edit.SortKeySpec.descending(Column.from0(0))),
        hasHeader = false
      ),
      Edit.Clear(Area.cell(Loc(None, ref"D2")), ClearWhat.contents),
      Edit.GroupRows(None, RowSpan(Row.from1(1), Row.from1(2)), 1, collapsed = false),
      Edit.ColWidth(None, ColSpan.single(Column.from0(0)), 20.0)
    )
    val sheet = book.editIn(Data)(edits*).flatMap(_(Data)).unsafe
    assertEquals(sheet(ref"A1").value, num(20))
    assertEquals(formulaText(sheet(ref"B2").value), "A2*2")
    assertEquals(
      sheet(ref"C1").styleId.flatMap(sheet.styleRegistry.get).map(_.numFmt),
      Some(NumFmt.Date)
    )
    assertEquals(
      sheet(ref"D1").styleId.flatMap(sheet.styleRegistry.get).map(_.numFmt),
      Some(NumFmt.Percent)
    )
    assert(!sheet.contains(ref"D2"))
    assertEquals(sheet.getRowProperties(Row.from1(1)).outlineLevel, Some(1))
    assertEquals(sheet.getColumnProperties(Column.from0(0)).width, Some(20.0))

  test("wb.edit with the prelude's given: a drag, an insert and a rename go through the evaluator"):
    val renamedName = SheetName.unsafe("Renamed")
    val r = book.edit(
      Edit.DragFormula(Area(Some(Data), ref"E1:E2"), "=A1*3", ref"E1", None),
      Edit.InsertRows(Some(Data), Row.from0(0), 1),
      Edit.RenameSheet(Data, renamedName)
    )
    val renamed = r.flatMap(_(renamedName)).unsafe
    assertEquals(formulaText(renamed(ref"E3").value), "A3*3")
    assertEquals(formulaText(r.flatMap(_(Other)).unsafe(ref"A1").value), "Renamed!A2+1")
    // an explicit text-only support refuses the same drag with the capability it lacks
    val drag = Edit.DragFormula(Area(Some(Data), ref"E1:E2"), "=A1*3", ref"E1", None)
    assert(book.edit(drag)(using FormulaSupport.textOnly).isLeft)

  test("sheet.edit, Edit.lower, Patch.toEdits, Edit.validate and EditSchema resolve"):
    val sheet = Sheet(Data).put(ref"A1", 1)
    val edited = sheet
      .edit(Edit.Put(Loc(None, ref"A2"), num(2), None), Edit.Merge(Area(None, ref"A1:B1")))
      .unsafe
    assertEquals(edited(ref"A2").value, num(2))
    assertEquals(edited.mergedRanges.size, 1)
    val lowered: Option[Patch] = Edit.lower(Edit.Merge(Area(None, ref"A1:B1")), sheet)
    assertEquals(lowered, Some(Patch.Merge(ref"A1:B1")))
    assertEquals(
      Patch.toEdits(Patch.Merge(ref"A1:B1"), sheet),
      Some(Vector(Edit.Merge(Area(None, ref"A1:B1"))))
    )
    assertEquals(EditSchema.find("putf").map(_.name), Some("put-formula"))
    assertEquals(EditSchema.all.size, EditSchema.all.map(_.name).distinct.size)
    assert(Edit.validate(Edit.MoveSheet(Data, None, None, None)).isLeft)
    // `ops.Scope` reaches scripts as `EditScope` (a bare `Scope` would collide with JMH's/ZIO's)
    assertEquals(EditScope.none.after(Edit.RenameSheet(Data, Other)), EditScope.none)
    assertEquals(EditScope.of(Data).after(Edit.RenameSheet(Data, Other)), EditScope.of(Other))

  test("the nested Edit enums resolve; xl-cli's FillDirection/SortDirection/SortMode do not"):
    import scala.compiletime.testing.typeChecks
    assert(typeChecks("Edit.FillDir.Down"))
    assert(typeChecks("Edit.SortDir.Descending"))
    assert(typeChecks("Edit.SortMode.Numeric"))
    assert(!typeChecks("FillDirection.Down"))
    assert(!typeChecks("SortDirection.Descending"))
    assert(!typeChecks("SortMode.Numeric"))
