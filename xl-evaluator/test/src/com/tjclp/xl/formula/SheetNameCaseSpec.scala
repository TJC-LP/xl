package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.DependentRecalculation.recalculateDependents
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{DefinedName, Workbook}
import munit.FunSuite

/**
 * #752: a reference's sheet qualifier resolves case-insensitively, as in Excel: on a book whose tab
 * is `Data`, `=data!B1`, `=SUM(DATA!A:A)`, `=INDIRECT("data!B1")` and `=data!rate` read `Data`, and
 * recalculation orders and evaluates them like the `Data` spelling.
 */
class SheetNameCaseSpec extends FunSuite:

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private def formula(expr: String): CellValue = CellValue.Formula(expr, None)
  private def a1(s: String): ARef = ARef.parse(s).fold(err => fail(err), identity)

  private def sheetWith(name: String, cells: (String, CellValue)*): Sheet =
    cells.foldLeft(Sheet(SheetName.unsafe(name))) { case (s, (ref, value)) =>
      s.put(a1(ref), value)
    }

  private val data = sheetWith("Data", "A1" -> num(1), "A2" -> num(2), "B1" -> num(7))
  private val summary = Sheet(SheetName.unsafe("Summary"))
  private val book = Workbook(Vector(data, summary))

  private def cached(wb: Workbook, sheet: String, ref: String): Option[CellValue] =
    wb.sheets
      .find(_.name.value == sheet)
      .flatMap(_.cells.get(a1(ref)))
      .map(_.value)
      .collect { case CellValue.Formula(_, Some(v), _) => v }

  test("a qualified cell or range in another case evaluates as the tab's spelling") {
    assertEquals(book.evaluateFormula("=data!B1", "Summary"), Right(num(7)))
    assertEquals(book.evaluateFormula("=SUM(DATA!A:A)", "Summary"), Right(num(3)))
    assertEquals(book.evaluateFormula("=dAtA!A1+Data!A2", "Summary"), Right(num(3)))
  }

  test("INDIRECT text names the sheet in any case") {
    assertEquals(book.evaluateFormula("=INDIRECT(\"data!B1\")", "Summary"), Right(num(7)))
    assertEquals(book.evaluateFormula("=SUM(INDIRECT(\"DATA!A1:A2\"))", "Summary"), Right(num(3)))
  }

  test("a sheet-qualified defined name finds the name scoped to the tab in any case") {
    val scoped = book.copy(metadata =
      book.metadata.copy(definedNames =
        Vector(DefinedName("rate", "Data!$B$1", localSheetId = Some(0)))
      )
    )
    assertEquals(scoped.evaluateFormula("=data!rate", "Summary"), Right(num(7)))
    assertEquals(scoped.evaluateFormula("=Data!rate", "Summary"), Right(num(7)))
  }

  test("recalculation orders =data!B1 after Data!B1 and caches its value") {
    // Data!B1 is itself a formula, so a wrong order or a failed lookup leaves Summary!A1 wrong
    val wb = Workbook(
      Vector(
        sheetWith("Summary", "A1" -> formula("data!B1*2")),
        sheetWith("Data", "A1" -> num(5), "B1" -> formula("A1+1"))
      )
    )
    val result = wb.recalculate()
    assertEquals(result.errors, Vector.empty)
    assertEquals(cached(result.workbook, "Data", "B1"), Some(num(6)))
    assertEquals(cached(result.workbook, "Summary", "A1"), Some(num(12)))
  }

  test(
    "tabs differing only by case: an exact qualifier reads its own tab, a third spelling fails"
  ) {
    val twins = Workbook(
      Vector(sheetWith("Data", "A1" -> num(1)), sheetWith("DATA", "A1" -> num(2)), summary)
    )
    assertEquals(twins.evaluateFormula("=Data!A1", "Summary"), Right(num(1)))
    assertEquals(twins.evaluateFormula("=DATA!A1", "Summary"), Right(num(2)))
    assert(twins.evaluateFormula("=data!A1", "Summary").isLeft)
  }

  test(
    "the dependency graph resolves a case-variant qualifier on a defined name as evaluation does"
  ) {
    // rate is scoped to Data (5) and to Other; Other!B1 reads Data's through `data!rate`
    val wb = Workbook(
      Vector(
        sheetWith("Data", "A1" -> num(5)),
        sheetWith(
          "Other",
          "A1" -> num(9),
          "B1" -> formula("data!rate"),
          "B2" -> formula("Data!rate"),
          "B3" -> formula("data!dyn")
        )
      )
    ).copy(metadata =
      Workbook().metadata.copy(definedNames =
        Vector(
          DefinedName("rate", "Data!$A$1", localSheetId = Some(0)),
          DefinedName("rate", "Other!$A$1", localSheetId = Some(1)),
          DefinedName("dyn", "INDIRECT(\"Data!A1\")", localSheetId = Some(0))
        )
      )
    )
    val other = SheetName.unsafe("Other")
    def at(ref: String) = QualifiedRef(other, a1(ref))
    assertEquals(wb.evaluateFormula("=data!rate", "Other"), Right(num(5)))
    // not an "unresolved name" reader: `xl audit` reported Other!B1 before
    assertEquals(DependencyGraph.unresolvedReaders(wb), Set.empty[QualifiedRef])
    // a name reaching INDIRECT through a case-variant qualifier is a dynamic reader
    assert(DependencyGraph.dynamicCells(wb).contains(at("B3")), "data!dyn reads INDIRECT")
    assert(!DependencyGraph.dynamicCells(wb).contains(at("B1")))
  }

  test("recalculateDependents seeds a case-variant sheet name onto the tab's cells") {
    val wb = Workbook(
      Vector(
        sheetWith("Data", "A1" -> num(1)),
        sheetWith("Summary", "A1" -> CellValue.Formula("Data!A1*2", Some(num(2))))
      )
    )
    val lower = SheetName.unsafe("data")
    val edited =
      wb(lower).map(s => wb.put(s.put(a1("A1"), num(10)))).fold(e => fail(e.message), identity)
    val recalculated = edited.recalculateDependents(lower, Set(a1("A1")))
    assertEquals(cached(recalculated, "Summary", "A1"), Some(num(20)))
  }
