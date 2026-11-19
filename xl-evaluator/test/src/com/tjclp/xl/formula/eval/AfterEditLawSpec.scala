package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.formula.graph.CompactGraphGens
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.{CellStore, Sheet}
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/**
 * #891: the after-edit pass is a full recalculation restricted to what the edit can change. On a
 * book a full recalculation has just settled, edit one constant (or clear it, or write a formula
 * into it) and recalculate after the edit: the workbook is exactly what a full recalculation of the
 * edited book gives, every cell, outside the cone as well as in it. The cone is the recalculation
 * graph's (#504's precise lookups); the generators cover ranges, whole columns and rows, running
 * totals, defined names, VLOOKUP strips, INDIRECT/OFFSET, cycles and unknown sheets.
 */
class AfterEditLawSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(200)

  private val clock = Clock.fixedDate(java.time.LocalDate.of(2026, 1, 2))

  private def options = RecalcOptions(
    clock = clock,
    rng = Rng.seeded(5L),
    iterative = IterativeMode.Off
  )

  /** The edit: a sheet, a cell, and what it becomes. */
  private final case class Edit(sheet: SheetName, ref: ARef, value: Option[CellValue])

  private val genEdit: Gen[Edit] =
    for
      sheet <- Gen.oneOf(CompactGraphGens.sheetNames).map(SheetName.unsafe)
      col <- Gen.choose(0, CompactGraphGens.Cols - 1)
      row <- Gen.choose(0, CompactGraphGens.Rows - 1)
      value <- Gen.frequency(
        6 -> Gen.choose(-5, 50).map(n => Some(CellValue.Number(BigDecimal(n)))),
        2 -> Gen.const(None),
        1 -> Gen.const(Some(CellValue.Text("x"))),
        1 -> Gen.const(Some(CellValue.Formula("A1+1")))
      )
    yield Edit(sheet, ARef.from0(col, row), value)

  private def applyEdit(wb: Workbook, edit: Edit): Workbook =
    wb.sheets.find(_.name == edit.sheet).fold(wb) { sheet =>
      val next: Sheet = edit.value match
        case Some(v) => sheet.put(Cell(edit.ref, v))
        case None => sheet.remove(edit.ref)
      wb.put(next)
    }

  private def cellsOf(wb: Workbook): Map[SheetName, CellStore] =
    wb.sheets.iterator.map(s => s.name -> s.cells).toMap

  /** Formula cells holding no cache (a data-table record's is its own). */
  private def uncachedFormulas(wb: Workbook): Set[QualifiedRef] =
    wb.sheets.iterator.flatMap { s =>
      s.cells.iterator.collect {
        case Cell(ref, CellValue.Formula(_, None, kind), _, _)
            if !kind.isInstanceOf[com.tjclp.xl.cells.FormulaKind.DataTable] =>
          QualifiedRef(s.name, ref)
      }
    }.toSet

  property("after an edit: the cone equals a full recalculation, the rest is untouched") {
    // the cone includes every formula holding no value (Excel has none), so nothing is exempt
    Prop.forAll(CompactGraphGens.genWorkbook, genEdit) { (wb, edit) =>
      val settled = wb.recalculate(options).workbook
      val edited = applyEdit(settled, edit)
      val after = cellsOf(
        edited.recalculateAfterEdit(edit.sheet, Set(edit.ref), options).workbook
      )
      val full = cellsOf(edited.recalculate(options).workbook)
      val before = cellsOf(edited)
      val cone = DependentRecalculation.editCone(edited, Set(QualifiedRef(edit.sheet, edit.ref)))
      def at(cells: Map[SheetName, CellStore], q: QualifiedRef) =
        cells.get(q.sheet).flatMap(_.get(q.ref))
      val all = (before.toVector ++ full.toVector)
        .flatMap((name, cells) => cells.refs.map(QualifiedRef(name, _)))
        .distinct
      val outside = all.filterNot(cone.contains).filter(q => at(after, q) != at(before, q))
      val inside = all
        .filter(cone.contains)
        .filter(q => at(after, q) != at(full, q))
      def show(qs: Vector[QualifiedRef], other: Map[SheetName, CellStore], label: String) =
        qs.map(q => s"$q: $label=${at(other, q).map(_.value)} after=${at(after, q).map(_.value)}")
      Prop.all(
        Prop(outside.isEmpty) :| s"edit $edit, outside the cone\n${show(outside, before, "before")
            .mkString("\n")}",
        Prop(inside.isEmpty) :| s"edit $edit, in the cone\n${show(inside, full, "full")
            .mkString("\n")}"
      )
    }
  }

  test("Excel parity: a lookup over an always-failing formula is withdrawn after an edit too") {
    // B1 fails on every pass (an unknown sheet), so a settled book holds it uncached; E1 looks up
    // through the table B1 sits in. Excel holds an error value in B1 and the lookup sees it; xl's
    // after-edit pass computes every uncached formula with the edit (#891), so B1 fails again and
    // E1 is withdrawn, exactly as a full recalculation has it.
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(1)))
      .put(ARef.from0(0, 1), CellValue.Number(BigDecimal(2)))
      .put(ARef.from0(1, 0), CellValue.Formula("'Gone'!A1"))
      .put(ARef.from0(1, 1), CellValue.Number(BigDecimal(7)))
      .put(ARef.from0(4, 0), CellValue.Formula("VLOOKUP(2,A1:B2,2,FALSE)"))
    val settled = Workbook(sheet).recalculate(options).workbook
    val edited = settled(sheet.name).fold(
      e => fail(e.toString),
      s => settled.put(s.put(ARef.from0(0, 1), CellValue.Number(BigDecimal(2))))
    )
    val after = edited.recalculateAfterEdit(sheet.name, Set(ARef.from0(0, 1)), options)
    val full = edited.recalculate(options)
    def e1(r: RecalcResult) =
      r.workbook(sheet.name).toOption.flatMap(_.cells.get(ARef.from0(4, 0))).map(_.value)
    assertEquals(e1(full), Some(CellValue.Formula("VLOOKUP(2,A1:B2,2,FALSE)", None)))
    assertEquals(e1(after), e1(full))
  }

  private def n(i: Int): CellValue = CellValue.Number(BigDecimal(i))

  test("#504: an edit inside a fixed-column VLOOKUP's table, off its strips, leaves it clean") {
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(1)))
      .put(ARef.from0(1, 0), CellValue.Number(BigDecimal(10)))
      .put(ARef.from0(2, 0), CellValue.Number(BigDecimal(100)))
      .put(ARef.from0(4, 0), CellValue.Formula("VLOOKUP(1,A1:C1,3,FALSE)"))
    val settled = Workbook(sheet).recalculate(options).workbook
    // B1 lies in the table but neither in the key column A nor the result column C
    val edited = settled(sheet.name).fold(
      e => fail(e.toString),
      s => settled.put(s.put(ARef.from0(1, 0), CellValue.Number(BigDecimal(20))))
    )
    val result = edited.recalculateAfterEdit(sheet.name, Set(ARef.from0(1, 0)), options)
    assertEquals(result.evaluated.getOrElse(sheet.name, Map.empty), Map.empty)
    // and the cell keeps its settled value
    assertEquals(
      result.workbook(sheet.name).toOption.flatMap(_.cells.get(ARef.from0(4, 0))).map(_.value),
      Some(CellValue.Formula("VLOOKUP(1,A1:C1,3,FALSE)", Some(CellValue.Number(BigDecimal(100)))))
    )
  }

  test("a seeded RAND reader draws the same value after an edit as in a full recalculation") {
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(1)))
      .put(ARef.from0(1, 0), CellValue.Formula("A1+RAND()"))
      .put(ARef.from0(1, 1), CellValue.Formula("A1*2+RAND()"))
    val edited = Workbook(sheet).put(sheet.put(ARef.from0(0, 0), CellValue.Number(BigDecimal(2))))
    val after =
      edited.recalculateAfterEdit(sheet.name, Set(ARef.from0(0, 0)), options).evaluated
    val full = edited.recalculate(options).evaluated
    assertEquals(after, full)
  }
