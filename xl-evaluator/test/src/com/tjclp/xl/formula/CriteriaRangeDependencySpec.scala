package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.DependentRecalculation
import com.tjclp.xl.formula.eval.DependentRecalculation.recalculateDependents
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/** Resizing named and LET ranges must invalidate and order every cell actually read. */
class CriteriaRangeDependencySpec extends FunSuite:
  private def n(value: Int): CellValue = CellValue.Number(BigDecimal(value))
  private val sheet = Sheet(SheetName.unsafe("S"))
    .put(ref"A1", n(1))
    .put(ref"A2", n(2))
    .put(ref"A3", n(3))
    .put(ref"C1", n(10))
    .put(ref"C2", n(20))
    .put(ref"C3", n(30))

  private def cache(sheet: Sheet, ref: ARef): Option[CellValue] = sheet(ref).value match
    case CellValue.Formula(_, cached, _) => cached
    case _ => None

  private def resultSheet(wb: Workbook, name: SheetName = sheet.name): Sheet =
    wb(name).fold(error => fail(error.message), identity)

  private def names(wb: Workbook): Workbook =
    wb.withDefinedName("criteria", "IF(TRUE,S!A1:A3,S!A1:A3)")
      .withDefinedName("target", "IF(TRUE,S!C1,S!C1)")
      .withDefinedName("staticTarget", "S!C1")

  List(("SUMIF", 60, 240), ("AVERAGEIF", 20, 80)).foreach { case (function, initial, expected) =>
    val letFormulas = List(
      s"LET(r,IF(TRUE,A1:A3,A1:A3),$function(r,\">0\",C1))",
      s"LET(t,C1,$function(A1:A3,\">0\",t))",
      s"LET(r,A1:A3,t,C1,LET(u,t,$function(r,\">0\",u)))"
    )

    letFormulas.foreach { formula =>
      test(s"$formula refreshes resized cells and downstream consumers after a sheet edit") {
        val original = sheet
          .put(ref"E1", CellValue.Formula(formula, Some(n(initial))))
          .put(ref"F1", CellValue.Formula("E1*2", Some(n(initial * 2))))
        assertEquals(original.evaluateCell(ref"E1"), Right(n(initial)))
        val updated = original.put(ref"C2", n(200)).recalculateDependents(Set(ref"C2"))
        assertEquals(cache(updated, ref"E1"), Some(n(expected)))
        assertEquals(cache(updated, ref"F1"), Some(n(expected * 2)))
      }
    }

    val namedFormulas = List(
      s"$function(criteria,\">0\",C1)",
      s"$function(A1:A3,\">0\",target)",
      s"$function(A1:A3,\">0\",staticTarget)",
      s"LET(r,criteria,t,target,$function(r,\">0\",t))"
    )
    namedFormulas.foreach { formula =>
      test(s"$formula refreshes resized named cells after a workbook edit") {
        val original = sheet.put(ref"E1", CellValue.Formula(formula, Some(n(initial))))
        val wb = names(Workbook(original))
        assertEquals(original.evaluateCell(ref"E1", Clock.system, Some(wb)), Right(n(initial)))
        val updated = wb
          .put(original.put(ref"C2", n(200)))
          .recalculateDependents(original.name, Set(ref"C2"))
        assertEquals(cache(resultSheet(updated), ref"E1"), Some(n(expected)))
      }
    }

    List("Z2*10", "INDIRECT(\"Z2\")*10").foreach { targetFormula =>
      test(s"$function full recalc refreshes a resized target with stale $targetFormula cache") {
        val formula = s"LET(t,Y1,$function(A1:A3,\">0\",t))"
        val original = sheet
          .put(ref"Y1", n(10))
          .put(ref"Y2", CellValue.Formula(targetFormula, Some(n(20))))
          .put(ref"Y3", n(30))
          .put(ref"Z2", n(20))
          .put(ref"E1", CellValue.Formula(formula, Some(n(initial))))
          .put(ref"F1", CellValue.Formula("E1*2", Some(n(initial * 2))))
        val wb = Workbook(original)
        val sequential = wb.recalculate()
        val parallel = wb.recalculateParallel(4)
        assertEquals(sequential.errors, Vector.empty)
        assertEquals(parallel.errors, sequential.errors)
        assertEquals(parallel.evaluated, sequential.evaluated)
        assertEquals(parallel.workbook, sequential.workbook)
        val updated = resultSheet(sequential.workbook)
        assertEquals(cache(updated, ref"Y2"), Some(n(200)))
        assertEquals(cache(updated, ref"E1"), Some(n(expected)))
        assertEquals(cache(updated, ref"F1"), Some(n(expected * 2)))
        val checked = original.evaluateWithDependencyCheck().fold(e => fail(e.message), identity)
        assertEquals(checked.get(ref"E1"), Some(n(expected)))
      }
    }

    test(s"$function orders an indirectly edited resized target before its dynamic consumer") {
      val formula = s"LET(t,Y1,$function(A1:A3,\">0\",t))"
      val original = sheet
        .put(ref"Y1", n(10))
        .put(ref"Y2", CellValue.Formula("Z2*10", Some(n(20))))
        .put(ref"Y3", n(30))
        .put(ref"Z2", n(2))
        .put(ref"E1", CellValue.Formula(formula, Some(n(initial))))
      val edited = original.put(ref"Z2", n(20))
      val updated = edited.recalculateDependents(Set(ref"Z2"))
      assertEquals(cache(updated, ref"Y2"), Some(n(200)))
      assertEquals(cache(updated, ref"E1"), Some(n(expected)))
      val cone = DependentRecalculation.recalculateAfterEdit(
        Workbook(edited),
        sheet.name,
        Set(ref"Z2"),
        Clock.system
      )
      assertEquals(cone.errors, Vector.empty)
      assertEquals(cache(resultSheet(cone.workbook), ref"E1"), Some(n(expected)))
    }

    test(s"$function finds a resized cross-sheet target hidden behind a defined formula") {
      val calcName = SheetName.unsafe("Calc")
      val calc = Sheet(calcName)
        .put(ref"E1", CellValue.Formula("total", Some(n(initial))))
      val original = names(Workbook(Vector(sheet, calc)))
        .withDefinedName("total", s"$function(criteria,\">0\",target)")
      assertEquals(DependencyGraph.dynamicCells(original), Set(QualifiedRef(calcName, ref"E1")))
      val updated = original
        .put(sheet.put(ref"C2", n(200)))
        .recalculateDependents(sheet.name, Set(ref"C2"))
      assertEquals(cache(resultSheet(updated, calcName), ref"E1"), Some(n(expected)))
    }
  }

  test("literal resized ranges and calls without a target retain static scheduling") {
    val original = sheet
      .put(ref"E1", CellValue.Formula("SUMIF(A1:A3,\">0\",C1)", None))
      .put(ref"E2", CellValue.Formula("AVERAGEIF(A1:A3,\">0\",C1)", None))
      .put(ref"E3", CellValue.Formula("SUMIF(criteria,\">0\")", None))
      .put(ref"E4", CellValue.Formula("LET(r,A1:A3,AVERAGEIF(r,\">0\"))", None))
    assertEquals(DependencyGraph.dynamicCells(original), Set.empty[ARef])
    assertEquals(DependencyGraph.dynamicCells(names(Workbook(original))), Set.empty[QualifiedRef])
  }

  test("a self-reference hidden by resizing cannot reuse its old cache") {
    val formula = "LET(t,C1,SUMIF(A1:A3,\">0\",t))"
    val original = sheet.put(ref"C2", CellValue.Formula(formula, Some(n(60))))
    val result = Workbook(original).recalculate()
    assert(result.errors.exists(_.ref == ref"C2"), result.errors.toString)
    assertEquals(cache(resultSheet(result.workbook), ref"C2"), None)
  }
