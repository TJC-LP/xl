package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.DependentRecalculation
import com.tjclp.xl.formula.functions.CriteriaRangeResize
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

  // ----- the text pre-filter: sound for names and LET names, silent for everything else -----

  List(
    "SUMIF(A1:A3,\">0\",amt)",
    "SUMIF(amt,\">0\")",
    "sumif(A1:A3,\">0\",amt)",
    "AVERAGEIF (A1:A3, 1, amt)",
    "1+SUMIF(criteria,\">0\",C1)",
    "SUMIF(Sheet1!amt,1,C1)",
    "SUMIF('My Sheet'!amt,1,C1)",
    "SUMIF(A1:A3,1,Sales.Total)",
    "SUMIF(A1:A3,1,_amt)",
    "SUMIF(A1:A3,1,\\amt)",
    "SUMIF(A1:A3,1,ümsatz)",
    // past column XFD: not a cell, so the parser reads a name
    "SUMIF(A1:A3,1,XFE1)",
    "LET(t,C1,SUMIF(A1:A3,\">0\",t))",
    "_xlfn.LET(_xlpm.t,C1,SUMIF(A1:A3,\">0\",_xlpm.t))",
    "SUMIF(A1:A3,\"x\",C1)+LET(t,C1,SUMIF(A1:A3,\">0\",t))",
    "SUMIFS(C1:C3,A1:A3,1)+SUMIF(A1:A3,1,amt)"
  ).foreach { formula =>
    test(s"pre-filter keeps =$formula, which may resize through a name") {
      assert(CriteriaRangeResize.mayResizeThroughName(formula))
    }
  }

  List(
    "SUMIFS(C1:C3,amt,\">0\")",
    "AVERAGEIFS(amt,A1:A3,1)",
    "SUM(amt)",
    "SUMIF(A1:A3,\">0\",C1:C3)",
    "SUMIF($A$1:$A$10,$B2,$C$1)",
    "SUMIF(A:A,\"x\",B:B)",
    "SUMIF($A:$A,\"x\",$B:$B)",
    "SUMIF(1:3,\"x\",5:7)",
    "SUMIF(A1:A3,\"amt\",C1)",
    "SUMIF(A1:A3,TRUE,C1)",
    "SUMIF(Sheet1!A1:A3,\">0\",Sheet1!C1)",
    "SUMIF('amt sheet'!A1:A3,\">0\",'amt sheet'!C1)",
    "SUMIF('O''Brien amt'!A1:A3,\">0\",C1)",
    "ROUND(SUMIF(A1:A3,\">0\",C1),2)",
    "SUMIF(A1:A3,#N/A,C1)+SUMIF(A1:A3,1.5E3,C1)",
    "SUMIF(A1:A3,\">0\",C1#)",
    "IFERROR(SUMIF(A1:A3,\"\"\"amt\"\"\",C1),0)"
  ).foreach { formula =>
    test(s"pre-filter skips =$formula, which cannot resize through a name") {
      assert(!CriteriaRangeResize.mayResizeThroughName(formula))
    }
  }

  test("dynamic-cell discovery parses no SUMIFS and no SUMIF over written ranges") {
    val literal = (1 to 50).foldLeft(sheet) { (s, i) =>
      s.put(ARef.from0(5, i), CellValue.Formula(s"SUMIFS(C1:C3,A1:A3,\">$i\")", None))
        .put(ARef.from0(6, i), CellValue.Formula(s"SUMIF(A1:A3,\">$i\",C1)", None))
        .put(ARef.from0(7, i), CellValue.Formula(s"AVERAGEIFS(C1:C3,criteria,$i)", None))
    }
    val wb = names(Workbook(literal))
    val trace = DependencyGraph.dynamicCellsTraced(wb)
    assertEquals(trace.cells, Set.empty[QualifiedRef])
    // the three definitions parse; no cell does
    assertEquals(trace.parses, 3)
    assertEquals(DependencyGraph.dynamicCells(literal), Set.empty[ARef])

    // one resizing reader through a name is found, and is the one cell parsed
    val withReader = literal.put(ref"E1", CellValue.Formula("SUMIF(A1:A3,\">0\",target)", None))
    val traced = DependencyGraph.dynamicCellsTraced(names(Workbook(withReader)))
    assertEquals(traced.cells, Set(QualifiedRef(sheet.name, ref"E1")))
    assertEquals(traced.parses, 4)
    assertEquals(DependencyGraph.dynamicCells(withReader), Set(ref"E1"))
  }
