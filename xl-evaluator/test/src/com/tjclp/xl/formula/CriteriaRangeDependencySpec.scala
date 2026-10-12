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
    ("SUMIF\u2003", "SUMIF with EM SPACE", 60, 240),
    ("SUMIF\u2028", "SUMIF with LINE SEPARATOR", 60, 240),
    ("\u017fUMIF", "SUMIF with long s", 60, 240),
    ("SUM\u0131F", "SUMIF with dotless i", 60, 240),
    ("AVERAGEIF\u2003", "AVERAGEIF with EM SPACE", 20, 80),
    ("AVERAGE\u0131F", "AVERAGEIF with dotless i", 20, 80)
  ).foreach { (function, label, initial, expected) =>
    test(s"$label keeps resized named dependencies discoverable") {
      // The parser accepts these spellings as the canonical function, so its pre-filter must too.
      // #733: a computed name (`target`) and a LET name are dynamic; a name bound to a plain range
      // (`staticTarget`) resizes statically, and its resized edges keep the reader fresh as well
      val computed = s"$function(A1:A3,\">0\",target)"
      val bound = s"LET(t,C1,$function(A1:A3,\">0\",t))"
      val static = s"$function(A1:A3,\">0\",staticTarget)"
      val original = sheet
        .put(ref"E1", CellValue.Formula(computed, Some(n(initial))))
        .put(ref"E2", CellValue.Formula(bound, Some(n(initial))))
        .put(ref"E3", CellValue.Formula(static, Some(n(initial))))
      val wb = names(Workbook(original))
      List(ref"E1", ref"E2", ref"E3").foreach { at =>
        assertEquals(original.evaluateCell(at, Clock.system, Some(wb)), Right(n(initial)))
      }
      val updated = wb
        .put(original.put(ref"C2", n(200)))
        .recalculateDependents(original.name, Set(ref"C2"))
      List(ref"E1", ref"E2", ref"E3").foreach { at =>
        assertEquals(cache(resultSheet(updated), at), Some(n(expected)), at.toA1)
      }
      List(computed, bound, static).foreach(f =>
        assert(CriteriaRangeResize.mayResizeThroughName(f))
      )
      assertEquals(DependencyGraph.dynamicCells(original), Set(ref"E2"))
      assertEquals(
        DependencyGraph.dynamicCells(wb),
        Set(QualifiedRef(sheet.name, ref"E1"), QualifiedRef(sheet.name, ref"E2"))
      )
    }
  }

  List(
    "SUMIF(A1:A3,\">0\",amt)",
    "SUMIF(amt,\">0\")",
    "sumif(A1:A3,\">0\",amt)",
    "AVERAGEIF (A1:A3, 1, amt)",
    "1+SUMIF(criteria,\">0\",C1)",
    "SUMIF(Sheet1!amt,1,C1)",
    "SUMIF(A1:A3,\">0\",Model!amt)",
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
    // without a workbook a defined name resolves to nothing, so the reader reads nothing through it
    assertEquals(DependencyGraph.dynamicCells(withReader), Set.empty[ARef])
    val withLet = literal.put(ref"E1", CellValue.Formula("LET(t,C1,SUMIF(A1:A3,\">0\",t))", None))
    assertEquals(DependencyGraph.dynamicCells(withLet), Set(ref"E1"))
  }

  // ----- #733: names bound to plain ranges resize statically; only computed geometry is dynamic -----

  private val S = sheet.name
  private def q(at: ARef): QualifiedRef = QualifiedRef(S, at)

  private def staticNames(wb: Workbook): Workbook =
    wb.withDefinedName("Categories", "S!$A$1:$A$3")
      .withDefinedName("Amounts", "S!$C$1:$C$3")
      .withDefinedName("staticTarget", "S!$C$1")
      .withDefinedName("alias", "staticTarget")
      .withDefinedName("offsetTarget", "OFFSET(S!$C$1,0,0)")

  // (formula, value before, value after C2 := 200)
  private val staticReaders = List(
    ref"E1" -> ("SUMIF(Categories,\">1\",Amounts)", 50, 230),
    ref"E2" -> ("SUMIF(A1:A3,\">0\",staticTarget)", 60, 240),
    ref"E3" -> ("AVERAGEIF(Categories,\">0\",staticTarget)", 20, 80),
    ref"E4" -> ("SUMIF(Categories,\">0\",alias)", 60, 240),
    ref"E5" -> ("SUMIF(S!Categories,\">0\",S!staticTarget)", 60, 240)
  )

  private def staticBook: Workbook =
    staticNames(Workbook(staticReaders.foldLeft(sheet) { case (s, (at, (formula, before, _))) =>
      s.put(at, CellValue.Formula(formula, Some(n(before))))
    }))

  test("#733: a SUMIF over names bound to plain ranges is static, with its resized edges") {
    val wb = staticBook
    assertEquals(DependencyGraph.dynamicCells(wb), Set.empty[QualifiedRef])
    assertEquals(com.tjclp.xl.formula.eval.WorkbookAudit.of(wb).dynamic, Vector.empty)
    val graph = DependencyGraph.fromWorkbook(wb)
    // staticTarget names C1, but the call reads C1:C3: the graph records the resized cells
    List(ref"E2", ref"E3", ref"E4", ref"E5").foreach { at =>
      val deps = graph.getOrElse(q(at), Set.empty)
      assert(Set(q(ref"C1"), q(ref"C2"), q(ref"C3")).subsetOf(deps), s"$at: $deps")
    }
    assert(Set(q(ref"C2"), q(ref"C3")).subsetOf(graph.getOrElse(q(ref"E1"), Set.empty)))
  }

  test("#733: an edit inside the resized target refreshes the static readers") {
    val wb = staticBook
    val viaDependents = wb
      .put(resultSheet(wb).put(ref"C2", n(200)))
      .recalculateDependents(S, Set(ref"C2"))
    val cone = DependentRecalculation.recalculateAfterEdit(
      wb.put(resultSheet(wb).put(ref"C2", n(200))),
      S,
      Set(ref"C2"),
      Clock.system
    )
    assertEquals(cone.errors, Vector.empty)
    staticReaders.foreach { case (at, (formula, _, after)) =>
      assertEquals(cache(resultSheet(viaDependents), at), Some(n(after)), formula)
      assertEquals(cache(resultSheet(cone.workbook), at), Some(n(after)), formula)
    }
  }

  test("#733: full recalc orders an uncached formula in the resized target before its reader") {
    val withFormula = resultSheet(staticBook)
      .put(ref"C2", CellValue.Formula("Z2*10", Some(n(20))))
      .put(ref"Z2", n(20))
    val wb = staticBook.put(withFormula)
    val sequential = wb.recalculate()
    val parallel = wb.recalculateParallel(4)
    assertEquals(sequential.errors, Vector.empty)
    assertEquals(parallel.errors, sequential.errors)
    assertEquals(parallel.evaluated, sequential.evaluated)
    assertEquals(parallel.workbook, sequential.workbook)
    staticReaders.foreach { case (at, (formula, _, after)) =>
      assertEquals(cache(resultSheet(sequential.workbook), at), Some(n(after)), formula)
    }
  }

  test("#733: a static self-read through a resized name is a cycle") {
    val looped = staticNames(
      Workbook(
        sheet.put(ref"C2", CellValue.Formula("SUMIF(A1:A3,\">0\",staticTarget)", Some(n(60))))
      )
    )
    assertEquals(DependencyGraph.dynamicCells(looped), Set.empty[QualifiedRef])
    val result = looped.recalculate()
    assert(result.errors.exists(_.ref == ref"C2"), result.errors.toString)
    assertEquals(cache(resultSheet(result.workbook), ref"C2"), None)
  }

  test("#733: LET-bound and computed names keep the dynamic schedule") {
    val wb = staticNames(
      Workbook(
        sheet
          .put(ref"E1", CellValue.Formula("SUMIF(A1:A3,\">0\",offsetTarget)", Some(n(60))))
          .put(ref"E2", CellValue.Formula("LET(t,staticTarget,SUMIF(A1:A3,\">0\",t))", Some(n(60))))
          .put(ref"E3", CellValue.Formula("SUMIF(Categories,\">0\",staticTarget)", Some(n(60))))
      )
    )
    assertEquals(DependencyGraph.dynamicCells(wb), Set(q(ref"E1"), q(ref"E2")))
    val updated = wb
      .put(resultSheet(wb).put(ref"C2", n(200)))
      .recalculateDependents(S, Set(ref"C2"))
    List(ref"E1", ref"E2", ref"E3").foreach { at =>
      assertEquals(cache(resultSheet(updated), at), Some(n(240)), at.toA1)
    }
  }

  test("#733: a data table over a static named SUMIF seeds again") {
    import com.tjclp.xl.cells.FormulaKind
    import com.tjclp.xl.formula.eval.SeedTableWarning
    val kind: FormulaKind.DataTable =
      FormulaKind.DataTable(ref"H10:H12", dt2D = false, dtr = false, r1 = Some(ref"A2"), r2 = None)
    val book = staticNames(
      Workbook(
        sheet
          .put(ref"H9", CellValue.Formula("SUMIF(Categories,\">1\",Amounts)"))
          .put(ref"G10", n(0))
          .put(ref"G11", n(5))
          .put(ref"G12", n(9))
          .put(ref"H10", CellValue.dataTable(kind, None))
      )
    )
    val report = book.seedDataTablesReport().fold(error => fail(error.message), identity)
    assert(
      !report.warnings.exists {
        case SeedTableWarning.Skipped(_, _, _, _) => true
        case _ => false
      },
      report.warnings.toString
    )
    // A2 takes each axis value: 0 drops C2 (30), 5 and 9 keep it (50)
    val seeded = resultSheet(report.workbook)
    List(ref"H10" -> 30, ref"H11" -> 50, ref"H12" -> 50).foreach { (at, expected) =>
      // the anchor keeps its dataTable formula; the other interiors are plain values
      val seededValue = seeded(at).value match
        case CellValue.Formula(_, cached, _) => cached
        case CellValue.Empty => None
        case other => Some(other)
      assertEquals(seededValue, Some(n(expected)), at.toA1)
    }
  }
