package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.graph.{CompactGraphGens, DependencyGraph}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/**
 * #891: the graph questions the data-table seeder asks, answered by the recalculation graph
 * ([[RecalcGraph]]), equal the answers of the used-range-bounded per-cell graph the seeder built
 * before (`DependencyGraph.fromWorkbookBounded`) on books without fixed-column lookups, the one
 * place #504's precise reads differ: the precedent formulas of a table's corner never walking past
 * its inputs, the formulas downstream of its inputs (constants or formulas), and the cyclic core.
 */
class SeederGraphOracleSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(300)

  private def withoutLookups(wb: Workbook): Workbook =
    wb.copy(sheets = wb.sheets.map { s =>
      s.copy(cells = s.cells.view.mapValues { cell =>
        cell.value match
          case CellValue.Formula(text, _, _) if text.toUpperCase.contains("LOOKUP(") =>
            cell.copy(value = CellValue.Number(BigDecimal(1)))
          case _ => cell
      }.toMap)
    })

  private val genCell: Gen[QualifiedRef] =
    for
      sheet <- Gen.oneOf(CompactGraphGens.sheetNames).map(SheetName.unsafe)
      col <- Gen.choose(0, CompactGraphGens.Cols - 1)
      row <- Gen.choose(0, CompactGraphGens.Rows - 1)
    yield QualifiedRef(sheet, ARef.from0(col, row))

  private val genCase =
    for
      wb <- CompactGraphGens.genWorkbook.map(withoutLookups)
      seeds <- Gen.listOfN(3, genCell)
      stop <- Gen.listOfN(2, genCell)
    yield (wb, seeds.toSet, stop.toSet)

  /**
   * The book with every sheet qualifier spelled as the sheet is named. Sheet names are
   * case-insensitive; the bounded graph resolved `'MID'!A1` to no sheet (a missing edge, so a
   * missed dependent or cycle), the recalculation graph canonicalises it. Spelling the qualifiers
   * canonically leaves only the questions both answer.
   */
  private def canonicalQualifiers(wb: Workbook): Workbook =
    val names = wb.sheets.map(_.name.value)
    val qualifier = """'([^']+)'!""".r
    def fix(text: String): String =
      qualifier.replaceAllIn(
        text,
        m =>
          java.util.regex.Matcher.quoteReplacement(
            names.find(_.equalsIgnoreCase(m.group(1))).fold(m.matched)(n => s"'$n'!")
          )
      )
    wb.copy(sheets = wb.sheets.map { s =>
      s.copy(cells = s.cells.view.mapValues { cell =>
        cell.value match
          case CellValue.Formula(text, cached, kind) =>
            cell.copy(value = CellValue.Formula(fix(text), cached, kind))
          case _ => cell
      }.toMap)
    })

  private def inUsed(wb: Workbook, q: QualifiedRef): Boolean =
    wb(q.sheet).toOption.flatMap(_.usedRange).exists(_.contains(q.ref))

  property("precedents, dependents of cells and the cyclic core equal the bounded graph's") {
    Prop.forAll(genCase) { (generated, seeds, stop) =>
      val wb = canonicalQualifiers(generated)
      val (deps, dependents) = DependencyGraph.fromWorkbookBounded(wb)
      val graph = RecalcGraph.build(wb, RecalcCounters.disabled)
      val formulas = deps.keySet
      // the seeder takes the closure with its sources added (`closure ++ srcQ`): the new walk lists
      // a source a cycle leads back to, the old one never listed a starting cell
      val sources = seeds.filter(formulas.contains)
      val oldClosure =
        DependencyGraph.qualifiedTransitivePrecedents(deps -- stop, seeds).intersect(formulas) ++
          sources
      val newClosure = graph.precedents(sources, stop) ++ sources
      val oldMoved = DependencyGraph.qualifiedTransitiveDependents(dependents, stop)
      val newMoved = graph.cellDependents(stop)
      val oldCore = DependencyGraph.qualifiedCyclicNodes(deps)
      Prop.all(
        Prop(newClosure == oldClosure) :|
          s"closure new-only ${newClosure -- oldClosure} old-only ${oldClosure -- newClosure}",
        // the bounded graph clipped every range to its sheet's used range, so an empty cell outside
        // it had no readers (pinned below); inside the used range the two agree
        Prop(oldMoved.subsetOf(newMoved)) :| s"moved old-only ${oldMoved -- newMoved}",
        Prop(!stop.forall(inUsed(wb, _)) || newMoved == oldMoved) :|
          s"moved new-only ${newMoved -- oldMoved}",
        Prop(graph.cyclicCore == oldCore) :|
          s"core new-only ${graph.cyclicCore -- oldCore} old-only ${oldCore -- graph.cyclicCore}"
      )
    }
  }

  test("a lower-case sheet qualifier is the sheet: the recalculation graph sees the edge") {
    val wb = Workbook(
      com.tjclp.xl.sheets
        .Sheet(SheetName.unsafe("Mid"))
        .put(ARef.from0(0, 0), CellValue.Formula("'mid'!B1+1"))
        .put(ARef.from0(1, 0), CellValue.Formula("A1*2"))
    )
    val graph = RecalcGraph.build(wb, RecalcCounters.disabled)
    val (deps, _) = DependencyGraph.fromWorkbookBounded(wb)
    assertEquals(graph.cyclicCore.size, 2)
    // the bounded graph the seeder used before #891 missed the cycle
    assertEquals(DependencyGraph.qualifiedCyclicNodes(deps), Set.empty[QualifiedRef])
  }

  test("an empty cell outside the used range still has its range readers") {
    // row 1 is empty, so the used range starts at row 2; A5 reads rows 1:3 all the same
    val wb = Workbook(
      com.tjclp.xl.sheets
        .Sheet(SheetName.unsafe("S"))
        .put(ARef.from0(0, 1), CellValue.Number(BigDecimal(1)))
        .put(ARef.from0(0, 4), CellValue.Formula("COUNT(1:3)"))
    )
    val b1 = QualifiedRef(SheetName.unsafe("S"), ARef.from0(1, 0))
    val graph = RecalcGraph.build(wb, RecalcCounters.disabled)
    assertEquals(graph.cellDependents(Set(b1)), Set(QualifiedRef(b1.sheet, ARef.from0(0, 4))))
    // the bounded graph the seeder used before #891 saw no reader of B1
    val (_, dependents) = DependencyGraph.fromWorkbookBounded(wb)
    assertEquals(DependencyGraph.qualifiedTransitiveDependents(dependents, Set(b1)), Set.empty)
  }
