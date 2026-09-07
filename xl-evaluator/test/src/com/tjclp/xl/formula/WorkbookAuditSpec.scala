package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.formula.eval.WorkbookAudit
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{CalcPr, Workbook}
import munit.FunSuite

/**
 * ADR-017 §2.10: `WorkbookAudit.of(wb)` buckets every reason a number can be wrong, from the
 * existing analyses, in one pass. `isClean` counts error values, uncached formulas, unparseable
 * formulas, cycles and unresolved names; volatile, dynamic and external references and `calcPr` are
 * reported but are not findings.
 */
class WorkbookAuditSpec extends FunSuite:

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private def cachedFormula(expr: String, value: Int): CellValue =
    CellValue.Formula(expr, Some(num(value)))
  private def a1(s: String): ARef = ARef.parse(s).fold(err => fail(err), identity)
  private def q(sheet: String, ref: String): QualifiedRef =
    QualifiedRef(SheetName.unsafe(sheet), a1(ref))

  private def sheetWith(name: String, cells: (String, CellValue)*): Sheet =
    cells.foldLeft(Sheet(SheetName.unsafe(name))) { case (s, (ref, value)) =>
      s.put(a1(ref), value)
    }

  private val iterative = CalcPr(
    iterativeCalculation = true,
    maxIterations = Some(100),
    maxChange = Some(BigDecimal("0.001"))
  )

  /** One cell per bucket, and nothing that lands in two buckets; `calcPr` left as read. */
  private val dirtyBook: Workbook = Workbook(
    Vector(
      sheetWith(
        "Calc",
        "A1" -> CellValue.Formula("1/0", Some(CellValue.Error(CellError.Div0))),
        "B1" -> CellValue.Formula("A1+1", None),
        "C1" -> cachedFormula("TODAY()", 45000),
        "D1" -> cachedFormula("INDIRECT(\"A2\")", 0),
        "E1" -> cachedFormula("F1+1", 0),
        "F1" -> cachedFormula("E1+1", 0),
        "G1" -> cachedFormula("'[1]Ext'!A1", 3),
        "H1" -> cachedFormula("UNSUPPORTED(1)", 0),
        "I1" -> cachedFormula("NoSuchName*2", 0)
      )
    )
  )

  /** The dirty book as an intentional circular model: iterative calculation declared. */
  private val dirty: Workbook = dirtyBook.withCalcPr(iterative)

  test("the dirty fixture populates exactly one entry per bucket") {
    val audit = WorkbookAudit.of(dirtyBook)
    assertEquals(audit.errorCells, Vector((q("Calc", "A1"), CellError.Div0)))
    assertEquals(audit.uncachedFormulas, Vector(q("Calc", "B1")))
    assertEquals(audit.volatile, Vector(q("Calc", "C1")))
    assertEquals(audit.dynamic, Vector(q("Calc", "D1")))
    assertEquals(audit.cycles.map(_.members.toSet), Vector(Set(q("Calc", "E1"), q("Calc", "F1"))))
    assert(audit.cycles.forall(_.cyclic))
    assertEquals(audit.iterativeCycles, Vector.empty)
    assertEquals(audit.externalRefs, Vector(q("Calc", "G1")))
    assertEquals(audit.unparseable.map(_._1), Vector(q("Calc", "H1")))
    assertEquals(audit.unresolvedReaders, Vector(q("Calc", "I1")))
    assertEquals(audit.calcPr, None)
    assertEquals(audit.isClean, false)
    assertEquals(audit.findings, 5)
  }

  test("with iterative calculation on, cycles are a note (iterativeCycles), not a finding") {
    val audit = WorkbookAudit.of(dirty)
    assertEquals(audit.cycles, Vector.empty)
    assertEquals(
      audit.iterativeCycles.map(_.members.toSet),
      Vector(Set(q("Calc", "E1"), q("Calc", "F1")))
    )
    assert(audit.iterativeCycles.forall(_.cyclic))
    assertEquals(audit.calcPr, Some(iterative))
    assertEquals(audit.findings, 4, "the cycle no longer counts")
    // every other bucket is exactly as without the calcPr
    val plain = WorkbookAudit.of(dirtyBook)
    assertEquals(
      audit.copy(cycles = plain.cycles, iterativeCycles = Vector.empty, calcPr = None),
      plain
    )
    // an intentional circular model with nothing else wrong is clean
    val circular = Workbook(
      Vector(
        sheetWith("Model", "A1" -> cachedFormula("B1*0.1", 1), "B1" -> cachedFormula("100+A1/2", 2))
      )
    )
    assertEquals(WorkbookAudit.of(circular).isClean, false)
    assertEquals(WorkbookAudit.of(circular).cycles.size, 1)
    val declared = WorkbookAudit.of(circular.withCalcPr(iterative))
    assert(declared.isClean, s"iterative model must be clean: $declared")
    assertEquals(declared.iterativeCycles.size, 1)
    // iterativeCalculation = false in an explicit calcPr is the same as no calcPr
    val off = WorkbookAudit.of(circular.withCalcPr(CalcPr(iterativeCalculation = false)))
    assertEquals(off.cycles.size, 1)
    assertEquals(off.iterativeCycles, Vector.empty)
  }

  test("an unparseable formula carries the parser's diagnostic with context") {
    val audit = WorkbookAudit.of(dirty)
    val message = audit.unparseable.headOption.map(_._2).getOrElse(fail("no unparseable entry"))
    assert(message.startsWith("UNSUPPORTED(1)"), message)
    assert(message.contains("UNSUPPORTED"), message)
    assert(message.linesIterator.size >= 2, s"expected the formula and a diagnostic line: $message")
  }

  test("a clean book is clean: every bucket empty, calcPr None") {
    val clean = Workbook(
      Vector(sheetWith("S", "A1" -> num(1), "A2" -> cachedFormula("A1*2", 2)))
    )
    val audit = WorkbookAudit.of(clean)
    assertEquals(audit, WorkbookAudit.clean)
    assert(audit.isClean)
    assertEquals(audit.findings, 0)
    // and a recalculated book with a volatile note is still clean
    val volatile = Workbook(Vector(sheetWith("S", "A1" -> CellValue.Formula("TODAY()", None))))
      .recalculate()
      .workbook
    val note = WorkbookAudit.of(volatile)
    assert(note.isClean, s"volatile is a note, not a finding: $note")
    assertEquals(note.volatile, Vector(q("S", "A1")))
  }

  test("a bare error value counts as an error cell; a dynamic reader is not an unresolved name") {
    val wb = Workbook(
      Vector(
        sheetWith(
          "S",
          "A1" -> CellValue.Error(CellError.NA),
          "B1" -> cachedFormula("OFFSET(A1,0,0)", 0)
        )
      )
    )
    val audit = WorkbookAudit.of(wb)
    assertEquals(audit.errorCells, Vector((q("S", "A1"), CellError.NA)))
    assertEquals(audit.dynamic, Vector(q("S", "B1")))
    assertEquals(audit.unresolvedReaders, Vector.empty)
    assertEquals(audit.unparseable, Vector.empty)
  }

  test("a data-table record without a cache is uncached but never parsed") {
    val kind: FormulaKind.DataTable = FormulaKind.DataTable(
      CellRange.parse("F2:G2").fold(err => fail(err), identity),
      dt2D = true,
      dtr = false,
      r1 = Some(a1("A1")),
      r2 = Some(a1("A2"))
    )
    val wb = Workbook(
      Vector(
        sheetWith(
          "DT",
          "A1" -> num(8),
          "A2" -> num(3),
          "F2" -> CellValue.dataTable(kind, None),
          "G2" -> CellValue.dataTable(kind, Some(num(7)))
        )
      )
    )
    val audit = WorkbookAudit.of(wb)
    assertEquals(audit.uncachedFormulas, Vector(q("DT", "F2")))
    assertEquals(audit.unparseable, Vector.empty)
    assertEquals(audit.unresolvedReaders, Vector.empty)
  }

  test("buckets are in workbook order, then row, then column") {
    val wb = Workbook(
      Vector(
        sheetWith(
          "Zed",
          "B2" -> CellValue.Formula("1", None),
          "A1" -> CellValue.Formula("2", None)
        ),
        sheetWith(
          "Alpha",
          "C1" -> CellValue.Formula("3", None),
          "A10" -> CellValue.Formula("4", None)
        )
      )
    )
    assertEquals(
      WorkbookAudit.of(wb).uncachedFormulas,
      Vector(q("Zed", "A1"), q("Zed", "B2"), q("Alpha", "C1"), q("Alpha", "A10"))
    )
  }

  test("cycles are in workbook order too, not the graph's dependency-first order") {
    // First's cycle READS Second's cycle, so the SCC condensation puts Second's component first;
    // the audit lists First's first, by the position of its earliest member.
    val wb = Workbook(
      Vector(
        sheetWith(
          "First",
          "A1" -> cachedFormula("B1+Second!A1", 0),
          "B1" -> cachedFormula("A1", 0)
        ),
        sheetWith("Second", "A1" -> cachedFormula("B1", 0), "B1" -> cachedFormula("A1", 0))
      )
    )
    val cycles = WorkbookAudit.of(wb).cycles
    assertEquals(
      cycles.map(_.members.toSet),
      Vector(Set(q("First", "A1"), q("First", "B1")), Set(q("Second", "A1"), q("Second", "B1")))
    )
  }

  test("restrictTo keeps one sheet's findings and the cycles that touch it") {
    val wb = Workbook(
      Vector(
        sheetWith("P", "A1" -> CellValue.Formula("Q!A1+1", None), "B1" -> cachedFormula("Q!B1", 0)),
        sheetWith("Q", "A1" -> CellValue.Formula("P!A1+1", None), "B1" -> cachedFormula("P!B1", 0)),
        sheetWith("R", "A1" -> CellValue.Formula("1", None))
      )
    )
    val audit = WorkbookAudit.of(wb)
    assertEquals(audit.uncachedFormulas.size, 3)
    assertEquals(audit.cycles.size, 2)
    val onlyR = audit.restrictTo(SheetName.unsafe("R"))
    assertEquals(onlyR.uncachedFormulas, Vector(q("R", "A1")))
    assertEquals(onlyR.cycles, Vector.empty)
    val onlyP = audit.restrictTo(SheetName.unsafe("P"))
    assertEquals(onlyP.uncachedFormulas, Vector(q("P", "A1")))
    assertEquals(onlyP.cycles.size, 2, "both cycles have a member on P")
    assertEquals(onlyP.calcPr, audit.calcPr)
    val nowhere = audit.restrictTo(SheetName.unsafe("Missing"))
    assert(nowhere.isClean)
    // the iterative twin restricts its note the same way
    val noted = WorkbookAudit.of(wb.withCalcPr(iterative))
    assertEquals(noted.iterativeCycles.size, 2)
    assertEquals(noted.restrictTo(SheetName.unsafe("R")).iterativeCycles, Vector.empty)
    assertEquals(noted.restrictTo(SheetName.unsafe("P")).iterativeCycles.size, 2)
  }

  test("the volatile function set is exactly TODAY, NOW, RAND and RANDBETWEEN, by name") {
    assertEquals(WorkbookAudit.volatileFunctions, Set("TODAY", "NOW", "RAND", "RANDBETWEEN"))
    val wb = Workbook(
      Vector(
        sheetWith(
          "S",
          "A1" -> cachedFormula("IF(A2>0, NOW(), 0)", 0),
          "A2" -> cachedFormula("rand()*10", 0),
          "A3" -> cachedFormula("RANDBETWEEN(1,6)", 0),
          "A4" -> cachedFormula("ROUND(A2,0)", 0),
          "A5" -> cachedFormula("LET(x, TODAY(), x+1)", 0)
        )
      )
    )
    assertEquals(
      WorkbookAudit.of(wb).volatile,
      Vector(q("S", "A1"), q("S", "A2"), q("S", "A3"), q("S", "A5"))
    )
  }
