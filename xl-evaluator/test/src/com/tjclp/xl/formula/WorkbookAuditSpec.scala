package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.formula.eval.WorkbookAudit
import com.tjclp.xl.formula.functions.FunctionRegistry
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

  test("an unparseable formula carries the parser's diagnostic on one line, as lint does (#676)") {
    val audit = WorkbookAudit.of(dirty)
    val message = audit.unparseable.headOption.map(_._2).getOrElse(fail("no unparseable entry"))
    assertEquals(message, "UNSUPPORTED(1): Unknown function 'UNSUPPORTED' at position 0")
  }

  test("#676: a long unparseable formula is quoted to 80 characters, never echoed whole") {
    val long = "NOSUCHFN(" + (1 to 60).map(i => s"A$i").mkString("+") + ")"
    val book = Workbook(Vector(sheetWith("S", "A1" -> cachedFormula(long, 0))))
    val message =
      WorkbookAudit.of(book).unparseable.headOption.map(_._2).getOrElse(fail("no entry"))
    assert(message.startsWith(long.take(80) + "…: "), message)
    assert(!message.contains(long), message)
    assertEquals(message.linesIterator.size, 1, message)
  }

  test("#676: a formula with line breaks or tabs is still ONE line in the audit") {
    // Alt+Enter in Excel's formula bar is stored as a newline inside <f>
    val broken = "SUM(A1,\nNOSUCHFN(1)\r\n)\t+"
    val book = Workbook(Vector(sheetWith("S", "A1" -> cachedFormula(broken, 0))))
    val message =
      WorkbookAudit.of(book).unparseable.headOption.map(_._2).getOrElse(fail("no entry"))
    assertEquals(message.linesIterator.size, 1, message)
    assert(!message.exists(c => c == '\n' || c == '\r' || c == '\t'), message)
    assert(message.startsWith("SUM(A1, NOSUCHFN(1) ) +: "), message)
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

  // ===== #678: data-table interiors stale against their corner formula =====

  private def columnTable(interior: String, input: String): FormulaKind.DataTable =
    FormulaKind.DataTable(
      CellRange.parse(interior).fold(err => fail(err), identity),
      dt2D = false,
      dtr = false,
      r1 = Some(a1(input)),
      r2 = None
    )

  /**
   * A2 is the input, F9 the corner (`A2*factor`), E10:E12 the axis 1..3, F10:F12 the interior with
   * caches `cached(i)` — a record at F10, plain values below, as Excel writes a column table.
   */
  private def tableBook(factor: Int, cached: Int => Int): Workbook =
    val kind = columnTable("F10:F12", "A2")
    Workbook(
      Vector(
        sheetWith(
          "S",
          "A2" -> num(4),
          "F9" -> cachedFormula(s"A2*$factor", 4 * factor),
          "E10" -> num(1),
          "E11" -> num(2),
          "E12" -> num(3),
          "F10" -> CellValue.dataTable(kind, Some(num(cached(1)))),
          "F11" -> num(cached(2)),
          "F12" -> num(cached(3))
        )
      )
    )

  test("#678: interiors that match the corner re-evaluated at their inputs carry no note") {
    val audit = WorkbookAudit.of(tableBook(2, i => i * 2))
    assertEquals(audit.staleDataTables, Vector.empty)
  }

  test("#678: an interior left behind by a model change is a data-table-stale NOTE") {
    // the corner now triples, but the interior still holds the ×2 results — what
    // `xl put` leaves on an iterative book (parity with Excel under autoNoTable)
    val audit = WorkbookAudit.of(tableBook(3, i => i * 2))
    val stale = audit.staleDataTables
    assertEquals(stale.map(t => (t.sheet.value, t.ref.toA1)), Vector(("S", "F10:F12")))
    val table = stale.head
    assertEquals(table.sampled.map(_.toA1), Vector("F10", "F11", "F12"))
    assertEquals(
      table.stale.map((r, c, n) => (r.toA1, c, n)),
      Vector(
        ("F10", num(2), num(3)),
        ("F11", num(4), num(6)),
        ("F12", num(6), num(9))
      )
    )
    assert(table.render.contains("xl recalc --tables"), table.render)
    assert(table.render.contains("3 of 3 sampled"), table.render)
    // a note, never a finding: the file is valid
    assert(audit.isClean, s"stale interiors must not make the book dirty: $audit")
    assertEquals(audit.findings, 0)
    assertEquals(audit.restrictTo(SheetName.unsafe("S")).staleDataTables, stale)
    assertEquals(audit.restrictTo(SheetName.unsafe("Other")).staleDataTables, Vector.empty)
  }

  test("#678: the check is SAMPLED — a bounded, named set of interior cells per table") {
    val kind = columnTable("F10:F1009", "A2")
    val axis = (0 until 1000).map(i => s"E${10 + i}" -> num(i))
    val interior = (1 until 1000).map(i => s"F${10 + i}" -> num(-1))
    val book = Workbook(
      Vector(
        sheetWith(
          "S",
          (Vector("A2" -> num(4), "F9" -> cachedFormula("A2*2", 8)) ++ axis ++ interior ++
            Vector("F10" -> CellValue.dataTable(kind, Some(num(-1)))))*
        )
      )
    )
    val table = WorkbookAudit.of(book).staleDataTables.headOption.getOrElse(fail("no note"))
    assertEquals(table.sampled.size, WorkbookAudit.DataTableSampleSize)
    assertEquals(table.sampled.headOption.map(_.toA1), Some("F10"))
    assertEquals(table.sampled.lastOption.map(_.toA1), Some("F1009"))
    assertEquals(table.sampled.distinct, table.sampled)
    assertEquals(table.stale.size, table.sampled.size)
  }

  test("#678: floating-point noise between Excel's cache and xl's re-evaluation is not stale") {
    val kind = columnTable("F10:F10", "A2")
    val book = Workbook(
      Vector(
        sheetWith(
          "S",
          "A2" -> num(0),
          "F9" -> cachedFormula("A2+0.2", 0),
          "E10" -> CellValue.Number(BigDecimal("0.1")),
          "F10" -> CellValue.dataTable(
            kind,
            Some(CellValue.Number(BigDecimal("0.30000000000000004")))
          )
        )
      )
    )
    assertEquals(WorkbookAudit.of(book).staleDataTables, Vector.empty)
  }

  test("#678: Excel-authored data tables (all three shapes) re-evaluate to their own caches") {
    val wb = com.tjclp.xl.ooxml.XlsxReader
      .read(com.tjclp.xl.ooxml.TestFixtures.copyToTemp("datatable-excel.xlsx"))
      .fold(err => fail(err.message), identity)
    assertEquals(WorkbookAudit.of(wb).staleDataTables, Vector.empty)
  }

  /** [[tableBook]] with the corner (and optionally a cone cell B2) given verbatim. */
  private def volatileTableBook(corner: String, cone: Option[String]): Workbook =
    val kind = columnTable("F10:F12", "A2")
    Workbook(
      Vector(
        sheetWith(
          "S",
          (Vector(
            "A2" -> num(4),
            "F9" -> cachedFormula(corner, 400),
            "E10" -> num(1),
            "E11" -> num(2),
            "E12" -> num(3),
            "F10" -> CellValue.dataTable(kind, Some(num(-1))),
            "F11" -> num(-1),
            "F12" -> num(-1)
          ) ++ cone.map(f => "B2" -> cachedFormula(f, 4)))*
        )
      )
    )

  test("#678: a table whose corner is volatile is never stale — the note stays deterministic") {
    val seeded = volatileTableBook("A2*100+RAND()", None)
      .seedDataTables()
      .fold(err => fail(err.message), identity)
    val first = WorkbookAudit.of(seeded)
    assertEquals(first.staleDataTables, Vector.empty)
    assertEquals(WorkbookAudit.of(seeded), first)
    // the volatile bucket still names the corner, so the table is not silently trusted
    assertEquals(first.volatile, Vector(q("S", "F9")))
  }

  test("#678: a volatile call anywhere in the corner's cone exempts the table too") {
    val book = volatileTableBook("B2*100", Some("A2+RAND()"))
    val seeded = book.seedDataTables().fold(err => fail(err.message), identity)
    assertEquals(WorkbookAudit.of(seeded).staleDataTables, Vector.empty)
    assertEquals(WorkbookAudit.of(seeded), WorkbookAudit.of(seeded))
    // the same shape without the volatile call is still checked (the caches are -1: stale)
    assertEquals(
      WorkbookAudit.of(volatileTableBook("B2*100", Some("A2+1"))).staleDataTables.map(_.ref.toA1),
      Vector("F10:F12")
    )
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

  test("GH-588: volatility is the FunctionFlags.volatile flag on the specs, not a name list") {
    // the flag is the single source of truth: the registry derives the names from it, the audit
    // reads the flag off the parsed Call, and the two agree
    assertEquals(
      FunctionRegistry.volatileFunctionNames,
      List("NOW", "RAND", "RANDBETWEEN", "TODAY")
    )
    assertEquals(WorkbookAudit.volatileFunctions, FunctionRegistry.volatileFunctionNames.toSet)
    FunctionRegistry.volatileFunctionNames.foreach { name =>
      assert(FunctionRegistry.lookup(name).exists(_.flags.volatile), s"$name must be flagged")
    }
    Vector("SUM", "INDIRECT", "OFFSET", "CELL", "IF").foreach { name =>
      assert(FunctionRegistry.lookup(name).exists(!_.flags.volatile), s"$name must not be flagged")
    }
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
