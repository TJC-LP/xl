package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{DefinedName, Workbook}
import munit.FunSuite

/**
 * #891: the after-edit pass against a full recalculation on hand-built books, one per shape the
 * cone could get wrong: lookups whose selector is a cell, a literal that needs truncating or
 * coercing, or an expression; tables reached through a name or another sheet; INDEX folded to one
 * cell; dynamic readers; defined names; whole-column readers and edits outside the used range;
 * cycles an edit creates or removes; spills; failures that appear and heal; criteria ranges.
 *
 * Every case settles a book with a full recalculation, applies an edit, and asserts that the
 * after-edit pass leaves every cell of every sheet, and the set of cells it reports as errors,
 * exactly as a full recalculation of the edited book does. So that a case cannot pass by the edit
 * reaching nothing, each also asserts that the edit moved at least one formula of its cone off its
 * settled value. [[AfterEditLawSpec]] holds the same law over generated books.
 */
class AfterEditAttackSpec extends FunSuite:

  private val clock = Clock.fixedDate(java.time.LocalDate.of(2026, 1, 2))

  private def options(seed: Long = 5L) =
    RecalcOptions(clock = clock, rng = Rng.seeded(seed), iterative = IterativeMode.Off)

  private def r(a1: String): ARef = ARef.parse(a1).fold(fail(_), identity)
  private def n(i: Double): CellValue = CellValue.Number(BigDecimal(i))
  private def f(text: String): CellValue = CellValue.Formula(text)
  private def sn(name: String): SheetName = SheetName.unsafe(name)
  private def q(sheet: String, a1: String): QualifiedRef = QualifiedRef(sn(sheet), r(a1))

  private def sheet(name: String, cells: (String, CellValue)*): Sheet =
    cells.foldLeft(Sheet(sn(name)))((s, kv) => s.put(r(kv._1), kv._2))

  private def set(wb: Workbook, name: String, a1: String, value: Option[CellValue]): Workbook =
    val target = wb(sn(name)).fold(e => fail(e.message), identity)
    wb.put(value.fold(target.remove(r(a1)))(v => target.put(Cell(r(a1), v))))

  private def named(wb: Workbook, names: DefinedName*): Workbook =
    wb.copy(metadata = wb.metadata.copy(definedNames = names.toVector))

  private def cellsOf(wb: Workbook): Map[QualifiedRef, CellValue] =
    wb.sheets.iterator
      .flatMap(s => s.cells.iterator.map((ref, c) => QualifiedRef(s.name, ref) -> c.value))
      .toMap

  private def cached(value: Option[CellValue]): Option[CellValue] =
    value.collect { case CellValue.Formula(_, cache, _) => cache }.flatten

  /** What one settle-edit-compare run found. */
  private final case class Run(
    after: Map[QualifiedRef, CellValue],
    differences: Vector[String],
    moved: Set[QualifiedRef]
  ):
    def cache(sheet: String, a1: String): Option[CellValue] = cached(after.get(q(sheet, a1)))

  /**
   * Settle `base`, apply `edits` to `on` (`None` clears the cell), then compare the after-edit pass
   * with a full recalculation of the edited book: every cell, and the cells reported as errors.
   * `moved` is the formulas of the cone whose cache differs from the settled book's.
   */
  private def run(base: Workbook, on: String, edits: (String, Option[CellValue])*): Run =
    val settled = base.recalculate(options()).workbook
    val edited = edits.foldLeft(settled)((wb, e) => set(wb, on, e._1, e._2))
    val refs = edits.map(e => r(e._1)).toSet
    val afterResult = edited.recalculateAfterEdit(sn(on), refs, options())
    val fullResult = edited.recalculate(options())
    val (after, full) = (cellsOf(afterResult.workbook), cellsOf(fullResult.workbook))
    val before = cellsOf(settled)
    val cone = DependentRecalculation.editCone(edited, refs.map(QualifiedRef(sn(on), _)))
    val keys = (after.keySet ++ full.keySet).toVector.sortBy(QualifiedRef.gridKey)
    val cellDifferences = keys.filter(k => after.get(k) != full.get(k)).map { k =>
      s"$k (in cone: ${cone.contains(k)}): after-edit ${after.get(k)}, full ${full.get(k)}"
    }
    def reported(result: RecalcResult) = result.errors.map(e => QualifiedRef(e.sheet, e.ref)).toSet
    val errorDifferences =
      if reported(afterResult) == reported(fullResult) then Vector.empty
      else Vector(s"errors: after-edit ${reported(afterResult)}, full ${reported(fullResult)}")
    val moved = cone.filter { k =>
      full.get(k).exists(_.isInstanceOf[CellValue.Formula]) &&
      cached(after.get(k)) != cached(before.get(k))
    }
    Run(after, cellDifferences ++ errorDifferences, moved)

  /** [[run]], asserting the law and that the edit moved a formula of its cone. */
  private def law(base: Workbook, on: String, edits: (String, Option[CellValue])*): Run =
    val result = run(base, on, edits*)
    val label = edits.map((a1, v) => s"$on!$a1 := $v").mkString(", ")
    assert(result.differences.isEmpty, result.differences.mkString(s"$label\n", "\n", ""))
    assert(result.moved.nonEmpty, s"$label moved no formula of its cone")
    result

  /** A 3x4 table: keys 1..3 in column A, then 10s, 100s and 1000s. */
  private val table: Seq[(String, CellValue)] = Seq(
    "A1" -> n(1),
    "B1" -> n(10),
    "C1" -> n(100),
    "D1" -> n(1000),
    "A2" -> n(2),
    "B2" -> n(20),
    "C2" -> n(200),
    "D2" -> n(2000),
    "A3" -> n(3),
    "B3" -> n(30),
    "C3" -> n(300),
    "D3" -> n(3000)
  )

  private def onTable(extra: (String, CellValue)*): Workbook =
    Workbook(sheet("S", (table ++ extra)*))

  // ---- lookups: the selector ------------------------------------------------------------------

  test("VLOOKUP whose column selector is a cell: editing the selector moves the strip it reads") {
    val wb = onTable("F1" -> n(3), "G1" -> f("VLOOKUP(2,A1:D3,F1,FALSE)"), "H1" -> f("G1+1"))
    val result = law(wb, "S", "F1" -> Some(n(4)))
    assertEquals(result.cache("S", "H1"), Some(n(2001)))
  }

  test("VLOOKUP by a selector cell: an edit upstream of a formula in the selected strip") {
    val wb = onTable(
      "D2" -> f("J1*2"),
      "J1" -> n(5),
      "F1" -> n(4),
      "G1" -> f("VLOOKUP(2,A1:D3,F1,FALSE)")
    )
    assertEquals(law(wb, "S", "J1" -> Some(n(9))).cache("S", "G1"), Some(n(18)))
  }

  test("VLOOKUP by a selector cell: the selector and the newly selected strip edited together") {
    val wb = onTable("F1" -> n(3), "G1" -> f("VLOOKUP(2,A1:D3,F1,FALSE)"))
    val result = law(wb, "S", "F1" -> Some(n(4)), "D2" -> Some(n(7)))
    assertEquals(result.cache("S", "G1"), Some(n(7)))
  }

  test("VLOOKUP with a fractional column index (literal and in a cell) reads the truncated strip") {
    val literal = onTable("G1" -> f("VLOOKUP(2,A1:D3,2.9,FALSE)"))
    assertEquals(law(literal, "S", "B2" -> Some(n(77))).cache("S", "G1"), Some(n(77)))
    val inCell = onTable("F1" -> n(2.9), "G1" -> f("VLOOKUP(2,A1:D3,F1,FALSE)"))
    assertEquals(law(inCell, "S", "B2" -> Some(n(77))).cache("S", "G1"), Some(n(77)))
  }

  test("VLOOKUP whose column index is +2, 1+1, a numeric string or TRUE+1 reads column 2") {
    val wb = onTable(
      "G1" -> f("VLOOKUP(2,A1:D3,+2,FALSE)"),
      "G2" -> f("VLOOKUP(2,A1:D3,1+1,FALSE)"),
      "G3" -> f("VLOOKUP(2,A1:D3,\"2\",FALSE)"),
      "G4" -> f("VLOOKUP(2,A1:D3,TRUE+1,FALSE)")
    )
    val result = law(wb, "S", "B2" -> Some(n(77)))
    Seq("G1", "G2", "G3", "G4").foreach(a1 => assertEquals(result.cache("S", a1), Some(n(77)), a1))
  }

  test("VLOOKUP whose selector is an IF: the branch and the strip it selects edited together") {
    val wb = onTable("K1" -> n(1), "G1" -> f("VLOOKUP(2,A1:D3,IF(K1=1,2,4),FALSE)"))
    val result = law(wb, "S", "D2" -> Some(n(9)), "K1" -> Some(n(0)))
    assertEquals(result.cache("S", "G1"), Some(n(9)))
  }

  // ---- lookups: the key and the table ----------------------------------------------------------

  test("approximate VLOOKUP: an edit in the key column changes the matched row") {
    val wb = onTable("G1" -> f("VLOOKUP(2.5,A1:D3,3)"), "G2" -> f("VLOOKUP(2.5,A1:D3,3,TRUE)"))
    val result = law(wb, "S", "A3" -> Some(n(2.4)))
    assertEquals(result.cache("S", "G1"), Some(n(300)))
    assertEquals(result.cache("S", "G2"), Some(n(300)))
  }

  test("VLOOKUP whose key is a formula: an edit moves the matched row") {
    val wb = onTable("K1" -> n(1), "G1" -> f("VLOOKUP(K1+1,A1:D3,3,FALSE)"))
    assertEquals(law(wb, "S", "K1" -> Some(n(2))).cache("S", "G1"), Some(n(300)))
  }

  test("VLOOKUP over whole columns: a key written far below the used range is found") {
    val wb = onTable("G1" -> f("VLOOKUP(9,A:D,3,FALSE)"), "C500" -> n(5))
    assertEquals(law(wb, "S", "A500" -> Some(n(9))).cache("S", "G1"), Some(n(5)))
  }

  test("VLOOKUP whose table and selector are defined names: result, key and selector edits") {
    val wb = named(
      onTable(
        "G1" -> f("VLOOKUP(2,Tbl,3,FALSE)"),
        "G2" -> f("VLOOKUP(2,Tbl,Sel,FALSE)"),
        "F1" -> n(3)
      ),
      DefinedName("Tbl", "S!$A$1:$D$3"),
      DefinedName("Sel", "S!$F$1")
    )
    assertEquals(law(wb, "S", "C2" -> Some(n(77))).cache("S", "G1"), Some(n(77)))
    law(wb, "S", "A2" -> Some(n(5)))
    assertEquals(law(wb, "S", "F1" -> Some(n(2))).cache("S", "G2"), Some(n(20)))
  }

  test("VLOOKUP into another sheet: an edit there, and one of the selector on the reader's") {
    val wb = Workbook(sheet("Data", table*)).put(
      sheet(
        "S",
        "F1" -> n(3),
        "G1" -> f("VLOOKUP(2,Data!A1:D3,F1,FALSE)"),
        "G2" -> f("VLOOKUP(2,Data!A1:D3,3,FALSE)")
      )
    )
    val edited = law(wb, "Data", "C2" -> Some(n(77)))
    assertEquals(edited.cache("S", "G1"), Some(n(77)))
    assertEquals(edited.cache("S", "G2"), Some(n(77)))
    assertEquals(law(wb, "S", "F1" -> Some(n(4))).cache("S", "G1"), Some(n(2000)))
  }

  test("HLOOKUP at a fixed row: an edit in the result row, and one that adds a second key match") {
    val wb = onTable("G1" -> f("HLOOKUP(10,A1:D3,3,FALSE)"), "G2" -> f("HLOOKUP(10,A1:D3,2,FALSE)"))
    assertEquals(law(wb, "S", "B3" -> Some(n(77))).cache("S", "G1"), Some(n(77)))
    // the key row gains an earlier match: A1 becomes 10, so the lookup moves from column B to A
    val firstMatch = onTable("G1" -> f("HLOOKUP(10,A1:D3,3,FALSE)"))
    assertEquals(law(firstMatch, "S", "A1" -> Some(n(10))).cache("S", "G1"), Some(n(3)))
  }

  test("VLOOKUP whose result strip holds a formula chain: an edit at the chain's far end") {
    val wb = onTable(
      "C2" -> f("J1+1"),
      "J1" -> f("K1*2"),
      "K1" -> n(1),
      "G1" -> f("VLOOKUP(2,A1:D3,3,FALSE)"),
      "H1" -> f("G1*2")
    )
    assertEquals(law(wb, "S", "K1" -> Some(n(4))).cache("S", "H1"), Some(n(18)))
  }

  // ---- INDEX ------------------------------------------------------------------------------------

  private val indexReaders: Seq[(String, CellValue)] = Seq(
    "G2" -> f("INDEX(A1:D3,ROW(),3)"),
    "G1" -> f("INDEX(A1:D3,2,2)"),
    "G3" -> f("SUM(INDEX(A1:D3,0,2))"),
    "H1" -> f("SUM(INDEX(A1:D3,2,0))"),
    "H2" -> f("INDEX(A1:D3,2,)"),
    "H3" -> f("SUM(INDEX(A1:D3,,2))")
  )

  test("INDEX at a literal, a ROW() and a zero or omitted position: an edit in a read column") {
    val result = law(onTable(indexReaders*), "S", "B3" -> Some(n(77)))
    assertEquals(result.cache("S", "G3"), Some(n(107)))
    assertEquals(result.cache("S", "H3"), Some(n(107)))
  }

  test("INDEX at a literal, a ROW() and a zero or omitted position: edits along the read row") {
    val result =
      law(onTable(indexReaders*), "S", "A2" -> Some(n(77)), "B2" -> Some(n(5)), "C2" -> Some(n(6)))
    assertEquals(result.cache("S", "G1"), Some(n(5)))
    assertEquals(result.cache("S", "G2"), Some(n(6)))
    assertEquals(result.cache("S", "H1"), Some(n(2088)))
  }

  test("INDEX as a range end, over one column or row, and over a multi-area reference") {
    val wb = onTable(
      "G1" -> f("SUM(INDEX(A1:D3,1,2):INDEX(A1:D3,3,3))"),
      "G2" -> f("INDEX(B1:B3,2)"),
      "G3" -> f("INDEX(A2:D2,3)"),
      "H1" -> f("INDEX((A1:B3,C1:D3),2,1,2)"),
      "H2" -> f("SUM(A1:INDEX(A1:D3,3,2))")
    )
    val result = law(wb, "S", "C2" -> Some(n(77)), "B3" -> Some(n(1)))
    assertEquals(result.cache("S", "G3"), Some(n(77)))
    assertEquals(result.cache("S", "H1"), Some(n(77)))
  }

  // ---- dynamic readers ---------------------------------------------------------------------------

  test("INDIRECT and OFFSET readers chained through static formulas") {
    val wb = Workbook(
      sheet(
        "S",
        "A1" -> f("INDIRECT(\"C1\")+1"),
        "B1" -> f("INDIRECT(\"D1\")"),
        "C1" -> f("B1+1"),
        "D1" -> n(5),
        "E1" -> f("A1*2"),
        "A2" -> f("OFFSET(A1,0,4)"),
        "F1" -> f("INDIRECT(\"A\"&D1)"),
        "A5" -> n(42),
        "A6" -> f("E1+1")
      )
    )
    val result = law(wb, "S", "D1" -> Some(n(6)))
    assertEquals(result.cache("S", "A1"), Some(n(8)))
    assertEquals(result.cache("S", "F1"), Some(n(17)))
  }

  test("an INDIRECT reader of a static formula: an edit upstream of the formula it names") {
    val wb =
      Workbook(sheet("S", "A1" -> f("INDIRECT(\"C\"&1)"), "C1" -> f("D1+1"), "D1" -> n(5)))
    assertEquals(law(wb, "S", "D1" -> Some(n(2))).cache("S", "A1"), Some(n(3)))
  }

  // ---- defined names -----------------------------------------------------------------------------

  private val namesBook: Workbook = named(
    Workbook(
      sheet(
        "S",
        "A1" -> n(1),
        "A2" -> n(2),
        "B1" -> f("Rate*2"),
        "B2" -> f("Total+1"),
        "B3" -> f("Alias"),
        "B4" -> f("SUM(Col)"),
        "B5" -> f("Loc")
      )
    ).put(sheet("T", "A1" -> n(7), "A2" -> n(3), "C1" -> f("S!B1+Loc"))),
    DefinedName("Rate", "T!$A$1"),
    DefinedName("Total", "SUM(T!$A$1:$A$2)+S!$A$1"),
    DefinedName("Alias", "Rate"),
    DefinedName("Col", "T!$A:$A"),
    DefinedName("Loc", "$A$2", localSheetId = Some(0)),
    DefinedName("Loc", "T!$A$2", localSheetId = Some(1))
  )

  test("names: a cell read through a name, an alias, a formula name and a whole-column name") {
    val result = law(namesBook, "T", "A1" -> Some(n(70)))
    assertEquals(result.cache("S", "B1"), Some(n(140)))
    assertEquals(result.cache("S", "B3"), Some(n(70)))
    assertEquals(result.cache("S", "B4"), Some(n(73)))
    assertEquals(result.cache("T", "C1"), Some(n(143)))
  }

  test("names: a sheet-scoped name resolves to its own sheet's cell on each sheet") {
    val onT = law(namesBook, "T", "A2" -> Some(n(30)))
    assertEquals(onT.cache("T", "C1"), Some(n(44)))
    assertEquals(onT.cache("S", "B5"), Some(n(2)))
    val onS = law(namesBook, "S", "A2" -> Some(n(20)))
    assertEquals(onS.cache("S", "B5"), Some(n(20)))
    assertEquals(onS.cache("T", "C1"), Some(n(17)))
  }

  test("names: a cell written below the used range of a whole-column name") {
    assertEquals(law(namesBook, "T", "A9" -> Some(n(1))).cache("S", "B4"), Some(n(11)))
  }

  // ---- whole columns and rows --------------------------------------------------------------------

  private val wholeAxisBook: Workbook = Workbook(
    sheet(
      "S",
      "A1" -> n(1),
      "F1" -> f("SUM(A:A)"),
      "F2" -> f("SUM(3:3)"),
      "F4" -> f("COUNTA(A:B)"),
      "F5" -> f("SUMIF(A:A,\">0\",B:B)"),
      "F6" -> f("MATCH(9,A:A,0)"),
      "F7" -> f("XLOOKUP(9,A:A,B:B)")
    )
  ).put(sheet("T", "A1" -> f("SUM(S!A:A)"), "A2" -> f("COUNT(S!1:1048576)")))

  test("whole-column readers, on the sheet and from another: cells written at row 5000") {
    val result = law(wholeAxisBook, "S", "A5000" -> Some(n(9)), "B5000" -> Some(n(4)))
    assertEquals(result.cache("S", "F1"), Some(n(10)))
    assertEquals(result.cache("S", "F6"), Some(n(5000)))
    assertEquals(result.cache("S", "F7"), Some(n(4)))
    assertEquals(result.cache("T", "A1"), Some(n(10)))
  }

  test("a whole-row reader: a cell written in a far column of the row") {
    assertEquals(law(wholeAxisBook, "S", "ZZ3" -> Some(n(9))).cache("S", "F2"), Some(n(9)))
  }

  test("whole-column readers: the only cell of the column removed") {
    val result = law(wholeAxisBook, "S", "A1" -> None)
    assertEquals(result.cache("S", "F1"), Some(n(0)))
    assertEquals(result.cache("S", "F4"), Some(n(0)))
    assertEquals(result.cache("T", "A1"), Some(n(0)))
  }

  // ---- cycles ------------------------------------------------------------------------------------

  test("an edit that creates a cycle withdraws its members and the cells they block") {
    val wb = Workbook(
      sheet(
        "S",
        "A1" -> f("B1+1"),
        "B1" -> n(1),
        "C1" -> f("A1+1"),
        "D1" -> f("C1+E1"),
        "E1" -> n(1),
        "G1" -> f("E1*2")
      )
    )
    val result = law(wb, "S", "B1" -> Some(f("C1+1")))
    Seq("A1", "B1", "C1", "D1").foreach(a1 => assertEquals(result.cache("S", a1), None, a1))
    assertEquals(result.cache("S", "G1"), Some(n(2)))
  }

  private val cycleBook: Workbook = Workbook(
    sheet(
      "S",
      "A1" -> f("B1+1"),
      "B1" -> f("C1+1"),
      "C1" -> f("A1+1"),
      "D1" -> f("C1+E1"),
      "E1" -> n(1),
      "G1" -> f("E1*2")
    )
  )

  test("an edit that removes a cycle computes its former members and what they blocked") {
    val result = law(cycleBook, "S", "B1" -> Some(n(3)))
    assertEquals(result.cache("S", "C1"), Some(n(5)))
    assertEquals(result.cache("S", "D1"), Some(n(6)))
  }

  test("an edit beside a standing cycle: its reader outside the cycle moves, the blocked stay") {
    val result = law(cycleBook, "S", "E1" -> Some(n(3)))
    assertEquals(result.cache("S", "G1"), Some(n(6)))
    assertEquals(result.cache("S", "D1"), None)
  }

  test("a formula that reads itself: an edit to what its reader also reads") {
    val wb = Workbook(
      sheet("S", "A1" -> f("A1+1"), "B1" -> f("A1+C1"), "C1" -> n(1), "D1" -> f("C1+1"))
    )
    val result = law(wb, "S", "C1" -> Some(n(2)))
    assertEquals(result.cache("S", "D1"), Some(n(3)))
    assertEquals(result.cache("S", "B1"), None)
  }

  test(
    "a cached cycle the edit does not reach keeps its caches; a cone cell behind it is blocked"
  ) {
    // a book another engine wrote: A1 and B1 are circular and hold caches; C1 reads A1. The edit
    // of D1 reaches E1 and F1 only. The pass leaves the cycle's caches and C1's alone (a full
    // recalculation withdraws them, so the every-cell law does not hold here, by design: cells
    // outside the cone are never touched). F1 is in the cone and behind the cycle: it is
    // withdrawn and reported as blocked, as a full recalculation has it
    def cachedFormula(text: String, value: Double) = CellValue.Formula(text, Some(n(value)))
    val book = Workbook(
      sheet(
        "S",
        "A1" -> cachedFormula("B1+1", 0),
        "B1" -> cachedFormula("A1+1", 0),
        "C1" -> cachedFormula("A1+5", 5),
        "D1" -> n(1),
        "E1" -> cachedFormula("D1+1", 2),
        "F1" -> cachedFormula("C1+D1", 6)
      )
    )
    val edited = set(book, "S", "D1", Some(n(10)))
    val result = edited.recalculateAfterEdit(sn("S"), Set(r("D1")), options())
    val after = cellsOf(result.workbook)
    assertEquals(after(q("S", "A1")), cachedFormula("B1+1", 0))
    assertEquals(after(q("S", "B1")), cachedFormula("A1+1", 0))
    assertEquals(after(q("S", "C1")), cachedFormula("A1+5", 5))
    assertEquals(cached(after.get(q("S", "E1"))), Some(n(11)))
    assertEquals(cached(after.get(q("S", "F1"))), None)
    assertEquals(
      result.errors.map(e => (e.ref.toA1, e.error.message.contains("Blocked"))),
      Vector("F1" -> true)
    )
    val full = edited.recalculate(options())
    assertEquals(cellsOf(full.workbook)(q("S", "F1")), after(q("S", "F1")))
    assertEquals(cellsOf(full.workbook)(q("S", "E1")), after(q("S", "E1")))
  }

  // ---- spills, formula edits, failures -----------------------------------------------------------

  test("a spill that shrinks: its `#` reader, a range over it and a cell it no longer covers") {
    val wb = Workbook(
      sheet(
        "S",
        "A1" -> n(3),
        "D1" -> f("SEQUENCE(A1)"),
        "F1" -> f("SUM(D1#)"),
        "F2" -> f("D3*2"),
        "F3" -> f("SUM(D1:D5)"),
        "F4" -> f("INDEX(D1#,2)")
      )
    )
    assertEquals(law(wb, "S", "A1" -> Some(n(2))).cache("S", "F1"), Some(n(3)))
  }

  test("a formula whose text is edited evaluates the new text, not a parse of the old one") {
    val wb = Workbook(sheet("S", "A1" -> n(1), "B1" -> f("A1+1"), "C1" -> f("B1*2")))
    assertEquals(law(wb, "S", "B1" -> Some(f("A1+100"))).cache("S", "C1"), Some(n(202)))
  }

  private val failingBook: Workbook = Workbook(
    sheet(
      "S",
      "A1" -> n(1),
      "B1" -> f("1/A1"),
      "C1" -> f("B1+1"),
      "D1" -> f("IFERROR(B1,9)"),
      "E1" -> f("SUM(B1:C1)")
    )
  )

  test("an edit that makes a precedent an error value: its readers, a guard and a range over it") {
    val result = law(failingBook, "S", "A1" -> Some(n(0)))
    assertEquals(result.cache("S", "D1"), Some(n(9)))
  }

  test("an edit that heals an error value") {
    val result = law(set(failingBook, "S", "A1", Some(n(0))), "S", "A1" -> Some(n(4)))
    assertEquals(result.cache("S", "C1"), Some(n(1.25)))
    assertEquals(result.cache("S", "D1"), Some(n(0.25)))
  }

  test("formulas that can never evaluate (unknown sheet, unknown function) beside an edit") {
    // B1 and F1 fail on every pass and hold no value, so every edit re-evaluates them and their
    // readers; the edit of A1 also moves D1 and E1
    val wb = Workbook(
      sheet(
        "S",
        "A1" -> n(1),
        "B1" -> f("Gone!A1+A1"),
        "C1" -> f("B1+1"),
        "D1" -> f("A1+1"),
        "E1" -> f("SUM(A1:D1)"),
        "F1" -> f("FOO(A1)"),
        "G1" -> f("F1+1")
      )
    )
    val result = law(wb, "S", "A1" -> Some(n(2)))
    assertEquals(result.cache("S", "D1"), Some(n(3)))
    Seq("B1", "C1", "F1", "G1").foreach(a1 => assertEquals(result.cache("S", a1), None, a1))
  }

  // ---- criteria and lazy arguments ---------------------------------------------------------------

  test("criteria functions whose sum range is resized to the criteria range") {
    val wb = onTable(
      "G1" -> f("SUMIF(A1:A3,\">1\",B1)"),
      "G2" -> f("SUMIFS(C1:C3,A1:A3,\">=2\",B1:B3,\"<30\")"),
      "G3" -> f("SUMPRODUCT(A1:A3,D1:D3)"),
      "G4" -> f("AVERAGEIF(A1:A2,\">0\",C1:C9)"),
      "G5" -> f("COUNTIFS(A:A,\">1\")")
    )
    // B3 lies outside SUMIF's written sum range B1 and inside the range it is resized to
    val result = law(wb, "S", "B3" -> Some(n(5)), "C3" -> Some(n(6)))
    assertEquals(result.cache("S", "G1"), Some(n(25)))
    assertEquals(result.cache("S", "G2"), Some(n(206)))
  }

  test("criteria functions: a criteria cell written below the table, a sum cell inside it") {
    val wb = onTable(
      "G1" -> f("SUMIF(A1:A3,\">1\",B1)"),
      "G4" -> f("AVERAGEIF(A1:A2,\">0\",C1:C9)"),
      "G5" -> f("COUNTIFS(A:A,\">1\")")
    )
    val result = law(wb, "S", "A9" -> Some(n(5)), "C2" -> Some(n(6)))
    assertEquals(result.cache("S", "G4"), Some(n(53)))
    assertEquals(result.cache("S", "G5"), Some(n(3)))
  }

  test("LET bindings, CHOOSE and IF branches, and a range intersection") {
    val wb = onTable(
      "G1" -> f("LET(x,B1:B3,SUM(x))"),
      "G2" -> f("CHOOSE(A1,B2,C2,D2)"),
      "G3" -> f("IF(A1=1,B3,C3)"),
      "G4" -> f("SUM(A1:B3 B2:D2)"),
      "G5" -> f("LET(t,A1:D3,VLOOKUP(2,t,3,FALSE))")
    )
    val result = law(wb, "S", "A1" -> Some(n(2)), "C3" -> Some(n(1)), "C2" -> Some(n(5)))
    assertEquals(result.cache("S", "G2"), Some(n(5)))
    assertEquals(result.cache("S", "G3"), Some(n(1)))
    // A1 became 2, so the lookup's first match is row 1
    assertEquals(result.cache("S", "G5"), Some(n(100)))
  }

  // ---- randomness --------------------------------------------------------------------------------

  test("seeded RAND after an edit: reproducible, seed-dependent, and untouched outside the cone") {
    val wb = Workbook(
      sheet(
        "S",
        "A1" -> n(1),
        "B1" -> f("A1+RAND()"),
        "B2" -> f("RAND()"),
        "B3" -> f("RANDBETWEEN(1,1000)+A1"),
        "C1" -> f("B1+RAND()")
      )
    )
    val settled = wb.recalculate(options()).workbook
    val edited = set(settled, "S", "A1", Some(n(2)))
    val first = edited.recalculateAfterEdit(sn("S"), Set(r("A1")), options())
    val second = edited.recalculateAfterEdit(sn("S"), Set(r("A1")), options())
    assertEquals(first.workbook, second.workbook)
    // B2 reads nothing the edit touched: it keeps the draw the settled book holds
    assertEquals(cellsOf(first.workbook)(q("S", "B2")), cellsOf(settled)(q("S", "B2")))
    val reseeded = edited.recalculateAfterEdit(sn("S"), Set(r("A1")), options(seed = 6L))
    assertNotEquals(cellsOf(first.workbook)(q("S", "B1")), cellsOf(reseeded.workbook)(q("S", "B1")))
  }
