package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.{RecalcCounters, RecalcEngine, RecalcOptions, WorkbookEvaluator}
import com.tjclp.xl.formula.functions.{FunctionSpec, ReferenceOperators}
import com.tjclp.xl.formula.functions.ReferenceOperators.{Envelope, Operand}
import com.tjclp.xl.formula.graph.CompactGraph
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.*

/**
 * #743: INDEX folds a position written with ROW() or COLUMN() relative to the formula that reads
 * it, so `B5 = SUM(B1:INDEX(B:B,ROW()-1))` depends on B1:B4 and not on its whole array B:B (which
 * holds B5). Every dependency extractor sees the same narrowing: the sheet graph, the workbook
 * graphs, the edit-cone index and the compact graph. #745: the narrowing is INDEX's
 * [[FunctionSpec.envelope]] hook, which any reference-computing spec can declare.
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
class IndexReaderEnvelopeSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(200)

  private val name = SheetName.unsafe("S")
  private def N(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def F(text: String): CellValue = CellValue.Formula(text, None)
  private def at(a1: String): ARef = ARef.parse(a1).fold(msg => fail(msg), identity)
  private def q(a1: String): QualifiedRef = QualifiedRef(name, at(a1))

  private def sheetOf(cells: (String, CellValue)*): Sheet =
    cells.foldLeft(Sheet(name))((s, c) => s.put(at(c._1), c._2))

  /** The formula cells every graph finds cyclic, the same in each (a mismatch fails here). */
  private def cyclic(sheet: Sheet): Set[QualifiedRef] =
    val wb = Workbook(Vector(sheet))
    val bySheet =
      DependencyGraph.detectCycles(DependencyGraph.fromSheet(sheet)).fold(_ => true, _ => false)
    val (bounded, _) = DependencyGraph.fromWorkbookBounded(wb)
    val (formulaGraph, _) = DependencyGraph.fromWorkbookFormulaGraph(wb)
    val legacy = DependencyGraph.qualifiedCyclicNodes(formulaGraph)
    val compactGraph = CompactGraph.build(wb)
    val compact = compactGraph.cyclicFormulas.iterator.map(compactGraph.nodes.qualified).toSet
    assertEquals(compact, legacy, "compact vs legacy")
    assertEquals(DependencyGraph.qualifiedCyclicNodes(bounded), legacy, "bounded vs legacy")
    assertEquals(bySheet, legacy.nonEmpty, "sheet graph vs legacy")
    legacy

  private def recalc(sheet: Sheet, engine: RecalcEngine): Map[ARef, CellValue] =
    val result = WorkbookEvaluator.recalculateCounted(
      Workbook(Vector(sheet)),
      RecalcOptions(),
      RecalcCounters.disabled,
      engine
    )
    assertEquals(result.errors, Vector.empty, s"$engine errors")
    result.evaluated.getOrElse(name, Map.empty)

  private val issue: Sheet =
    sheetOf(
      "B1" -> N(1),
      "B2" -> N(2),
      "B3" -> N(3),
      "B4" -> N(4),
      "B5" -> F("SUM(B1:INDEX(B:B,ROW()-1))")
    )

  test("#743: B5 = SUM(B1:INDEX(B:B,ROW()-1)) is acyclic in every graph and sums B1:B4") {
    assertEquals(cyclic(issue), Set.empty[QualifiedRef])
    RecalcEngine.values.foreach { engine =>
      assertEquals(recalc(issue, engine).get(at("B5")), Some(N(10)), engine.toString)
    }
  }

  test("#743: the edges are B1:B4, in the workbook graph and the edit-cone index") {
    val wb = Workbook(Vector(issue))
    val (bounded, _) = DependencyGraph.fromWorkbookBounded(wb)
    assertEquals(bounded(q("B5")), Set("B1", "B2", "B3", "B4").map(q))
    // an edit to B4 reaches B5 through its edges; one to B6 (outside its read) does not
    val index = DependencyGraph.fromWorkbookDependencyIndex(wb)
    assertEquals(index.transitiveDependents(Set(q("B4"))), Set(q("B5")))
    assertEquals(index.transitiveDependents(Set(q("B6"))), Set.empty[QualifiedRef])
  }

  test("#743: a true self-read SUM(B1:INDEX(B:B,ROW())) is still circular") {
    val own = issue.put(at("B5"), F("SUM(B1:INDEX(B:B,ROW()))"))
    assertEquals(cyclic(own), Set(q("B5")))
    // ROW()-1 in row 1 is INDEX's 0, the whole column: circular as Excel has it
    val top = sheetOf("B1" -> F("SUM(B2:INDEX(B:B,ROW()-1))"), "B2" -> N(1))
    assertEquals(cyclic(top), Set(q("B1")))
  }

  test("#743: COLUMN() folds like ROW(): a running total along a row") {
    val row = sheetOf(
      "A3" -> N(5),
      "B3" -> N(6),
      "C3" -> F("SUM(A3:INDEX(3:3,COLUMN()-1))"),
      "D3" -> F("INDEX($A$3:$Z$3,1,COLUMN()-1)*2")
    )
    assertEquals(cyclic(row), Set.empty[QualifiedRef])
    RecalcEngine.values.foreach { engine =>
      val values = recalc(row, engine)
      assertEquals(values.get(at("C3")), Some(N(11)), engine.toString)
      assertEquals(values.get(at("D3")), Some(N(22)), engine.toString)
    }
  }

  test("#743: a 1,000-row running-total column computes with no cycle") {
    val rows = 1000
    // A: 1..n; B: the running total through INDEX(B:B,ROW()-1); C: the row number through MAX
    // over the computed range C1:C{n-1}
    val sheet = (2 to rows).foldLeft(sheetOf("A1" -> N(1), "B1" -> F("A1"), "C1" -> N(1))) {
      (s, r) =>
        s.put(ARef.from1(1, r), N(r))
          .put(ARef.from1(2, r), F(s"A$r+INDEX(B:B,ROW()-1)"))
          .put(ARef.from1(3, r), F("MAX(C$1:INDEX(C:C,ROW()-1))+1"))
    }
    assertEquals(cyclic(sheet), Set.empty[QualifiedRef])
    RecalcEngine.values.foreach { engine =>
      val values = recalc(sheet, engine)
      assertEquals(values.get(ARef.from1(2, rows)), Some(N(rows * (rows + 1) / 2)), s"$engine B")
      assertEquals(values.get(ARef.from1(3, rows)), Some(N(rows)), s"$engine C")
      assertEquals(values.get(ARef.from1(2, 500)), Some(N(500 * 501 / 2)), s"$engine B500")
    }
  }

  test("#743 limit: INDEX under a `:` operand not fixed for the reader, or in a name, keeps B:B") {
    // Excel gives 10 for both; xl still reports the whole array, so both are (false) cycles. A
    // `:` whose operand is CHOOSE is classified without a reader, and a name body is expanded once
    // for every reader, so neither folds ROW(). LIMITATIONS.md states both.
    def legacyCyclic(wb: Workbook): Set[QualifiedRef] =
      val (formulaGraph, _) = DependencyGraph.fromWorkbookFormulaGraph(wb)
      val legacy = DependencyGraph.qualifiedCyclicNodes(formulaGraph)
      val compactGraph = CompactGraph.build(wb)
      val compact = compactGraph.cyclicFormulas.iterator.map(compactGraph.nodes.qualified).toSet
      assertEquals(compact, legacy, "compact vs legacy")
      legacy
    val choose = issue.put(at("B5"), F("SUM(B1:CHOOSE(1,INDEX(B:B,ROW()-1)))"))
    assertEquals(legacyCyclic(Workbook(Vector(choose))), Set(q("B5")))
    val named = Workbook(Vector(issue.put(at("B5"), F("SUM(B1:prev)"))))
      .withDefinedName("prev", "INDEX(S!$B:$B,ROW()-1)")
    assertEquals(legacyCyclic(named), Set(q("B5")))
  }

  // ===== laws =====

  /** A1:F10 of numbers, so every INDEX the laws write selects a value. */
  private val grid: Sheet =
    (for c <- 0 until 6; r <- 0 until 10 yield ARef.from0(c, r)).foldLeft(Sheet(name)) {
      (acc, ref) => acc.put(ref, N(ref.col.index0 * 10 + ref.row.index0 + 1))
    }

  /** The formula cell: column H or below the grid, so its own row and column vary. */
  private val genHost: Gen[ARef] =
    Gen.oneOf(
      Gen.choose(0, 14).map(r => ARef.from0(7, r)),
      Gen.choose(0, 7).map(c => ARef.from0(c, 12))
    )

  private def offset(base: String, d: Int): String =
    if d == 0 then base else if d > 0 then s"$base+$d" else s"$base-${-d}"

  /** A position INDEX folds for its reader, or one only evaluation knows (`$H$20`, blank: 0). */
  private def genPosition(host: ARef, target: Int, alongRows: Boolean): Gen[String] =
    val own = if alongRows then host.row.index0 + 1 else host.col.index0 + 1
    val fn = if alongRows then "ROW()" else "COLUMN()"
    Gen.frequency(
      4 -> Gen.const(offset(fn, target - own)),
      2 -> Gen.choose(-2, 2).map(d => offset(fn, target - own + d)),
      1 -> Gen.const(s"2*$fn-${2 * own - target}"),
      1 -> Gen.const(target.toString),
      1 -> Gen.const("0"),
      1 -> Gen.const("$H$20")
    )

  private val genArray: Gen[CellRange] =
    Gen.frequency(
      4 -> (for
        c1 <- Gen.choose(0, 5)
        c2 <- Gen.choose(c1, 5)
        r1 <- Gen.choose(0, 9)
        r2 <- Gen.choose(r1, 9)
      yield CellRange(ARef.from0(c1, r1), ARef.from0(c2, r2))),
      1 -> Gen.choose(0, 5).map(c => CellRange(ARef.from0(c, 0), ARef.from0(c, Row.MaxIndex0))),
      1 -> Gen.choose(0, 9).map(r => CellRange(ARef.from0(0, r), ARef.from0(Column.MaxIndex0, r)))
    )

  /** An INDEX call at `host`: two- or three-argument, its positions aimed inside its array. */
  private val genIndex: Gen[(ARef, CellRange, String)] =
    for
      host <- genHost
      array <- genArray
      row <- Gen.choose(1, math.min(array.height, 12))
      col <- Gen.choose(1, math.min(array.width, 8))
      rowText <- genPosition(host, array.rowStart.index0 + row, alongRows = true)
      colText <- genPosition(host, array.colStart.index0 + col, alongRows = false)
      relRow = rowText.replace("ROW()", s"(ROW()-${array.rowStart.index0})")
      relCol = colText.replace("COLUMN()", s"(COLUMN()-${array.colStart.index0})")
      three <- Gen.oneOf(true, false)
    yield
      val args = if three then s"$relRow,$relCol" else relRow
      (host, array, s"INDEX(${array.toA1},$args)")

  private def plain(host: ARef, formula: String): Option[BigDecimal] =
    val placed = grid.put(host, F(formula))
    placed.evaluateCell(host, Clock.system, Some(Workbook(Vector(placed)))).toOption.collect {
      case CellValue.Number(n) => n
    }

  /** The area an envelope bounds: its one leaf or its range. */
  private def area(e: Envelope): Option[CellRange] = e match
    case Envelope.Exact(Vector((None, range)), _) => Some(range)
    case Envelope.Within(None, range, _) => Some(range)
    case _ => None

  private def holds(outer: CellRange, inner: CellRange): Boolean =
    outer.contains(inner.start) && outer.contains(inner.end)

  property("#745 law: INDEX's declared envelope holds every cell it reads, at every reader") {
    forAllNoShrink(genIndex) { (host, array, index) =>
      val expr = parsed(s"=$index")
      val envelope = ReferenceOperators.envelope(
        Right(expr.asInstanceOf[TExpr[Any]]),
        ReferenceOperators.staticShape,
        Some(host)
      )
      // the reference evaluation selects, as its corner and extent (none when it is an error)
      val selected =
        for
          row <- plain(host, s"ROW($index)")
          col <- plain(host, s"COLUMN($index)")
          rows <- plain(host, s"ROWS($index)")
          cols <- plain(host, s"COLUMNS($index)")
        yield CellRange(
          ARef.from1(col.toInt, row.toInt),
          ARef.from1(col.toInt + cols.toInt - 1, row.toInt + rows.toInt - 1)
        )
      selected.foreach { cells =>
        assert(area(envelope).exists(holds(_, cells)), s"$index at $host: $envelope vs $cells")
        // an exact envelope is exactly the selection
        envelope match
          case Envelope.Exact(_, _) => assertEquals(area(envelope), Some(cells), s"$index at $host")
          case _ => ()
      }
      // and never past the array
      assert(area(envelope).forall(holds(array, _)), s"$index at $host: $envelope")
      true
    }
  }

  /** `SUM(a:b)` with each operand a cell INDEX picks from A1:F10 by ROW() and COLUMN(). */
  private val genRange: Gen[(ARef, CellRange, String)] =
    def pick(host: ARef, cell: ARef): Gen[String] =
      val r = offset("ROW()", cell.row.index0 - host.row.index0)
      val c = offset("COLUMN()", cell.col.index0 - host.col.index0)
      Gen.oneOf(cell.toA1, s"INDEX($$A$$1:$$F$$10,$r,$c)", s"INDEX($$A:$$F,$r,$c)")
    val genCell = Gen.zip(Gen.choose(0, 5), Gen.choose(0, 9)).map((c, r) => ARef.from0(c, r))
    for
      host <- genHost
      a <- genCell
      b <- genCell
      sa <- pick(host, a)
      // XLOOKUP's envelope is b, covered by its own edges: the range is static only through
      // INDEX's array, whose edges must stay whole however INDEX folds (operandReader)
      looked = s"XLOOKUP(${b.col.index0 * 10 + b.row.index0 + 1},${b.toA1},${b.toA1})"
      sb <- Gen.frequency(3 -> pick(host, b), 1 -> Gen.const(looked))
      other <- Gen.oneOf("", ",IF(TRUE,$B$2:$B$9,0)", "+COUNT(INDEX($C:$C,ROW()))")
    yield (host, CellRange(a, a).expand(b), s"SUM($sa:$sb$other)")

  property("#743 law: edges at the reader hold every cell the range reads, in every extractor") {
    forAllNoShrink(genRange) { (host, box, formula) =>
      val expr = parsed(s"=$formula")
      val read = box.cells.toSet
      val dynamic = DependencyGraph.containsDynamicReference(expr)
      val placed = grid.put(host, F(formula))
      val wb = Workbook(Vector(placed))
      val sheetEdges = DependencyGraph.fromSheet(placed).dependencies.getOrElse(host, Set.empty)
      val (bounded, _) = DependencyGraph.fromWorkbookBounded(wb)
      val qualified = bounded.getOrElse(QualifiedRef(name, host), Set.empty).map(_.ref)
      assert(dynamic || read.subsetOf(sheetEdges), s"$formula at $host: sheet $sheetEdges")
      assert(dynamic || read.subsetOf(qualified), s"$formula at $host: workbook $qualified")
      // the value reads the CLI's cycle check takes hold them too (a dynamic range's are found at
      // run time)
      val reads =
        graph.ValueDependencies.Reader(placed.usedRange, Some(name)).reads(expr, Some(host))
      assert(dynamic || read.subsetOf(reads), s"$formula at $host: value reads $reads")
      // a range of cells INDEX picks by position is static, with no edge outside what it reads
      if !formula.contains(",IF(") && !formula.contains("+") && !formula.contains("XLOOKUP") then
        assert(!dynamic, formula)
        assertEquals(sheetEdges, read, s"$formula at $host")
      true
    }
  }

  private def parsed(formula: String): TExpr[?] =
    FormulaParser.parse(formula).fold(err => fail(s"$formula: $err"), identity)

  // ===== #745: the hook =====

  /** A test reference function PICK(ref) whose result is some reference inside `ref`. */
  private def pick(hook: Option[FunctionSpec.EnvelopeFn[Operand]]) =
    FunctionSpec.referencing[eval.ArrayResult, Operand]("PICK", Arity.one, envelope = hook)(
      (_, _) => Left(EvalError.EvalFailed("PICK is a test spec"))
    )((_, _) => Left(EvalError.EvalFailed("PICK is a test spec")))(using
      ReferenceOperators.referenceOperand
    )

  private val inside: FunctionSpec.EnvelopeFn[Operand] = (ref, of, _) =>
    of(ref) match
      case Envelope.Exact(Vector((sheet, range)), covered) => Envelope.Within(sheet, range, covered)
      case other => other

  test("#745: a referencing spec that declares an envelope hook gets static edges") {
    val a1a9 = TExpr.RangeLocation.Local(CellRange(at("A1"), at("A9")))
    def range(spec: FunctionSpec[eval.ArrayResult] { type Args = Operand }): TExpr.Call[?] =
      val call = TExpr.Call(spec, Left(a1a9))
      TExpr.Call(
        ReferenceOperators.range,
        (
          Left(TExpr.RangeLocation.Local(CellRange(at("A1"), at("A1")))),
          Right(call.asInstanceOf[TExpr[Any]])
        )
      )
    val hooked = range(pick(Some(inside)))
    val bare = range(pick(None))
    // without a hook the graph cannot bound PICK: the range is dynamic, as before #745
    assert(ReferenceOperators.isComputedRangeDynamic(bare, ReferenceOperators.staticShape))
    assert(DependencyGraph.containsDynamicReference(bare))
    // with one, PICK's own edges (A1:A9) cover A1:PICK(A1:A9): static, its edges those cells
    assert(!ReferenceOperators.isComputedRangeDynamic(hooked, ReferenceOperators.staticShape))
    assert(!DependencyGraph.containsDynamicReference(hooked))
    assertEquals(
      DependencyGraph.extractDependencies(hooked),
      CellRange(at("A1"), at("A9")).cells.toSet
    )
  }
