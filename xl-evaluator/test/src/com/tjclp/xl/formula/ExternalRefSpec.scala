package com.tjclp.xl.formula

import com.tjclp.xl.{Anchor, CellRange}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.formula.eval.{SheetEvaluator, WorkbookEvaluator}
import munit.ScalaCheckSuite
import org.scalacheck.Prop.*
import org.scalacheck.Gen

/**
 * GH-353: external-workbook references ([2]Book1!A1, '[3]Sheet Name'!B2).
 *
 * Parser: the three surface forms parse to ExternalRef/ExternalRange instead of failing with
 * UnexpectedChar([. Printer: the canonical forms round-trip textually, and parse ∘ print = id on
 * the AST. Graph: external refs contribute no edges, so graph construction succeeds. Recalc: Excel
 * closed-workbook semantics — a formula cell touching an external workbook keeps its Excel-written
 * cached value verbatim (pinned, never re-evaluated); dependents compute from that cache; an
 * UNCACHED external cell is a clear per-cell error that propagates to dependents without aborting
 * the rest of the pass.
 */
class ExternalRefSpec extends ScalaCheckSuite:

  import SheetEvaluator.*
  import WorkbookEvaluator.*

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private def formula(expr: String): CellValue = CellValue.Formula(expr, None)
  private val a3 = ARef.from0(0, 2)
  private val d1 = ARef.from0(3, 0)
  private val d2 = ARef.from0(3, 1)
  private val e1 = ARef.from0(4, 0)

  private def cached(wb: Workbook, sheetName: String, ref: ARef): Option[CellValue] =
    wb.sheets
      .find(_.name.value == sheetName)
      .flatMap(_.cells.get(ref))
      .map(_.value)
      .collect { case CellValue.Formula(_, Some(v), _) => v }

  // ==================== Parser ====================

  test("parse: unquoted external cell ref [2]Book1!A1"):
    FormulaParser.parse("=[2]Book1!A1") match
      case Right(TExpr.ExternalRef(index, name, at, anchor)) =>
        assertEquals(index, 2)
        assertEquals(name, "Book1")
        assertEquals(at, ARef.from0(0, 0))
        assertEquals(anchor, Anchor.Relative)
      case other => fail(s"Expected ExternalRef, got $other")

  test("parse: unquoted external range [2]Book1!A1:B2"):
    FormulaParser.parse("=[2]Book1!A1:B2") match
      case Right(TExpr.ExternalRange(index, name, range)) =>
        assertEquals(index, 2)
        assertEquals(name, "Book1")
        assertEquals(range.toA1, "A1:B2")
      case other => fail(s"Expected ExternalRange, got $other")

  test("parse: quoted external ref '[3]Sheet Name'!B2 (bracket inside quotes)"):
    FormulaParser.parse("='[3]Sheet Name'!B2") match
      case Right(TExpr.ExternalRef(index, name, at, _)) =>
        assertEquals(index, 3)
        assertEquals(name, "Sheet Name")
        assertEquals(at, ARef.from0(1, 1))
      case other => fail(s"Expected ExternalRef, got $other")

  test("parse: external name with dots (the field repro [2]Consolidation.xlsx!D5:D9)"):
    FormulaParser.parse("=SUM([2]Consolidation.xlsx!D5:D9)") match
      case Right(expr) => assert(TExpr.containsExternalRef(expr))
      case Left(err) => fail(s"Expected parse success, got $err")

  test("parse: anchored external ref [2]Book1!$A$1 keeps the anchor"):
    FormulaParser.parse("=[2]Book1!$A$1") match
      case Right(TExpr.ExternalRef(_, _, _, anchor)) => assertEquals(anchor, Anchor.Absolute)
      case other => fail(s"Expected ExternalRef, got $other")

  test("parse: invalid workbook index is a clear error, not UnexpectedChar"):
    assert(FormulaParser.parse("=[0]Book1!A1").isLeft)
    assert(FormulaParser.parse("=[x]Book1!A1").isLeft)
    assert(FormulaParser.parse("=[2]!A1").isLeft)

  test("parse: external ranges are accepted in range-typed argument slots (GH-353 review)"):
    val forms = List(
      """=SUMIF([2]Book1!A1:A9, ">0")""",
      """=SUMIFS(D1:D9, [2]Book1!A1:A9, ">0")""",
      """=COUNTIF([2]Book1!A1:A9, ">0")""",
      """=AVERAGEIF([2]Book1!A1:A9, ">0")""",
      "=VLOOKUP(A1, [2]Book1!A1:B9, 2, FALSE)",
      "=SUMPRODUCT([2]Book1!A1:A9, B1:B9)"
    )
    forms.foreach { text =>
      FormulaParser.parse(text) match
        case Right(expr) => assert(TExpr.containsExternalRef(expr), s"not external: $text")
        case Left(err) => fail(s"parse failed for $text: $err")
    }

  test("parse: former LOCAL-literal-range slots now accept external ranges (GH-394)"):
    // GH-394 migrated the MATCH/INDEX/XLOOKUP/XIRR array slots from local-only CellRange to
    // RangeLocation, so external ranges PARSE like the SUMIF slots (GH-353): the parsed shape
    // carries External, evaluation is the friendly external-unsupported error, and cached
    // cells bearing these shapes pin via containsExternalRef (no longer via the parse-failure
    // fallback)
    List(
      "=XLOOKUP(1, [2]Book1!A1:A9, B1:B9)",
      "=MATCH(1, [2]Book1!A1:A9, 0)"
    ).foreach { text =>
      FormulaParser.parse(text) match
        case Right(expr) => assert(TExpr.containsExternalRef(expr), s"not external: $text")
        case Left(err) => fail(s"parse failed for $text: $err")
    }

  // ==================== Printer round-trips ====================

  test("print: canonical external forms round-trip the original text exactly"):
    val forms = List(
      "=[2]Book1!A1",
      "=[2]Book1!A1:B2",
      "='[3]Sheet Name'!B2",
      "=[2]Book1!$A$1",
      "=SUM([2]Consolidation.xlsx!D5:D9)",
      "=SUM([2]Book1!A1:A2)+1",
      """=SUMIF([2]Book1!A1:A9, ">0")"""
    )
    forms.foreach { text =>
      FormulaParser.parse(text) match
        case Right(expr) => assertEquals(FormulaPrinter.print(expr), text)
        case Left(err) => fail(s"parse failed for $text: $err")
    }

  private val genExternalName: Gen[String] =
    Gen.oneOf(
      Gen.identifier.map(_.take(20)).suchThat(_.nonEmpty),
      Gen
        .zip(Gen.identifier.suchThat(_.nonEmpty), Gen.identifier.suchThat(_.nonEmpty))
        .map((a, b) => s"${a.take(10)} ${b.take(10)}") // forces the quoted form
    )

  private val genExternalRef: Gen[TExpr[?]] =
    for
      index <- Gen.choose(1, 99)
      name <- genExternalName
      col <- Gen.choose(0, 100)
      row <- Gen.choose(0, 100)
    yield TExpr.ExternalRef(index, name, ARef.from0(col, row))

  private val genExternalRange: Gen[TExpr[?]] =
    for
      index <- Gen.choose(1, 99)
      name <- genExternalName
      col <- Gen.choose(0, 50)
      row <- Gen.choose(0, 50)
      w <- Gen.choose(0, 5)
      h <- Gen.choose(0, 5)
    yield TExpr.ExternalRange(
      index,
      name,
      CellRange(ARef.from0(col, row), ARef.from0(col + w, row + h))
    )

  property("round-trip: parse ∘ print = id for external refs"):
    forAll(genExternalRef) { expr =>
      val printed = FormulaPrinter.print(expr)
      FormulaParser.parse(printed) == Right(expr)
    }

  property("round-trip: parse ∘ print = id for external ranges"):
    forAll(genExternalRange) { expr =>
      val printed = FormulaPrinter.print(expr)
      FormulaParser.parse(printed) == Right(expr)
    }

  // ==================== Shifter ====================

  test("shift: external refs drag anchor-aware like sheet-qualified refs"):
    val shifted = FormulaParser
      .parse("=[2]Book1!A1+[2]Book1!$A$1")
      .map(FormulaShifter.shift(_, 1, 1))
      .map(FormulaPrinter.print(_))
    assertEquals(shifted, Right("=[2]Book1!B2+[2]Book1!$A$1"))

  // ==================== Dependency graph ====================

  test("graph: construction succeeds and external refs contribute no edges"):
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(d1, formula("=SUM([2]Book1!A1:A2)"))
      .put(e1, formula("=SUM(D1)+1"))
    val graph = DependencyGraph.fromSheet(sheet)
    assertEquals(DependencyGraph.precedents(graph, d1), Set.empty[ARef])
    assertEquals(DependencyGraph.precedents(graph, e1), Set(d1))

    val (deps, _) = DependencyGraph.fromWorkbookBounded(Workbook(sheet))
    val qd1 = DependencyGraph.QualifiedRef(SheetName.unsafe("S"), d1)
    assertEquals(deps.get(qd1), Some(Set.empty[DependencyGraph.QualifiedRef]))

  // ==================== Recalculate: cache pinning ====================

  test("recalculate: external formula keeps its Excel cache verbatim, dependent computes from it"):
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(d1, CellValue.Formula("=SUM([2]Book1!A1:A2)", Some(num(42))))
      .put(e1, formula("=SUM(D1)+1"))
    val result = Workbook(sheet).recalculate()
    assert(result.isClean, s"expected clean recalc, got ${result.errors.map(_.render)}")
    assertEquals(cached(result.workbook, "S", d1), Some(num(42)), "cache must be preserved")
    assertEquals(cached(result.workbook, "S", e1), Some(num(43)))

  test("recalculate: cached external SUMIF/SUMIFS pin verbatim and stay clean (GH-353 review)"):
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(d1, CellValue.Formula("""=SUMIF([2]Book1!A1:A9, ">0")""", Some(num(42))))
      .put(d2, CellValue.Formula("""=SUMIFS(A1:A9, [2]Book1!A1:A9, ">0")""", Some(num(7))))
      .put(e1, formula("=D1+D2"))
    val result = Workbook(sheet).recalculate()
    assert(result.isClean, s"expected clean recalc, got ${result.errors.map(_.render)}")
    assertEquals(cached(result.workbook, "S", d1), Some(num(42)))
    assertEquals(cached(result.workbook, "S", d2), Some(num(7)))
    assertEquals(cached(result.workbook, "S", e1), Some(num(49)))

  test("recalculate: UNCACHED external SUMIF is the friendly per-cell error (GH-353 review)"):
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(d1, formula("""=SUMIF([2]Book1!A1:A9, ">0")"""))
    val result = Workbook(sheet).recalculate()
    assertEquals(result.errors.map(_.ref), Vector(d1))
    assert(
      result.errors.forall(_.error.message.contains("External workbook reference")),
      result.errors.map(_.render).mkString("; ")
    )

  test("recalculate: MIXED local+external cached formula stays pinned when the local part changes"):
    // The pin is verbatim: xl does NOT recompute from the external cache parts the way Excel
    // does, so the cache is served even though A1 changed (documented in LIMITATIONS.md GH-353)
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(ARef.from0(0, 0), num(5))
      .put(d1, CellValue.Formula("=A1+[2]Book1!B1", Some(num(42))))
      .put(e1, formula("=D1+1"))
    val result = Workbook(sheet).recalculate()
    assert(result.isClean, s"expected clean recalc, got ${result.errors.map(_.render)}")
    assertEquals(cached(result.workbook, "S", d1), Some(num(42)), "pin must survive verbatim")
    assertEquals(cached(result.workbook, "S", e1), Some(num(43)))

  test("recalculate: cached external cell whose shape is beyond the parser still pins (fallback)"):
    // XLOOKUP's array slots require LOCAL literal ranges, so this formula fails to parse —
    // the lexical '[n]' fallback in pinnedExternalCache must keep the Excel cache anyway
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(d1, CellValue.Formula("=XLOOKUP(1, [2]Book1!A1:A9, [2]Book1!B1:B9)", Some(num(42))))
      .put(e1, formula("=D1+1"))
    val result = Workbook(sheet).recalculate()
    assert(result.isClean, s"expected clean recalc, got ${result.errors.map(_.render)}")
    assertEquals(cached(result.workbook, "S", d1), Some(num(42)))
    assertEquals(cached(result.workbook, "S", e1), Some(num(43)))

  test("recalculate: UNCACHED external formula errors per-cell and propagates to dependents"):
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(d1, formula("=SUM([2]Book1!A1:A2)"))
      .put(e1, formula("=SUM(D1)+1"))
      .put(a3, formula("=1+1")) // unrelated cell still computes
    val result = Workbook(sheet).recalculate()
    assertEquals(cached(result.workbook, "S", a3), Some(num(2)), "unrelated cell must compute")
    assertEquals(cached(result.workbook, "S", d1), None)
    assertEquals(cached(result.workbook, "S", e1), None)
    assertEquals(result.errors.map(_.ref).toSet, Set(d1, e1))
    assert(
      result.errors.forall(_.error.message.contains("External workbook reference")),
      result.errors.map(_.render).mkString("; ")
    )

  // ==================== Direct evaluation ====================

  test("evaluateFormula: direct eval of an external formula is a clear error"):
    val sheet = Sheet(SheetName.unsafe("S"))
    Workbook(sheet).evaluateFormula("=SUM([2]Book1!A1:A2)", SheetName.unsafe("S")) match
      case Left(err) =>
        assert(err.message.contains("External workbook reference"), err.message)
        assert(!err.message.contains("UnexpectedChar"), err.message)
      case Right(v) => fail(s"expected error, got $v")

  test("evaluateFormula: standalone external ref is a clear error"):
    val sheet = Sheet(SheetName.unsafe("S"))
    sheet.evaluateFormula("=[2]Book1!A1") match
      case Left(err) => assert(err.message.contains("External workbook reference"), err.message)
      case Right(v) => fail(s"expected error, got $v")

  test("evaluateFormula: the issue repro — SUM over a cached external precedent uses the cache"):
    val sheet = Sheet(SheetName.unsafe("Data"))
      .put(d1, CellValue.Formula("=SUM([2]Book1!A1:A2)", Some(num(42))))
      .put(d2, num(8))
    val wb = Workbook(sheet)
    assertEquals(
      wb.evaluateFormula("=SUM(D1:D2)", SheetName.unsafe("Data")),
      Right(num(50))
    )

  test("evaluateFormula: SUM over an UNCACHED external precedent errors that cell's chain"):
    val sheet = Sheet(SheetName.unsafe("Data"))
      .put(d1, formula("=SUM([2]Book1!A1:A2)"))
      .put(d2, num(8))
    val wb = Workbook(sheet)
    wb.evaluateFormula("=SUM(D1:D2)", SheetName.unsafe("Data")) match
      case Left(err) => assert(err.message.contains("External workbook reference"), err.message)
      case Right(v) => fail(s"expected error, got $v")

  test("evaluateCell: cached external cell is pinned to its cache"):
    val sheet = Sheet(SheetName.unsafe("S"))
      .put(d1, CellValue.Formula("=SUM([2]Book1!A1:A2)", Some(num(42))))
    assertEquals(sheet.evaluateCell(d1), Right(num(42)))
