package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.{IterativeCalc, IterativeMode, RecalcOptions, RecalcResult}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.testing.SmallStack
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * #829 (and #828): a cycle that closes through a reference computed at run time — INDIRECT, OFFSET,
 * INDEX at ROW(), a range reaching an uncached cell, a spill — is invisible to the dependency
 * graph. It used to run until either the recursion-depth guard or the thread's stack gave out,
 * whichever came first on the JIT tier at hand, so the same workbook recalculated twice could fail
 * with different errors. It is now the circular reference at the first re-entry into a cell whose
 * evaluation is in progress: the same result on every run, on any stack, warm or cold.
 */
class DynamicCycleSpec extends FunSuite:

  private val S = SheetName.unsafe("S")
  private val Prefix = "Circular reference through "
  private val Guards = List("exhausted the stack", "recursion depth exceeded")

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private def formula(text: String): CellValue = CellValue.Formula(text, None)
  private val options = RecalcOptions(
    clock = Clock.fixedDate(java.time.LocalDate.of(2026, 1, 2)),
    iterative = IterativeMode.Off
  )

  private def recalc(sheet: Sheet, opts: RecalcOptions = options): RecalcResult =
    Workbook(sheet).recalculate(opts)

  /** A cell's cached value (a formula's cache, a constant's value), None when uncached. */
  private def valueAt(result: RecalcResult, ref: ARef): Option[CellValue] =
    result.workbook.sheets.find(_.name == S).flatMap(_.cells.get(ref)).map(_.value).flatMap {
      case CellValue.Formula(_, cached, _) => cached
      case other => Some(other)
    }

  private def messages(result: RecalcResult): Map[ARef, String] =
    result.errors.map(e => e.ref -> e.error.message).toMap

  private def through(ref: ARef): String =
    s"${Prefix}S!${ref.toA1}: the cell is read again while its own evaluation is in progress"

  /** The same result on the runner thread, a 1 MB thread and a 512 KB thread, with no guard. */
  private def assertStackIndependent(sheet: Sheet, opts: RecalcOptions = options): RecalcResult =
    val onRunner = recalc(sheet, opts)
    val onWide = SmallStack.run(1024L * 1024L, "dynamic-cycle-1mb")(recalc(sheet, opts))
    val onNarrow = SmallStack.run(512L * 1024L, "dynamic-cycle-512kb")(recalc(sheet, opts))
    assertEquals(onWide, onRunner, "a 1 MB stack")
    assertEquals(onNarrow, onRunner, "a 512 KB stack")
    assertEquals(recalc(sheet, opts), onRunner, "a second run on the runner thread")
    onRunner.errors.foreach { e =>
      assert(!Guards.exists(e.error.message.contains), s"a guard answered: ${e.render}")
    }
    onRunner

  test("#829: INDIRECT closing on its own cell is the circular reference through that cell") {
    val sheet = Sheet(S).put(ref"A1", formula("=INDIRECT(\"A1\")")).put(ref"B1", num(5))
    val result = assertStackIndependent(sheet)
    assertEquals(
      messages(result),
      Map(ref"A1" -> s"Formula error in '=INDIRECT(\"A1\")': ${through(ref"A1")}")
    )
    assertEquals(valueAt(result, ref"A1"), None)
    assertEquals(valueAt(result, ref"B1"), Some(num(5)))
  }

  test(
    "#829: two cells reading each other through INDIRECT each report the cycle from their own root"
  ) {
    val sheet = Sheet(S)
      .put(ref"A1", formula("=INDIRECT(\"B1\")"))
      .put(ref"B1", formula("=INDIRECT(\"A1\")"))
      .put(ref"X1", num(5))
      .put(ref"Y1", formula("=X1*2"))
    val result = assertStackIndependent(sheet)
    assertEquals(messages(result).keySet, Set(ref"A1", ref"B1"))
    // each member is evaluated from its own root and meets itself first
    assert(messages(result)(ref"A1").endsWith(through(ref"A1")), messages(result)(ref"A1"))
    assert(messages(result)(ref"B1").endsWith(through(ref"B1")), messages(result)(ref"B1"))
    assertEquals(valueAt(result, ref"A1"), None)
    assertEquals(valueAt(result, ref"B1"), None)
    assertEquals(valueAt(result, ref"Y1"), Some(num(10)))
  }

  test("#829: OFFSET cycles, and a range read that reaches the reading cell") {
    // OFFSET's base is a static edge (a cycle through bases is Tarjan's); the cycle here closes
    // only through the offsets: A1 reads B1 and B1 reads A1, both from the constant C1
    val sheet = Sheet(S)
      .put(ref"A1", formula("=OFFSET(C1,0,-1)"))
      .put(ref"B1", formula("=OFFSET(C1,0,-2)"))
      .put(ref"C1", num(7))
      .put(ref"D1", num(1))
      .put(ref"D2", num(2))
      .put(ref"D3", formula("=SUM(OFFSET(D1,0,0,3,1))"))
    val result = assertStackIndependent(sheet)
    assertEquals(messages(result).keySet, Set(ref"A1", ref"B1", ref"D3"))
    assert(messages(result)(ref"A1").endsWith(through(ref"A1")), messages(result)(ref"A1"))
    assert(messages(result)(ref"B1").endsWith(through(ref"B1")), messages(result)(ref"B1"))
    // D3 reads D1:D3, itself included: the re-entry is through D3
    assert(messages(result)(ref"D3").endsWith(through(ref"D3")), messages(result)(ref"D3"))
    assertEquals(valueAt(result, ref"C1"), Some(num(7)))
  }

  test("#829: a dynamic reader of a static cycle member re-enters the member (the CI shape)") {
    // A5 = SUM(A1:INDEX(A:A,ROW())) reads itself through the static graph (#743: INDEX at ROW()
    // reads one cell), so Tarjan reports it; B1 reads it at run time, recurses into its range and
    // meets A5 again. This is the book shape RecalcEngineParitySpec found nondeterministic.
    val sheet = (1 to 4)
      .foldLeft(Sheet(S))((s, r) => s.put(ARef.from0(0, r - 1), num(r)))
      .put(ref"A5", formula("=SUM(A1:INDEX(A:A,ROW()))"))
      .put(ref"B1", formula("=INDIRECT(\"A5\")+1"))
    val result = assertStackIndependent(sheet)
    assertEquals(messages(result).keySet, Set(ref"A5", ref"B1"))
    assert(messages(result)(ref"A5").endsWith("Circular reference"), messages(result)(ref"A5"))
    assert(messages(result)(ref"B1").endsWith(through(ref"A5")), messages(result)(ref"B1"))
  }

  /** Static and dynamic cycles, a lookup whose result cell is a dynamic member, and a sibling. */
  private val mixed: Sheet = Sheet(S)
    .put(ref"A1", formula("=B1+1"))
    .put(ref"B1", formula("=A1+1"))
    .put(ref"C1", formula("=INDIRECT(\"D1\")"))
    .put(ref"D1", formula("=INDIRECT(\"C1\")"))
    .put(ref"F1", num(1))
    .put(ref"G1", formula("=INDIRECT(\"G1\")"))
    .put(ref"E1", formula("=VLOOKUP(1,F1:G1,2,FALSE)"))
    .put(ref"X1", num(5))
    .put(ref"Y1", formula("=X1*2"))

  test("#829: static and dynamic cycles side by side, on every path") {
    val result = assertStackIndependent(mixed)
    assertEquals(
      messages(result).keySet,
      Set(ref"A1", ref"B1", ref"C1", ref"D1", ref"E1", ref"G1")
    )
    assert(messages(result)(ref"A1").endsWith("Circular reference"), messages(result)(ref"A1"))
    assert(messages(result)(ref"B1").endsWith("Circular reference"), messages(result)(ref"B1"))
    assert(messages(result)(ref"C1").endsWith(through(ref"C1")), messages(result)(ref"C1"))
    assert(messages(result)(ref"D1").endsWith(through(ref"D1")), messages(result)(ref"D1"))
    assert(messages(result)(ref"G1").endsWith(through(ref"G1")), messages(result)(ref"G1"))
    // a static dependent of a failed member is blocked before any lookup runs, as before
    assert(
      messages(result)(ref"E1").endsWith("Blocked by an upstream formula evaluation failure"),
      messages(result)(ref"E1")
    )
    assertEquals(valueAt(result, ref"Y1"), Some(num(10)))
    // the parallel pass agrees with the sequential one
    assertEquals(recalc(mixed, options.copy(parallelism = 4)), result)
    // read directly, GH-715's lookup contract is kept: a result cell in a cycle is #VALUE!, as it
    // was behind the depth guard (Evaluator.isCycleFailure), never a failure of the lookup
    assertEquals(
      mixed.evaluateFormula("=VLOOKUP(1,F1:G1,2,FALSE)"),
      Right(CellValue.Error(CellError.Value))
    )
  }

  test("#829: an iterative pass still reports a dynamic cycle; it is not a fixpoint") {
    val sheet = Sheet(S)
      .put(ref"A1", formula("=INDIRECT(\"B1\")"))
      .put(ref"B1", formula("=INDIRECT(\"A1\")"))
    val iterative =
      options.copy(iterative = IterativeMode.Force(IterativeCalc(10, BigDecimal("0.001"))))
    val result = assertStackIndependent(sheet, iterative)
    assertEquals(messages(result).keySet, Set(ref"A1", ref"B1"))
    messages(result).values.foreach(m => assert(m.contains(Prefix), m))
  }

  test("#829: COUNTBLANK over a range holding its own cell reads it as the cycle's value") {
    // a cyclic cell is not blank to COUNTBLANK (as a cyclic key is a non-match to a lookup): the
    // re-entry failure is read as the cycle's value, the count completes, and no error is reported
    val sheet = Sheet(S).put(ref"A1", formula("=COUNTBLANK(INDIRECT(\"A1:A2\"))"))
    val result = assertStackIndependent(sheet)
    assertEquals(result.errors, Vector.empty)
    assertEquals(valueAt(result, ref"A1"), Some(num(1)))
  }

  test("#829: a direct evaluation meets the same circular reference") {
    val sheet = Sheet(S).put(ref"A1", formula("=INDIRECT(\"A1\")"))
    sheet.evaluateFormula("=INDIRECT(\"A1\")") match
      case Left(err) => assert(err.message.contains(through(ref"A1")), err.message)
      case Right(v) => fail(s"a self-reference must not evaluate, got $v")
  }

  test("#829: spill references that reach their own anchor are circular, on any stack") {
    // a spill reference is a static edge to its anchor, so a recalculation reports these cycles
    // from the graph; a direct evaluation meets them at the first re-entry into the spill
    val self = Sheet(S).put(ref"A1", formula("=ROWS(A1#)"))
    val result = assertStackIndependent(self)
    assertEquals(messages(result).keySet, Set(ref"A1"))
    assert(messages(result)(ref"A1").endsWith("Circular reference"), messages(result)(ref"A1"))
    val pair = Sheet(S).put(ref"A1", formula("=B1#")).put(ref"B1", formula("=A1#"))
    val paired = assertStackIndependent(pair)
    assertEquals(messages(paired).keySet, Set(ref"A1", ref"B1"))
    self.evaluateFormula("=SUM(A1#)", Clock.system, None, Some(ref"B1")) match
      case Left(err) => assert(err.message.contains(s"${Prefix}S!A1#"), err.message)
      case Right(v) => fail(s"a circular spill must not evaluate, got $v")
    pair.evaluateFormula("=SUM(A1#)", Clock.system, None, Some(ref"C1")) match
      case Left(err) => assert(err.message.contains(s"${Prefix}S!A1#"), err.message)
      case Right(v) => fail(s"a circular spill must not evaluate, got $v")
  }
