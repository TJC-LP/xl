package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.formula.graph.CompactGraphGens
import com.tjclp.xl.testing.SmallStack
import munit.ScalaCheckSuite
import org.scalacheck.Prop

/**
 * #829 (#828): a recalculation is a function of the workbook and its options alone — the same
 * `RecalcResult` on every run, on the runner's stack and on a small one, however warm the JIT. The
 * generated books mix static cycles with INDIRECT/OFFSET readers, so dynamic cycles are common;
 * every one must end at the re-entry guard, never at a depth or stack guard.
 */
class RecalcDeterminismSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(80)

  private val clock = Clock.fixedDate(java.time.LocalDate.of(2026, 1, 2))
  private def opts =
    RecalcOptions(clock = clock, rng = Rng.seeded(7L), iterative = IterativeMode.Off)
  private val Guards = List("exhausted the stack", "recursion depth exceeded")

  private def dynamicCycles(result: RecalcResult): Int =
    result.errors.count(_.error.message.contains("Circular reference through "))

  property("recalculating a book again, on any stack, gives the same result") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val first = wb.recalculate(opts)
      val second = wb.recalculate(opts)
      val wide = SmallStack.run(1024L * 1024L, "recalc-1mb")(wb.recalculate(opts))
      val narrow = SmallStack.run(512L * 1024L, "recalc-512kb")(wb.recalculate(opts))
      val guarded = first.errors.filter(e => Guards.exists(e.error.message.contains))
      Prop.collect(s"dynamic cycles: ${dynamicCycles(first)}")(
        Prop(first == second) :| "twice on the runner thread" &&
          Prop(first == wide) :| "on a 1 MB stack" &&
          Prop(first == narrow) :| "on a 512 KB stack" &&
          Prop(guarded.isEmpty) :| s"a cycle reached a depth or stack guard: $guarded"
      )
    }
  }

  property("the parallel and iterative passes are deterministic too") {
    Prop.forAll(CompactGraphGens.genWorkbookOf(40)) { wb =>
      val par = opts.copy(parallelism = 3)
      val iter = opts.copy(iterative = IterativeMode.Force(IterativeCalc(20, BigDecimal("0.001"))))
      val iterated = wb.recalculate(iter)
      Prop(wb.recalculate(par) == wb.recalculate(par)) :| "parallel twice" &&
      Prop(iterated == wb.recalculate(iter)) :| "iterative twice" &&
      Prop(iterated == SmallStack.run(1024L * 1024L, "recalc-iter-1mb")(wb.recalculate(iter))) :|
        "iterative on a 1 MB stack"
    }
  }
