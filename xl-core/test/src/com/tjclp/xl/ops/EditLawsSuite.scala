package com.tjclp.xl.ops

import com.tjclp.xl.addressing.{CellRange, SheetName}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.{Arbitrary, Gen}
import org.scalacheck.Prop.forAll

/**
 * The seven algebra laws of ADR-017 §6 for `enum Edit`, as ScalaCheck properties over a `genEdit`
 * that reaches every case (the coverage test pins that): identity, fold, idempotence for the class
 * `EditSchema` declares, lowering coherence with the `Patch` kernel, desugar coherence
 * (`Patch.toEdits`), determinism, and all-or-nothing. The suite is parameterised by the
 * `FormulaSupport` the interpreter runs under, because a law must hold whether or not the formula
 * engine is present: [[EditLawsSpec]] runs it under `FormulaSupport.textOnly`, where every
 * fill/copy/sort/drag over a formula and all four structural edits take the REFUSAL path, and
 * xl-evaluator's `EditLawsEvalSpec` runs the same properties under `EvalFormulaSupport`, where they
 * take the semantic one — the only place an idempotence or fold violation in `insert-rows` or a
 * shifted `fill` can surface.
 */
abstract class EditLawsSuite(support: FormulaSupport, label: String) extends ScalaCheckSuite:

  import EditGenerators.*

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(200)

  private given FormulaSupport = support

  test(s"the laws below run under the $label FormulaSupport") {
    assertEquals(summon[FormulaSupport], support)
  }

  test("validateFill's refusal is reached through genEdit (fills are not all direction-legal)") {
    // Deterministic draw: a fixed seed over a large sample, so the coverage claim cannot flake on
    // an unlucky run (fills are ~1 case in 49 and only one target in four breaks the rule).
    val seeded = Gen
      .listOfN(3000, genEdit)
      .pureApply(Gen.Parameters.default, org.scalacheck.rng.Seed(0x5eed5eedL))
    val fills = seeded.collect { case f: Edit.Fill => f }
    val refused = fills.filter(f => Edit.validate(f).isLeft)
    assert(fills.size >= 4, s"genEdit yielded only ${fills.size} fills")
    assert(refused.nonEmpty, "no generated fill was refused by the direction rule")
    assert(refused.size < fills.size, "every generated fill was refused")
    refused.foreach(f =>
      Edit.validate(f) match
        case Left(XLError.InvalidReference(msg)) => assert(msg.startsWith("Fill "), msg)
        case other => fail(s"expected the direction rule's InvalidReference, got $other")
    )
  }
  private given Arbitrary[Edit] = Arbitrary(genEdit)
  private given Arbitrary[Vector[Edit]] = Arbitrary(genEdits)

  private val scope: Scope = Scope.of(data)

  private def run(wb: Workbook, edits: Vector[Edit], s: Scope = scope): XLResult[Applied] =
    Edit.applyAll(wb, edits, s)

  /** Results compared on their outcome: the workbook, or the failure's root cause. */
  private def outcome(r: XLResult[Applied]): Either[XLError, Workbook] =
    r.map(_.workbook).left.map(_.root)

  test("genEdit reaches every Edit case (the laws quantify over all of them)") {
    val seen = (1 to 3000).flatMap(_ => genEdit.sample).map(EditSchema.nameOf).toSet
    val all = EditSchema.all.map(_.name).toSet
    assertEquals(all.diff(seen), Set.empty[String])
    assertEquals(caseGenerators.size, EditSchema.all.size)
  }

  property("identity: applying no edits returns the workbook unchanged") {
    forAll(Gen.const(baseWorkbook)) { (wb: Workbook) =>
      assertEquals(run(wb, Vector.empty), Right(Applied(wb, Vector.empty, scope)))
      true
    }
  }

  property("fold: applyAll(a ++ b) == applyAll(a) then applyAll(b) under the resulting scope") {
    forAll { (a: Vector[Edit], b: Vector[Edit]) =>
      val together = outcome(run(baseWorkbook, a ++ b))
      val stepwise = outcome(
        run(baseWorkbook, a).flatMap(first => run(first.workbook, b, first.scope))
      )
      assertEquals(together, stepwise)
      true
    }
  }

  property("idempotence: the declared class gives the same workbook applied once or twice") {
    forAll(genEdit.retryUntil(e => EditSchema.specOf(e).idempotent)) { (e: Edit) =>
      val once = outcome(run(baseWorkbook, Vector(e)))
      val twice = outcome(run(baseWorkbook, Vector(e, e)))
      assertEquals(twice, once, s"not idempotent: $e")
      true
    }
  }

  property("determinism: the same edits on the same workbook give equal results") {
    forAll { (edits: Vector[Edit]) =>
      assertEquals(run(baseWorkbook, edits), run(baseWorkbook, edits))
      true
    }
  }

  property(
    "all-or-nothing: a failing edit fails the whole sequence at its index, nothing escapes"
  ) {
    forAll { (good: Vector[Edit]) =>
      val bad = Edit.RemoveSheet(missing)
      run(baseWorkbook, good) match
        case Right(_) =>
          run(baseWorkbook, good :+ bad) match
            case Left(XLError.EditFailed(index, op, cause)) =>
              assertEquals(index, good.size + 1)
              assertEquals(op, "remove-sheet")
              // GH-615: the candidates are whatever sheets the prefix left; the name is fixed
              cause match
                case XLError.SheetNotFound(name, _) => assertEquals(name, missing.value)
                case other => fail(s"expected SheetNotFound, got $other")
            case other => fail(s"expected EditFailed at ${good.size + 1}, got $other")
        case Left(XLError.EditFailed(index, _, _)) =>
          // The prefix already fails: the failure is reported at the SAME index with `bad` appended
          assertEquals(run(baseWorkbook, good :+ bad).left.map(_.opIndex), Left(Some(index)))
        case Left(other) => fail(s"applyAll must wrap every failure in EditFailed, got $other")
      true
    }
  }

  // ----- coherence with the Patch kernel -----

  private def sheetOf(r: XLResult[Applied], name: SheetName): XLResult[Sheet] =
    r.flatMap(_.workbook(name))

  private val genLocalEdit: Gen[Edit] =
    genSheetEdit.retryUntil(e => Edit.lower(e, baseData).isDefined)

  property("lowering coherence: lower(e) == Some(p) implies applying e equals applyPatch(p)") {
    forAll(genLocalEdit) { (e: Edit) =>
      val wb = Workbook(baseData)
      (Edit.lower(e, baseData), run(wb, Vector(e), Scope.of(data))) match
        case (Some(p), Right(applied)) =>
          assertEquals(applied.workbook(data), Right(Patch.applyPatch(baseData, p)), s"edit: $e")
        case (Some(_), Left(XLError.EditFailed(_, _, cause))) =>
          // `lower` refuses what `validate` refuses and honours the qualifier (an edit aimed at a
          // sheet other than `existing` lowers to None), so a lowerable edit fails only when the
          // support rejects its formula text (text-only: empty; evaluator: unparsable).
          cause match
            case XLError.FormulaError(_, _) => ()
            case unexpected =>
              fail(s"lowerable edit $e failed for an unexpected reason: $unexpected")
        case (None, _) => fail("genLocalEdit only yields lowerable edits")
        case (Some(_), Left(unwrapped)) => fail(s"unwrapped failure $unwrapped")
      true
    }
  }

  test("lowering coherence is not vacuous: most lowerable edits apply") {
    val samples = (1 to 500).flatMap(_ => genLocalEdit.sample)
    val applied = samples.count(e => run(Workbook(baseData), Vector(e), Scope.of(data)).isRight)
    assert(applied > samples.size / 2, s"only $applied of ${samples.size} lowerable edits applied")
  }

  test("lowering coherence pins the Clear→unmerge branch: a contents clear over a merged range") {
    val range = CellRange.parse("E4:E5").fold(e => fail(e), identity)
    val merged = CellRange.parse("E5:F6").fold(e => fail(e), identity)
    assert(baseData.mergedRanges.contains(merged), "fixture: baseData merges E5:F6")
    val e = Edit.Clear(Area(None, range), ClearWhat.contents)
    val viaPatch = Edit.lower(e, baseData).map(Patch.applyPatch(baseData, _))
    val viaEdit = sheetOf(run(Workbook(baseData), Vector(e), Scope.of(data)), data).toOption
    assert(viaPatch.isDefined, "a contents clear lowers")
    assertEquals(viaPatch, viaEdit)
    assert(viaPatch.exists(s => !s.mergedRanges.contains(merged)), "the patch side unmerges")
    assert(viaEdit.exists(s => !s.mergedRanges.contains(merged)), "the edit side unmerges")
  }

  property("desugar coherence: Patch.toEdits(p, sheet) applied equals Patch.applyPatch(sheet, p)") {
    forAll(genPatch(baseData)) { (p: Patch) =>
      Patch.toEdits(p, baseData) match
        case Some(edits) =>
          val viaEdits = sheetOf(run(Workbook(baseData), edits, Scope.of(data)), data)
          assertEquals(viaEdits, Right(Patch.applyPatch(baseData, p)), s"patch: $p")
        case None => ()
      true
    }
  }

  test("desugar coherence is not vacuous: most generated patches desugar") {
    val samples = (1 to 500).flatMap(_ => genPatch(baseData).sample)
    val desugared = samples.count(p => Patch.toEdits(p, baseData).isDefined)
    assert(desugared > samples.size / 2, s"only $desugared of ${samples.size} patches desugar")
  }

  // ===== GH-612: whole-column references through DragFormula and Fill =====
  // Under the semantic support `$A:$A` keeps its form while the criterion drags; under textOnly the
  // edit takes the refusal path like every other formula shift (the suite's standing split).

  private def formulaTextAt(wb: Workbook, ref: com.tjclp.xl.addressing.ARef): Option[String] =
    wb(data).toOption.flatMap(_.cells.get(ref)).map(_.value).collect {
      case com.tjclp.xl.cells.CellValue.Formula(text, _, _) => text
    }

  private def assertShiftedOrRefused(
    result: XLResult[Applied],
    expected: Map[com.tjclp.xl.addressing.ARef, String]
  ): Unit =
    outcome(result) match
      case Right(wb) =>
        assertNotEquals(support, FormulaSupport.textOnly, "textOnly cannot shift a formula")
        expected.foreach { (ref, text) =>
          assertEquals(formulaTextAt(wb, ref), Some(text), s"at ${ref.toA1}")
        }
      case Left(XLError.UnsupportedCapability(op, _, _)) =>
        assertEquals(support, FormulaSupport.textOnly)
        assertEquals(op, "shift")
      case Left(other) => fail(s"unexpected failure: $other")

  test(
    "GH-612: DragFormula keeps $A:$A and drags only the criterion (or refuses without a parser)"
  ) {
    import com.tjclp.xl.addressing.ARef
    val range = CellRange(ARef.from0(25, 0), ARef.from0(25, 2)) // Z1:Z3
    val edit =
      Edit.DragFormula(Area(Some(data), range), "=COUNTIF($A:$A,B1)", ARef.from0(25, 0), None)
    assertShiftedOrRefused(
      run(baseWorkbook, Vector(edit)),
      Map(
        ARef.from0(25, 0) -> "COUNTIF($A:$A,B1)",
        ARef.from0(25, 1) -> "COUNTIF($A:$A,B2)",
        ARef.from0(25, 2) -> "COUNTIF($A:$A,B3)"
      )
    )
  }

  test("GH-612: Fill down keeps $A:$A in the repeated formula (or refuses without a parser)") {
    import com.tjclp.xl.addressing.ARef
    import com.tjclp.xl.cells.CellValue
    val z1 = ARef.from0(25, 0)
    val seeded = baseWorkbook.update(data, _.put(z1, CellValue.Formula("COUNTIF($A:$A,B1)", None)))
    val source = Area(Some(data), CellRange(z1, z1))
    val target = CellRange(z1, ARef.from0(25, 2)) // Z1:Z3
    val wb = seeded.fold(e => fail(e.message), identity)
    assertShiftedOrRefused(
      run(wb, Vector(Edit.Fill(source, target, Edit.FillDir.Down))),
      Map(
        ARef.from0(25, 1) -> "COUNTIF($A:$A,B2)",
        ARef.from0(25, 2) -> "COUNTIF($A:$A,B3)"
      )
    )
  }

  test("PatchSpec's kernel is untouched: Patch.applyPatch on a Put still puts") {
    val p =
      Patch.Put(com.tjclp.xl.addressing.ARef.from0(0, 0), com.tjclp.xl.cells.CellValue.Text("x"))
    assertEquals(
      Patch
        .applyPatch(Sheet(data), p)
        .cells
        .get(com.tjclp.xl.addressing.ARef.from0(0, 0))
        .map(_.value),
      Some(com.tjclp.xl.cells.CellValue.Text("x"))
    )
  }
