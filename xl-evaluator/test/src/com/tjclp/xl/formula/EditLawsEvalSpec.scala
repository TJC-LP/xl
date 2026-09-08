package com.tjclp.xl.formula

import com.tjclp.xl.formula.eval.EvalFormulaSupport
import com.tjclp.xl.ops.EditLawsSuite

/**
 * The seven `Edit` laws (ADR-017 §6) under `EvalFormulaSupport`: the same properties xl-core runs
 * under the text-only support, now with fill/copy/sort/drag shifting real formulas and
 * insert/delete rows/cols and rename rewriting the workbook — the semantics, not the refusal.
 */
class EditLawsEvalSpec extends EditLawsSuite(EvalFormulaSupport, "evaluator")
