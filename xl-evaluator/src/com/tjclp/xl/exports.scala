package com.tjclp.xl

/**
 * Formula system exports.
 *
 * When xl-evaluator is a dependency, this makes formula types available via:
 * {{{
 * import com.tjclp.xl.{*, given}
 *
 * sheet.evaluateFormula("=SUM(A1:A10)")  // SheetEvaluator extension
 * FormulaParser.parse("=A1+B1")          // Parser
 * DependencyGraph.fromSheet(sheet)       // Dependency analysis
 * }}}
 */
object formulaExports:
  // Parser and printer
  export formula.parser.FormulaParser
  export formula.printer.FormulaPrinter
  export formula.functions.FunctionRegistry

  // Typed expression AST (GH-612: RangeForm is the whole-column / whole-row marker on range nodes)
  export formula.ast.TExpr
  export formula.ast.RangeForm

  // Dependency analysis
  export formula.graph.DependencyGraph

  // Evaluation
  export formula.eval.Evaluator
  export formula.Clock
  // GH-115: randomness capability for RAND/RANDBETWEEN (Rng.system / Rng.seeded)
  export formula.Rng

  // Error types
  export formula.eval.EvalError
  export formula.parser.ParseError

  // Array formula support
  export formula.eval.ArrayResult

  // SheetEvaluator object and extension methods
  export formula.eval.SheetEvaluator
  export formula.eval.SheetEvaluator.*

  // Number-format inheritance for formula cells (GH-184, opt-in via putFormulaInheriting)
  export formula.eval.FormulaFormatting

  // WorkbookEvaluator extension methods (withCachedFormulas, recalculate, evaluateFormula)
  export formula.eval.WorkbookEvaluator
  export formula.eval.WorkbookEvaluator.*

  // Recalc result types (workbook-level total recalculation with per-cell errors)
  export formula.eval.{CellEvalError, IterativeCalc, RecalcResult}
  // GH-492: per-strongly-connected-component fixpoint verdicts on RecalcResult.cycles
  export formula.eval.SccReport

  // GH-419: explicit data-table cache seeding (autoNoTable books never recompute tables on open)
  export formula.eval.DataTableSeeder
  export formula.eval.DataTableSeeder.*
  // GH-453: seeding report + per-table warnings for circular books
  export formula.eval.{DataTableSeedReport, SeedTableWarning}

  // DependentRecalculation extension methods (GH-163)
  // Note: Extension methods with default parameters must be imported directly,
  // not via wildcard export (Scala 3 compiler bug)
  export formula.eval.DependentRecalculation

  // ADR-017 §2.8/§2.9 (GH-559): the options record behind wb.recalculate(options) /
  // recalculateAfterEdit / recalculateUncached, the reference-rewriting sheet renamer, and the
  // structural editor + string-level formula rewriting the CLI and scripts share. The objects are
  // exported, never their `.*` members: StructuralEditor's extension block carries default
  // arguments (the wildcard-export landmine above).
  export formula.eval.{IterativeMode, RecalcOptions, SheetRenamer, StructuralEditor}
  export formula.printer.{FormulaOps, FormulaShifter}
  // ADR-017 §2.12 (W2.1): the evaluator-backed FormulaSupport behind wb.edit / sheet.edit.
  export formula.eval.EvalFormulaSupport
  // Explicit type + val pair (an exported companion loses its type alias for external consumers).
  type QualifiedRef = formula.graph.DependencyGraph.QualifiedRef
  val QualifiedRef: formula.graph.DependencyGraph.QualifiedRef.type =
    formula.graph.DependencyGraph.QualifiedRef

  // ADR-017 §2.10: inspection — the bounded cross-sheet graph behind `cell`/`deps`, the audit
  // buckets behind `audit`, the summary behind `describe --full`. `wb.describe` / `wb.audit` come
  // from WorkbookInspect's extension block, which carries no default arguments (wildcard-safe).
  export formula.graph.QualifiedGraph
  export formula.eval.{SheetSummary, WorkbookAudit, WorkbookInspect, WorkbookSummary}
  export formula.eval.WorkbookInspect.*

  // Display strategy with formula evaluation
  // The evaluating given has higher priority than default due to LowPriority pattern
  export formula.display.EvaluatingFormulaDisplay
  export formula.display.EvaluatingFormulaDisplay.given

// Make available at com.tjclp.xl.*
export formulaExports.*

/**
 * The `given FormulaSupport` of `import com.tjclp.xl.{*, given}` once xl-evaluator is on the
 * classpath (ADR-017 §2.12): `wb.edit` / `sheet.edit` shift, restructure and rename through the
 * evaluator. xl-core deliberately has no default — a formula-blind interpreter would write `#REF!`
 * — so the choice is made here, explicitly, rather than through a low-priority trait (which does
 * not survive export forwarding).
 */
given ops.FormulaSupport = formula.eval.EvalFormulaSupport
