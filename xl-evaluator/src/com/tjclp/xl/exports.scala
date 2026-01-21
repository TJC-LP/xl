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

  // Typed expression AST
  export formula.ast.TExpr

  // Dependency analysis
  export formula.graph.DependencyGraph

  // Evaluation
  export formula.eval.Evaluator
  export formula.Clock

  // Error types
  export formula.eval.EvalError
  export formula.parser.ParseError

  // SheetEvaluator object and extension methods
  export formula.eval.SheetEvaluator
  export formula.eval.SheetEvaluator.*

  // WorkbookEvaluator extension methods (withCachedFormulas)
  export formula.eval.WorkbookEvaluator
  export formula.eval.WorkbookEvaluator.*

  // DependentRecalculation extension methods (GH-163)
  // Note: Extension methods with default parameters must be imported directly,
  // not via wildcard export (Scala 3 compiler bug)
  export formula.eval.DependentRecalculation

  // Display strategy with formula evaluation
  // The evaluating given has higher priority than default due to LowPriority pattern
  export formula.display.EvaluatingFormulaDisplay
  export formula.display.EvaluatingFormulaDisplay.given

// Make available at com.tjclp.xl.*
export formulaExports.*
