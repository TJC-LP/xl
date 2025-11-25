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
  export formula.FormulaParser
  export formula.FormulaPrinter
  export formula.FunctionParser

  // Typed expression AST
  export formula.TExpr

  // Dependency analysis
  export formula.DependencyGraph

  // Evaluation
  export formula.Evaluator
  export formula.Clock

  // Error types
  export formula.EvalError
  export formula.ParseError

  // SheetEvaluator object and extension methods
  export formula.SheetEvaluator
  export formula.SheetEvaluator.*

  // Display strategy with formula evaluation
  // Enables: given FormulaDisplayStrategy = EvaluatingFormulaDisplay.evaluating
  export formula.display.EvaluatingFormulaDisplay

// Make available at com.tjclp.xl.*
export formulaExports.*
