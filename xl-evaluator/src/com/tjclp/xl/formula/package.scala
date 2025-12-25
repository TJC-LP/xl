package com.tjclp.xl

/**
 * Formula package re-exports.
 *
 * Re-exports types from subpackages to maintain backward compatibility with imports like:
 * {{{
 * import com.tjclp.xl.formula.{FormulaParser, TExpr, Evaluator}
 * }}}
 */
package object formula:
  // Parser
  export parser.FormulaParser
  export parser.ParseError

  // Printer
  export printer.FormulaPrinter
  export printer.FormulaShifter

  // AST
  export ast.TExpr
  export ast.ExprValue

  // Evaluation
  export eval.Evaluator
  export eval.EvalError
  export eval.Aggregator
  export eval.CriteriaMatcher
  export eval.SheetEvaluator
  export eval.WorkbookEvaluator

  // Functions
  export functions.FunctionSpec
  export functions.FunctionSpecs
  export functions.FunctionRegistry
  export functions.EvalContext
  export functions.ArgPrinter
  export functions.ArgValue

  // Graph
  export graph.DependencyGraph

  // Display
  export display.EvaluatingFormulaDisplay
