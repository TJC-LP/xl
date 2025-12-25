package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

object FunctionSpecs
    extends FunctionSpecsMath
    with FunctionSpecsAggregate
    with FunctionSpecsLogical
    with FunctionSpecsConditional
    with FunctionSpecsTypeCheck
    with FunctionSpecsDateTime
    with FunctionSpecsFinancial
    with FunctionSpecsLookup
    with FunctionSpecsReference
    with FunctionSpecsText
