package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

object FunctionRegistry:
  inline def all: List[FunctionSpec[?]] =
    ${ FunctionRegistryMacro.collect[FunctionSpecs.type] }

  private lazy val byName: Map[String, FunctionSpec[?]] =
    all.map(spec => spec.name.toUpperCase -> spec).toMap

  def lookup(name: String): Option[FunctionSpec[?]] =
    byName.get(name.toUpperCase)

  def allNames: List[String] =
    byName.keys.toList.sorted
