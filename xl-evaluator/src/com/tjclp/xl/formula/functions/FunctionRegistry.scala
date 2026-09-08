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

  /**
   * GH-274: upper-case names of functions flagged `FunctionFlags.dynamicDeps` (INDIRECT, and OFFSET
   * since GH-301). Used by `DependencyGraph.dynamicCells` as a cheap substring pre-filter before
   * parsing — zero parse cost for sheets without dynamic references.
   */
  lazy val dynamicFunctionNames: List[String] =
    byName.values.filter(_.flags.dynamicDeps).map(_.name.toUpperCase).toList.sorted

  /**
   * GH-588: upper-case names of functions flagged `FunctionFlags.volatile` (TODAY, NOW, RAND,
   * RANDBETWEEN). `WorkbookAudit.volatileFunctions` is this list as a set — the flag on the spec is
   * the single source of truth, there is no name table to keep in step.
   */
  lazy val volatileFunctionNames: List[String] =
    byName.values.filter(_.flags.volatile).map(_.name.toUpperCase).toList.sorted
