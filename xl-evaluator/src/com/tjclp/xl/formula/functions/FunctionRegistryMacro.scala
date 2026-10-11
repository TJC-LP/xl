package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import scala.quoted.*

object FunctionRegistryMacro:
  def collect[T: Type](using Quotes): Expr[List[FunctionSpec[?]]] =
    import quotes.reflect.*
    val specSym = TypeRepr.of[FunctionSpec[Any]].typeSymbol
    collectWhere[T](_.dealias.typeSymbol == specSym)

  /**
   * GH-714: the specs of `T` whose declared type is a `FunctionSpec[R]` — the functions whose
   * result is an `R` by their type, so a list derived from it needs no hand upkeep.
   */
  def collectReturning[T: Type, R: Type](using Quotes): Expr[List[FunctionSpec[?]]] =
    import quotes.reflect.*
    val target = TypeRepr.of[FunctionSpec[R]]
    collectWhere[T](_ <:< target)

  private def collectWhere[T: Type](using
    q: Quotes
  )(keep: q.reflect.TypeRepr => Boolean): Expr[List[FunctionSpec[?]]] =
    import q.reflect.*

    val targetType = TypeRepr.of[T]
    val moduleSym = targetType.termSymbol
    val moduleRef =
      if moduleSym != Symbol.noSymbol then Ref(moduleSym)
      else Ref(targetType.typeSymbol.companionModule)

    val specFields = targetType.baseClasses
      .flatMap(_.declaredFields)
      .distinctBy(_.name)
      .filter { field =>
        field.tree match
          case v: ValDef => keep(v.tpt.tpe)
          case _ => false
      }

    val entries = specFields.map { field =>
      Select.unique(moduleRef, field.name).asExprOf[FunctionSpec[?]]
    }

    Expr.ofList(entries)
