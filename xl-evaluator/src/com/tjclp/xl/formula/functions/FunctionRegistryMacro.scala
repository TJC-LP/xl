package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import scala.quoted.*

object FunctionRegistryMacro:
  def collect[T: Type](using Quotes): Expr[List[FunctionSpec[?]]] =
    import quotes.reflect.*

    val targetType = TypeRepr.of[T]
    val moduleSym = targetType.termSymbol
    val classSym =
      if moduleSym != Symbol.noSymbol then moduleSym.moduleClass
      else targetType.typeSymbol
    val moduleRef =
      if moduleSym != Symbol.noSymbol then Ref(moduleSym)
      else Ref(targetType.typeSymbol.companionModule)
    val specSym = TypeRepr.of[FunctionSpec[Any]].typeSymbol

    val specFields = targetType.baseClasses
      .flatMap(_.declaredFields)
      .distinctBy(_.name)
      .filter { field =>
        field.tree match
          case v: ValDef => v.tpt.tpe.dealias.typeSymbol == specSym
          case _ => false
      }

    val entries = specFields.map { field =>
      Select.unique(moduleRef, field.name).asExprOf[FunctionSpec[?]]
    }

    Expr.ofList(entries)
