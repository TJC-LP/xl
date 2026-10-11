package com.tjclp.xl.formula.graph

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.formula.ast.{BinarySpine, TExpr}
import com.tjclp.xl.formula.eval.Evaluator
import com.tjclp.xl.formula.functions.{ArgValue, CriteriaRangeResize, ReferenceOperators}

/**
 * Value reads for the positioned-evaluation cycle check. Reference inspection (ROW, COLUMN, ROWS,
 * COLUMNS and AREAS) needs the address or shape, not the referenced cells' values. A computed
 * reference still reads its selectors: ROW(INDEX(A:A, B1)) depends on B1, not A:A. The ordinary
 * dependency graph remains conservative for structural edits and impact analysis.
 */
private[xl] object ValueDependencies:
  private final case class Bound(expr: TExpr[?], env: Map[String, Bound])
  private val inspectors = Set("ROW", "COLUMN", "ROWS", "COLUMNS", "AREAS")

  def extract(expr: TExpr[?], bounds: Option[CellRange]): Set[ARef] =
    def cells(range: CellRange): Set[ARef] =
      bounds.fold(range.cells.toSet)(b => range.intersect(b).fold(Set.empty[ARef])(_.cells.toSet))

    def loop(e: TExpr[?], referenceOnly: Boolean, env: Map[String, Bound]): Set[ARef] =
      def value(child: TExpr[?]): Set[ARef] = loop(child, false, env)
      def binding(name: String): Set[ARef] =
        env.get(name).fold(Set.empty[ARef])(b => loop(b.expr, referenceOnly, b.env))
      def arg(a: ArgValue, addressOnly: Boolean): Set[ARef] = a match
        case ArgValue.Expr(child) => loop(child, addressOnly, env)
        case ArgValue.Range(location) =>
          // Non-static lexical locations print as their declared identifier. This also covers
          // range-slot bindings without confusing qualified names, cell addresses or errors with
          // legal LET names. Follow the binding in its defining scope and the consumer's mode.
          val bound = if location.staticRange.isEmpty then env.get(location.toA1) else None
          bound match
            case Some(b) => loop(b.expr, addressOnly, b.env)
            case None => if addressOnly then Set.empty else location.localCellsBounded(bounds)
        case ArgValue.Cells(range) => if addressOnly then Set.empty else cells(range)

      e match
        case chain if BinarySpine.isBinary(chain) =>
          BinarySpine.operandList(chain).foldLeft(Set.empty[ARef])(_ ++ value(_))
        case TExpr.Coerced(inner, _) => loop(inner, referenceOnly, env)
        case TExpr.ToInt(inner) => value(inner)
        case TExpr.UnaryPlus(inner) => value(inner)
        case TExpr.Percent(inner) => value(inner)
        case TExpr.DateToSerial(inner) => value(inner)
        case TExpr.DateTimeToSerial(inner) => value(inner)
        case TExpr.BindingRef(name) => binding(name)
        case TExpr.CoercedBindingRef(name, _) => binding(name)
        case TExpr.Let(bindings, body) =>
          val (scoped, bindingReads) = bindings.foldLeft((env, Set.empty[ARef])) {
            case ((scope, reads), (name, definition)) =>
              (
                scope.updated(name, Bound(definition, scope)),
                reads ++ loop(definition, referenceDefinition(definition), scope)
              )
          }
          // Plain-cell LET stores a reference without reading its values. Eager scalar bindings
          // and reference selectors still evaluate, even if the body does not use the binding.
          bindingReads ++ loop(body, referenceOnly, scoped)
        case TExpr.Aggregate(_, location) => arg(ArgValue.Range(location), false)
        case call: TExpr.Call[?] =>
          val name = call.spec.name
          val args = ReferenceOperators.dependencyValues(
            name,
            CriteriaRangeResize.resizedArgs(name, call.spec.argSpec.toValues(call.args))
          )
          args.zipWithIndex.foldLeft(Set.empty[ARef]) { case (acc, (a, index)) =>
            // CELL's supported info types inspect metadata; info_type itself is a value read.
            val addressOnly = inspectors.contains(name) || (name == "CELL" && index == 1) ||
              (referenceOnly && (name match
                case "INDEX" | "OFFSET" => index == 0
                case "IF" | "CHOOSE" => index > 0
                case "IFS" => index % 2 == 1
                case "SWITCH" => index > 0 && (index % 2 == 0 || index == args.size - 1)
                case "XLOOKUP" => index == 2
                case "(union)" | "(intersection)" | "(range)" => true
                case _ => false))
            acc ++ arg(a, addressOnly)
          }
        case _ =>
          if referenceOnly then Set.empty
          else DependencyGraph.extractDependenciesBounded(e, bounds)
    loop(expr, false, Map.empty)

  private def referenceDefinition(expr: TExpr[?]): Boolean = expr match
    // Literal range bindings are substituted by the parser and skipped by LET evaluation.
    case _: TExpr.RangeRef | _: TExpr.SheetRange => true
    case _ => Evaluator.denotesReference(expr)
