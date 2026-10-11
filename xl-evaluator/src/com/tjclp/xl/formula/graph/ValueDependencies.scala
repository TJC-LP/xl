package com.tjclp.xl.formula.graph

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
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

  /**
   * The functions whose result can be a reference, each with the arguments (by index, given the
   * argument count) its reference passes through. In a reference-only position (ROWS(...), a LET
   * binding that ROW inspects) those arguments are not read; the others still are. INDIRECT passes
   * none: its text is a value read, and the reference it builds is not static.
   *
   * This duplicates what the evaluator knows. Keep it in sync with [[Evaluator.referenceFunctions]]
   * (the calls that select or return a reference) and with every
   * [[com.tjclp.xl.formula.functions.FunctionSpec.Referencing]] spec (OFFSET, INDIRECT, INDEX, the
   * intersection operator). ValueDependenciesSpec walks `FunctionRegistry.all` and fails when a
   * name from either is missing here.
   */
  private[graph] val referencePassing: Map[String, (Int, Int) => Boolean] = Map(
    "INDEX" -> ((index, _) => index == 0),
    "OFFSET" -> ((index, _) => index == 0),
    "INDIRECT" -> ((_, _) => false),
    "IF" -> ((index, _) => index > 0),
    "CHOOSE" -> ((index, _) => index > 0),
    "IFS" -> ((index, _) => index % 2 == 1),
    "SWITCH" -> ((index, count) => index > 0 && (index % 2 == 0 || index == count - 1)),
    "XLOOKUP" -> ((index, _) => index == 2),
    ReferenceOperators.UnionName -> ((_, _) => true),
    ReferenceOperators.IntersectionName -> ((_, _) => true)
  )

  /**
   * The cells whose values `expr` reads on its own sheet, with ranges bounded by `bounds`. `home`
   * is that sheet's name. A reference qualified with it (`Data!D5` in a formula on Data) is a read
   * like `D5`. Sheet names compare case-insensitively, as Excel's do and as
   * [[ReferenceOperators.dependencyValues]] compares them. Any other qualified reference belongs to
   * another sheet and reads nothing here.
   */
  def extract(expr: TExpr[?], bounds: Option[CellRange], home: Option[SheetName]): Set[ARef] =
    Reader(bounds, home).reads(expr)

  /**
   * [[extract]] for many formulas on one sheet. Each distinct range is expanded once for all of
   * them, as `DependencyGraph.fromSheet` does (GH-522).
   */
  final class Reader(bounds: Option[CellRange], home: Option[SheetName]):
    private val expanded = scala.collection.mutable.HashMap.empty[(ARef, ARef), Set[ARef]]

    def reads(expr: TExpr[?]): Set[ARef] = loop(expr, false, Map.empty)

    private def cells(range: CellRange): Set[ARef] =
      expanded.getOrElseUpdate(
        (range.start, range.end),
        bounds.fold(range.cells.toSet)(b => range.intersect(b).fold(Set.empty[ARef])(_.cells.toSet))
      )

    private def local(sheet: SheetName): Boolean =
      home.exists(_.value.equalsIgnoreCase(sheet.value))

    private def loop(e: TExpr[?], referenceOnly: Boolean, env: Map[String, Bound]): Set[ARef] =
      def value(child: TExpr[?]): Set[ARef] = loop(child, false, env)
      def binding(name: String): Set[ARef] =
        env.get(name).fold(Set.empty[ARef])(b => loop(b.expr, referenceOnly, b.env))
      def cell(at: ARef): Set[ARef] = if referenceOnly then Set.empty else Set(at)
      def range(r: CellRange): Set[ARef] = if referenceOnly then Set.empty else cells(r)
      def arg(a: ArgValue, addressOnly: Boolean): Set[ARef] = a match
        case ArgValue.Expr(child) => loop(child, addressOnly, env)
        case ArgValue.Range(location) =>
          // Non-static lexical locations print as their declared identifier. This also covers
          // range-slot bindings without confusing qualified names, cell addresses or errors with
          // legal LET names. Follow the binding in its defining scope and the consumer's mode.
          val bound = if location.staticRange.isEmpty then env.get(location.toA1) else None
          bound match
            case Some(b) => loop(b.expr, addressOnly, b.env)
            case None if addressOnly => Set.empty
            case None =>
              location match
                case TExpr.RangeLocation.Local(r, _) => cells(r)
                case TExpr.RangeLocation.CrossSheet(sheet, r, _) if local(sheet) => cells(r)
                case _ => Set.empty
        case ArgValue.Cells(r) => if addressOnly then Set.empty else cells(r)

      e match
        case chain if BinarySpine.isBinary(chain) =>
          BinarySpine.operandList(chain).foldLeft(Set.empty[ARef])(_ ++ value(_))
        case TExpr.Ref(at, _, _) => cell(at)
        case TExpr.PolyRef(at, _) => cell(at)
        case TExpr.SheetRef(sheet, at, _, _) => if local(sheet) then cell(at) else Set.empty
        case TExpr.SheetPolyRef(sheet, at, _) => if local(sheet) then cell(at) else Set.empty
        case TExpr.RangeRef(r, _) => range(r)
        case TExpr.SheetRange(sheet, r, _) => if local(sheet) then range(r) else Set.empty
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
          val passes = if referenceOnly then referencePassing.get(name) else None
          args.zipWithIndex.foldLeft(Set.empty[ARef]) { case (acc, (a, index)) =>
            // CELL's supported info types inspect metadata; info_type itself is a value read.
            val addressOnly = inspectors.contains(name) || (name == "CELL" && index == 1) ||
              passes.exists(_(index, args.size))
            acc ++ arg(a, addressOnly)
          }
        // literals, names and external references: no cell of this sheet
        case _ =>
          if referenceOnly then Set.empty
          else DependencyGraph.extractDependenciesBounded(e, bounds)

  private def referenceDefinition(expr: TExpr[?]): Boolean = expr match
    // Literal range bindings are substituted by the parser and skipped by LET evaluation.
    case _: TExpr.RangeRef | _: TExpr.SheetRange => true
    case _ => Evaluator.denotesReference(expr)
