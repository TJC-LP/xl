package com.tjclp.xl.formula.graph

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.{FunctionSpecs, FunctionRegistry, ArgValue}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}

import com.tjclp.xl.workbooks.Workbook

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.CellRange
import com.tjclp.xl.cells.{Cell, CellValue, FormulaKind}
import com.tjclp.xl.sheets.Sheet
import scala.annotation.{nowarn, tailrec}

/**
 * Dependency graph for formula cells.
 *
 * Represents the dependency relationships between cells containing formulas. Each node is a cell
 * reference (ARef), and each directed edge A → B means cell A depends on cell B (A uses B's value).
 *
 * Supports:
 *   - Cycle detection using Tarjan's strongly connected components algorithm
 *   - Topological sorting using Kahn's algorithm for correct evaluation order
 *   - Precedent/dependent queries for impact analysis
 *
 * Design principles:
 *   - Pure functional (no mutations, all operations return new data)
 *   - Total error handling (cycles reported via Either)
 *   - O(V + E) complexity for graph algorithms
 *   - O(1) lookups for precedent/dependent queries (Map-based adjacency lists)
 *
 * Example:
 * {{{
 * // Sheet with formulas: A1="=B1+C1", B1="=10", C1="=20"
 * val graph = DependencyGraph.fromSheet(sheet)
 * graph.precedents(ref"A1") // Set(B1, C1)
 * graph.dependents(ref"B1") // Set(A1)
 *
 * // Detect cycles
 * DependencyGraph.detectCycles(graph) // Right(()) - no cycles
 *
 * // Get evaluation order
 * DependencyGraph.topologicalSort(graph) // Right(List(B1, C1, A1))
 * }}}
 */
final case class DependencyGraph(
  // Forward edges: ref → cells this ref depends on
  dependencies: Map[ARef, Set[ARef]],
  // Reverse edges: ref → cells that depend on this ref
  dependents: Map[ARef, Set[ARef]]
)

object DependencyGraph:
  private def depsFromArgValues(
    values: List[ArgValue],
    exprDeps: TExpr[?] => Set[ARef],
    rangeDeps: TExpr.RangeLocation => Set[ARef],
    cellRangeDeps: CellRange => Set[ARef]
  ): Set[ARef] =
    values.foldLeft(Set.empty[ARef]) { (acc, value) =>
      val next = value match
        case ArgValue.Expr(expr) => exprDeps(expr)
        case ArgValue.Range(range) => rangeDeps(range)
        case ArgValue.Cells(range) => cellRangeDeps(range)
      // Preserve a memoized range Set when it is the first argument. Besides avoiding a needless
      // rebuild, this lets every one-range aggregate in a graph share the exact cached expansion.
      if acc.isEmpty then next else if next.isEmpty then acc else acc ++ next
    }

  private def boundedCells(range: CellRange, bounds: Option[CellRange]): Set[ARef] =
    bounds match
      case Some(b) => range.intersect(b).map(_.cells.toSet).getOrElse(Set.empty)
      case None => range.cells.toSet

  /**
   * Build dependency graph from Sheet.
   *
   * Iterates through all cells, extracts references from Formula cells, and constructs the
   * dependency graph. Non-formula cells (constants, text, etc.) are ignored.
   *
   * @param sheet
   *   The sheet to analyze
   * @return
   *   Dependency graph with nodes for all formula cells and edges for all references
   *
   * Example:
   * {{{
   * val sheet = Sheet.empty
   *   .put(ref"A1", CellValue.Formula("=B1+C1"))
   *   .put(ref"B1", CellValue.Number(10))
   *   .put(ref"C1", CellValue.Formula("=D1*2"))
   *
   * val graph = DependencyGraph.fromSheet(sheet)
   * // graph.dependencies = Map(A1 -> Set(B1, C1), C1 -> Set(D1))
   * // graph.dependents = Map(B1 -> Set(A1), C1 -> Set(A1), D1 -> Set(C1))
   * }}}
   */
  def fromSheet(sheet: Sheet): DependencyGraph =
    // Get bounds once for all extractions - constrains full column/row ranges
    val bounds = sheet.usedRange
    // GH-522: formulas commonly reuse the same driver range. Expand each geometric range once for
    // the whole sheet build (anchors affect dragging, not dependency membership).
    val rangeCache =
      scala.collection.mutable.HashMap.empty[(ARef, ARef), Set[ARef]]

    val formulaCells = sheet.cells.flatMap { case (ref, cell) =>
      cell.value match
        // GH-430: dataTable-kind cells never enter the graph — their TABLE(...) display text is
        // not evaluable (guaranteed-failing parse) and their cache is pinned, so they are pure
        // value sources, not computation nodes.
        case CellValue.Formula(_, _, _: FormulaKind.DataTable) => None
        case CellValue.Formula(expression, _, _) => Some(ref -> expression)
        case _ => None
    }

    // Build forward edges (dependencies) - use bounded extraction to avoid 1M+ cells
    val dependencies = formulaCells.map { case (ref, formulaStr) =>
      val deps = FormulaParser.parse(formulaStr) match
        case scala.util.Right(expr) =>
          extractDependenciesBoundedCached(expr, bounds, rangeCache)
        case scala.util.Left(_) => Set.empty[ARef] // Parse error: no dependencies
      ref -> deps
    }.toMap

    // Build reverse edges with one mutable set per precedent, then freeze once. Growing an
    // immutable Set for every edge recreates the N×range-size allocation this build avoids above.
    val reverse = scala.collection.mutable.HashMap.empty[
      ARef,
      scala.collection.mutable.LinkedHashSet[ARef]
    ]
    dependencies.foreach { case (ref, deps) =>
      deps.foreach { dep =>
        reverse.getOrElseUpdate(dep, scala.collection.mutable.LinkedHashSet.empty) += ref
      }
    }
    val dependents = reverse.iterator.map((ref, refs) => ref -> refs.toSet).toMap

    DependencyGraph(dependencies, dependents)

  /**
   * Check if an expression contains any cell references.
   *
   * GH-197: This is a structural check that doesn't enumerate cells in ranges. Use this for quick
   * boolean checks (e.g., "does this formula need a sheet?") instead of extractDependencies which
   * would enumerate 1M+ cells for full-column ranges.
   *
   * @param expr
   *   The expression to analyze
   * @return
   *   true if the expression contains any Ref, PolyRef, RangeRef, or cross-sheet references
   */
  @nowarn("msg=Unreachable case")
  def containsCellReferences[A](expr: TExpr[A]): Boolean =
    expr match
      // Cell references
      case TExpr.Ref(_, _, _) => true
      case TExpr.PolyRef(_, _) => true
      case TExpr.RangeRef(_) => true
      case TExpr.SheetRef(_, _, _, _) => true
      case TExpr.SheetPolyRef(_, _, _) => true
      case TExpr.SheetRange(_, _) => true
      // GH-353: external-workbook refs ARE cell references (they just live outside the workbook)
      case TExpr.ExternalRef(_, _, _, _) => true
      case TExpr.ExternalRange(_, _, _) => true
      case TExpr.Aggregate(_, _) => true

      // Function calls - check arguments
      case call: TExpr.Call[?] =>
        call.spec.argSpec
          .toValues(call.args)
          .exists {
            case ArgValue.Expr(e) => containsCellReferences(e)
            case ArgValue.Range(_) => true
            case ArgValue.Cells(_) => true
          }

      // Binary operators - check both sides
      case TExpr.Add(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Sub(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Mul(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Div(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Pow(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Concat(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Eq(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Neq(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Lt(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Lte(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Gt(l, r) => containsCellReferences(l) || containsCellReferences(r)
      case TExpr.Gte(l, r) => containsCellReferences(l) || containsCellReferences(r)

      // Unary operators
      case TExpr.ToInt(e) => containsCellReferences(e)
      case TExpr.UnaryPlus(e) => containsCellReferences(e)
      case TExpr.Percent(e) => containsCellReferences(e)
      case TExpr.DateToSerial(e) => containsCellReferences(e)
      case TExpr.DateTimeToSerial(e) => containsCellReferences(e)

      // GH-193: LET — check binding values and the body; BindingRef is a name, not a cell
      case TExpr.Let(bindings, body) =>
        bindings.exists((_, value) => containsCellReferences(value)) ||
        containsCellReferences(body)
      case TExpr.BindingRef(_) => false
      case TExpr.CoercedBindingRef(_, _) => false

      // GH-384: a defined name resolves to cells (or a formula over cells) at evaluation time —
      // it behaves like a reference for "does this formula read data?" checks
      case TExpr.NameRef(_) => true
      // GH-394: sheet-qualified names likewise
      case TExpr.SheetNameRef(_, _) => true

      // GH-306: runtime coercion wrapper — transparent for analysis
      case TExpr.Coerced(inner, _) => containsCellReferences(inner)

      // Literals and constants
      case TExpr.Lit(_) => false

  /**
   * GH-274: Check whether an expression contains a dynamic-reference function call.
   *
   * A dynamic reference is a `TExpr.Call` whose spec is flagged `FunctionFlags.dynamicDeps` (e.g.
   * INDIRECT): its data dependencies are decided by evaluated text, so the static graph cannot see
   * them. The walk recurses into Call arguments; `ArgValue.Range`/`Cells` arguments cannot contain
   * calls, and reference/literal leaves are never dynamic.
   */
  @nowarn("msg=Unreachable case")
  private def containsDynamicReferenceResolved[A](
    expr: TExpr[A],
    resolveName: (String, Option[SheetName]) => Boolean
  ): Boolean =
    expr match
      case call: TExpr.Call[?] =>
        call.spec.flags.dynamicDeps || call.spec.argSpec
          .toValues(call.args)
          .exists {
            case ArgValue.Expr(e) => containsDynamicReferenceResolved(e, resolveName)
            case ArgValue.Range(TExpr.RangeLocation.Name(name, scope)) =>
              resolveName(name, scope)
            case ArgValue.Range(_) => false
            case ArgValue.Cells(_) => false
          }
      case TExpr.Add(l, r) =>
        containsDynamicReferenceResolved(l, resolveName) ||
        containsDynamicReferenceResolved(r, resolveName)
      case TExpr.Sub(l, r) =>
        containsDynamicReferenceResolved(l, resolveName) ||
        containsDynamicReferenceResolved(r, resolveName)
      case TExpr.Mul(l, r) =>
        containsDynamicReferenceResolved(l, resolveName) ||
        containsDynamicReferenceResolved(r, resolveName)
      case TExpr.Div(l, r) =>
        containsDynamicReferenceResolved(l, resolveName) ||
        containsDynamicReferenceResolved(r, resolveName)
      case TExpr.Pow(l, r) =>
        containsDynamicReferenceResolved(l, resolveName) ||
        containsDynamicReferenceResolved(r, resolveName)
      case TExpr.Concat(l, r) =>
        containsDynamicReferenceResolved(l, resolveName) ||
        containsDynamicReferenceResolved(r, resolveName)
      case TExpr.Eq(l, r) =>
        containsDynamicReferenceResolved(l, resolveName) ||
        containsDynamicReferenceResolved(r, resolveName)
      case TExpr.Neq(l, r) =>
        containsDynamicReferenceResolved(l, resolveName) ||
        containsDynamicReferenceResolved(r, resolveName)
      case TExpr.Lt(l, r) =>
        containsDynamicReferenceResolved(l, resolveName) ||
        containsDynamicReferenceResolved(r, resolveName)
      case TExpr.Lte(l, r) =>
        containsDynamicReferenceResolved(l, resolveName) ||
        containsDynamicReferenceResolved(r, resolveName)
      case TExpr.Gt(l, r) =>
        containsDynamicReferenceResolved(l, resolveName) ||
        containsDynamicReferenceResolved(r, resolveName)
      case TExpr.Gte(l, r) =>
        containsDynamicReferenceResolved(l, resolveName) ||
        containsDynamicReferenceResolved(r, resolveName)
      case TExpr.ToInt(e) => containsDynamicReferenceResolved(e, resolveName)
      case TExpr.UnaryPlus(e) => containsDynamicReferenceResolved(e, resolveName)
      case TExpr.Percent(e) => containsDynamicReferenceResolved(e, resolveName)
      case TExpr.DateToSerial(e) => containsDynamicReferenceResolved(e, resolveName)
      case TExpr.DateTimeToSerial(e) => containsDynamicReferenceResolved(e, resolveName)
      // GH-306: runtime coercion wrapper — a coerced INDIRECT/OFFSET is still dynamic
      case TExpr.Coerced(inner, _) => containsDynamicReferenceResolved(inner, resolveName)
      // GH-193: LET — binding values and the body may carry dynamic calls
      case TExpr.Let(bindings, body) =>
        bindings.exists((_, value) => containsDynamicReferenceResolved(value, resolveName)) ||
        containsDynamicReferenceResolved(body, resolveName)
      case TExpr.NameRef(name) => resolveName(name, None)
      case TExpr.SheetNameRef(sheet, name) => resolveName(name, Some(sheet))
      case TExpr.Aggregate(_, TExpr.RangeLocation.Name(name, scope)) => resolveName(name, scope)
      case _ => false

  def containsDynamicReference[A](expr: TExpr[A]): Boolean =
    containsDynamicReferenceResolved(expr, (_, _) => false)

  /**
   * GH-274: Find all formula cells in the sheet bearing dynamic references (INDIRECT et al).
   *
   * A case-insensitive substring pre-filter over the registry's dynamicDeps-flagged names runs
   * before any parsing, so sheets without dynamic functions pay no parse cost. Formulas that fail
   * to parse are not dynamic (matching `fromSheet`, where they contribute no edges).
   */
  def dynamicCells(sheet: Sheet): Set[ARef] =
    val names = FunctionRegistry.dynamicFunctionNames
    if names.isEmpty then Set.empty
    else
      sheet.cells.iterator.flatMap { case (ref, cell) =>
        cell.value match
          case CellValue.Formula(expression, _, _)
              if names.exists(n => expression.toUpperCase(java.util.Locale.ROOT).contains(n)) =>
            FormulaParser.parse(expression) match
              case scala.util.Right(expr) if containsDynamicReference(expr) => Some(ref)
              case _ => None
          case _ => None
      }.toSet

  /**
   * GH-520: workbook-aware dynamic classification, including parseable defined-name chains.
   *
   * A name is resolved with the same case-insensitive, sheet-scoped-shadowing rules as evaluation.
   * The lookup sheet and the formula's ambient sheet both belong to the memo key: for
   * `Other!GlobalName`, lookup occurs as seen from `Other`, while an unqualified name inside a
   * workbook-scoped definition still resolves from the original formula's sheet. Name cycles stop
   * cleanly, matching dependency extraction's total posture.
   *
   * Ordinary names do not make their users dynamic. Definitions are parsed and memoized first; only
   * names whose parseable chains actually reach a dynamic function join the cheap substring
   * pre-filter used for cell formulas. This matters for financial models with tens of thousands of
   * formulas referring to a static scenario name such as `Case`.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def dynamicCells(workbook: Workbook): Set[QualifiedRef] =
    type NameKey = (SheetName, SheetName, String)

    val dynamicFunctions = FunctionRegistry.dynamicFunctionNames
    val memo = scala.collection.mutable.HashMap.empty[NameKey, Boolean]

    def expressionIsDynamic(
      expr: TExpr[?],
      currentSheet: SheetName,
      visiting: Set[NameKey]
    ): Boolean =
      containsDynamicReferenceResolved(
        expr,
        (name, scope) =>
          nameIsDynamic(
            name,
            lookupFrom = scope.getOrElse(currentSheet),
            fallbackSheet = currentSheet,
            visiting = visiting
          )
      )

    def nameIsDynamic(
      name: String,
      lookupFrom: SheetName,
      fallbackSheet: SheetName,
      visiting: Set[NameKey]
    ): Boolean =
      val key = (lookupFrom, fallbackSheet, name.toUpperCase(java.util.Locale.ROOT))
      memo.get(key) match
        case Some(dynamic) => dynamic
        case None if visiting.contains(key) => false
        case None =>
          val dynamic =
            (for
              definedName <- Evaluator.lookupDefinedName(workbook, lookupFrom, name)
              target <- FormulaParser.parse(definedName.formula).toOption
            yield
              val definingSheet =
                Evaluator
                  .definedNameScope(workbook, definedName)
                  .map(_.name)
                  .getOrElse(fallbackSheet)
              expressionIsDynamic(target, definingSheet, visiting + key)
            ).getOrElse(false)
          memo(key) = dynamic
          dynamic

    val definedNameIds = workbook.metadata.definedNames.iterator.map(_.name).toSet
    val dynamicNameTokens = workbook.sheets.iterator.flatMap { sheet =>
      definedNameIds.iterator.collect {
        case name if nameIsDynamic(name, sheet.name, sheet.name, Set.empty) =>
          name.toUpperCase(java.util.Locale.ROOT)
      }
    }.toSet
    val candidateTokens = dynamicFunctions ++ dynamicNameTokens

    if candidateTokens.isEmpty then Set.empty
    else
      workbook.sheets.iterator.flatMap { sheet =>
        sheet.cells.iterator.flatMap { case (ref, cell) =>
          cell.value match
            case CellValue.Formula(expression, _, _)
                if candidateTokens.exists(
                  expression.toUpperCase(java.util.Locale.ROOT).contains
                ) =>
              FormulaParser.parse(expression) match
                case scala.util.Right(expr) if expressionIsDynamic(expr, sheet.name, Set.empty) =>
                  Some(QualifiedRef(sheet.name, ref))
                case _ => None
            case _ => None
        }
      }.toSet

  /**
   * GH-274: A dynamic cell set closed under "depends on me": the cells themselves plus every
   * transitive static dependent. `transitiveDependents` excludes its seeds, so the union is
   * required. This is the bucket `deferDynamic` moves to the end of evaluation order.
   */
  def dynamicClosure(graph: DependencyGraph, dynamic: Set[ARef]): Set[ARef] =
    dynamic ++ transitiveDependents(graph, dynamic)

  /**
   * GH-274: Stable partition of a topological order — non-closure cells first, closure cells last,
   * relative order preserved within each part.
   *
   * Lemma (order preservation): if `order` is a valid topological order of the graph and `closure`
   * is dependent-closed (contains every static dependent of each of its members), the result is
   * also a valid topological order. For any edge u-depends-on-v: if u and v are in the same part,
   * their relative order is preserved (stable partition); if u is deferred and v is not, v sits in
   * the front part, before u; v deferred with u not deferred is impossible — u depends on v, so u
   * is a dependent of v and dependent-closure would have pulled u into the closure. Property-tested
   * in DependencyGraphSpec.
   */
  def deferDynamic(order: List[ARef], closure: Set[ARef]): List[ARef] =
    if closure.isEmpty then order
    else order.filterNot(closure.contains) ++ order.filter(closure.contains)

  /**
   * Check if an expression contains any **unqualified** cell references.
   *
   * GH-210: Fully-qualified references (SheetRef, SheetPolyRef, SheetRange, CrossSheet aggregates)
   * already name their target sheet, so the formula doesn't require a `-s` flag. Only unqualified
   * refs (Ref, PolyRef, RangeRef, Local aggregates) need an ambient sheet context.
   *
   * @param expr
   *   The expression to analyze
   * @return
   *   true if the expression contains any unqualified cell reference
   */
  @nowarn("msg=Unreachable case")
  def containsUnqualifiedCellReferences[A](expr: TExpr[A]): Boolean =
    expr match
      // Unqualified cell references - need ambient sheet
      case TExpr.Ref(_, _, _) => true
      case TExpr.PolyRef(_, _) => true
      case TExpr.RangeRef(_) => true
      case TExpr.Aggregate(_, TExpr.RangeLocation.Local(_)) => true

      // Qualified cell references - sheet already specified
      case TExpr.SheetRef(_, _, _, _) => false
      case TExpr.SheetPolyRef(_, _, _) => false
      case TExpr.SheetRange(_, _) => false
      // GH-353: external refs are fully qualified (workbook + sheet) — no ambient sheet needed
      case TExpr.ExternalRef(_, _, _, _) => false
      case TExpr.ExternalRange(_, _, _) => false
      case TExpr.Aggregate(_, TExpr.RangeLocation.CrossSheet(_, _)) => false
      case TExpr.Aggregate(_, TExpr.RangeLocation.External(_, _, _)) => false
      // GH-394: an unqualified name's lookup depends on the ambient sheet (sheet-scoped names
      // shadow workbook-scoped ones); a sheet-qualified name carries its own context
      case TExpr.Aggregate(_, TExpr.RangeLocation.Name(_, scope)) => scope.isEmpty

      // Function calls - check arguments
      case call: TExpr.Call[?] =>
        call.spec.argSpec
          .toValues(call.args)
          .exists {
            case ArgValue.Expr(e) => containsUnqualifiedCellReferences(e)
            case ArgValue.Range(loc) =>
              loc match
                case TExpr.RangeLocation.Local(_) => true
                case TExpr.RangeLocation.CrossSheet(_, _) => false
                // GH-353: fully qualified (workbook + sheet) — no ambient sheet needed
                case TExpr.RangeLocation.External(_, _, _) => false
                // GH-394: unqualified name lookup depends on the ambient sheet
                case TExpr.RangeLocation.Name(_, scope) => scope.isEmpty
            case ArgValue.Cells(_) => true
          }

      // Binary operators - check both sides
      case TExpr.Add(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Sub(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Mul(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Div(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Pow(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Concat(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Eq(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Neq(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Lt(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Lte(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Gt(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)
      case TExpr.Gte(l, r) =>
        containsUnqualifiedCellReferences(l) || containsUnqualifiedCellReferences(r)

      // Unary operators
      case TExpr.ToInt(e) => containsUnqualifiedCellReferences(e)
      case TExpr.UnaryPlus(e) => containsUnqualifiedCellReferences(e)
      case TExpr.Percent(e) => containsUnqualifiedCellReferences(e)
      case TExpr.DateToSerial(e) => containsUnqualifiedCellReferences(e)
      case TExpr.DateTimeToSerial(e) => containsUnqualifiedCellReferences(e)

      // GH-193: LET — check binding values and the body; BindingRef is a name, not a cell
      case TExpr.Let(bindings, body) =>
        bindings.exists((_, value) => containsUnqualifiedCellReferences(value)) ||
        containsUnqualifiedCellReferences(body)
      case TExpr.BindingRef(_) => false
      case TExpr.CoercedBindingRef(_, _) => false

      // GH-384: name resolution depends on the ambient sheet (sheet-scoped names SHADOW
      // workbook-scoped ones), so a name-bearing formula needs a sheet context
      case TExpr.NameRef(_) => true
      // GH-394: a sheet-qualified name carries its own lookup context
      case TExpr.SheetNameRef(_, _) => false

      // GH-306: runtime coercion wrapper — transparent for analysis
      case TExpr.Coerced(inner, _) => containsUnqualifiedCellReferences(inner)

      // Literals and constants
      case TExpr.Lit(_) => false

  /**
   * Extract all cell references from TExpr.
   *
   * Recursively traverses the expression AST and collects all cell references, including:
   *   - Single cell references (Ref)
   *   - Range references (RangeRef) expanded to all cells in range
   *
   * @param expr
   *   The expression to analyze
   * @return
   *   Set of all cell references used in the expression
   *
   * Example:
   * {{{
   * val expr = TExpr.Add(TExpr.Ref(ref"A1", ...), TExpr.Lit(10))
   * extractDependencies(expr) // Set(A1)
   *
   * val sumExpr = TExpr.sum(CellRange.parse("B1:B10").toOption.get)
   * extractDependencies(sumExpr) // Set(B1, B2, ..., B10)
   * }}}
   */
  // nowarn: Compiler incorrectly reports PolyRef as unreachable, but tests confirm it IS reached at runtime
  @nowarn("msg=Unreachable case")
  def extractDependencies[A](expr: TExpr[A]): Set[ARef] =
    expr match
      // Single cell reference
      case TExpr.Ref(at, _, _) => Set(at)

      // Polymorphic reference (type resolved at evaluation time)
      case TExpr.PolyRef(at, _) => Set(at)

      // Cross-sheet references return Set.empty in same-sheet dependency extraction.
      // This is intentional: extractDependencies builds intra-sheet graphs only.
      // For workbook-level dependency tracking, use extractQualifiedDependencies + fromWorkbook.
      case TExpr.SheetRef(_, _, _, _) => Set.empty
      case TExpr.SheetPolyRef(_, _, _) => Set.empty
      case TExpr.SheetRange(_, _) => Set.empty
      // GH-353: external-workbook refs target cells OUTSIDE the workbook — no edges ever
      case TExpr.ExternalRef(_, _, _, _) => Set.empty
      case TExpr.ExternalRange(_, _, _) => Set.empty
      case TExpr.RangeRef(range) =>
        range.cells.toSet

      case call: TExpr.Call[?] =>
        depsFromArgValues(
          call.spec.argSpec.toValues(call.args),
          expr => extractDependencies(expr),
          _.localCells,
          _.cells.toSet
        )

      // Recursive cases (binary operators)
      case TExpr.Add(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Sub(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Mul(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Div(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Pow(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Concat(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Eq(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Neq(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Lt(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Lte(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Gt(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Gte(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.ToInt(expr) =>
        extractDependencies(expr) // Type conversion - extract from wrapped expr
      // GH-374: unary plus is transparent — dependencies under it feed recalc edges
      case TExpr.UnaryPlus(expr) => extractDependencies(expr)
      // GH-355: postfix percent — dependencies under it feed recalc edges
      case TExpr.Percent(expr) => extractDependencies(expr)
      case TExpr.Aggregate(_, location) => location.localCells

      // GH-193: LET — union of binding-value and body dependencies; BindingRef has none
      case TExpr.Let(bindings, body) =>
        bindings.foldLeft(extractDependencies(body)) { case (acc, (_, value)) =>
          acc ++ extractDependencies(value)
        }
      case TExpr.BindingRef(_) => Set.empty
      case TExpr.CoercedBindingRef(_, _) => Set.empty
      // GH-384: name targets live behind workbook metadata this sheet-level walk cannot see —
      // no intra-sheet edges (consistent with cross-sheet refs; the qualified extractor resolves).
      // GH-394: sheet-qualified names likewise — NOTE this means an edit to a name's TARGET does
      // not trigger sheet-level targeted recalc of name users; the workbook-level qualified
      // extractor below carries the real edges.
      case TExpr.NameRef(_) => Set.empty
      case TExpr.SheetNameRef(_, _) => Set.empty

      // GH-306: runtime coercion wrapper — transparent for analysis
      case TExpr.Coerced(inner, _) => extractDependencies(inner)

      // Literals and nullary functions (no dependencies)
      case TExpr.Lit(_) => Set.empty
      case TExpr.DateToSerial(dateExpr) => extractDependencies(dateExpr)
      case TExpr.DateTimeToSerial(dtExpr) => extractDependencies(dtExpr)

  /**
   * Extract all cell references from TExpr, bounded by the sheet's used range.
   *
   * This optimized version constrains full column/row references (like A:A or 1:1) to the
   * intersection with bounds, avoiding iteration over 1M+ cells. Use this when building dependency
   * graphs from sheets.
   *
   * @param expr
   *   The expression to analyze
   * @param bounds
   *   Optional bounding range (typically sheet.usedRange) to constrain full ranges
   * @return
   *   Set of all cell references used in the expression, bounded by the given range
   */
  def extractDependenciesBounded[A](expr: TExpr[A], bounds: Option[CellRange]): Set[ARef] =
    extractDependenciesBoundedCached(
      expr,
      bounds,
      scala.collection.mutable.HashMap.empty[(ARef, ARef), Set[ARef]]
    )

  @nowarn("msg=Unreachable case")
  private def extractDependenciesBoundedCached[A](
    expr: TExpr[A],
    bounds: Option[CellRange],
    rangeCache: scala.collection.mutable.HashMap[(ARef, ARef), Set[ARef]]
  ): Set[ARef] =
    // Helper to bound a CellRange
    def boundRange(range: CellRange): Set[ARef] =
      rangeCache.getOrElseUpdate(
        (range.start, range.end),
        bounds match
          case Some(b) => range.intersect(b).map(_.cells.toSet).getOrElse(Set.empty)
          case None => range.cells.toSet
      )

    def recurse[B](child: TExpr[B]): Set[ARef] =
      extractDependenciesBoundedCached(child, bounds, rangeCache)

    def localCells(location: TExpr.RangeLocation): Set[ARef] = location match
      case TExpr.RangeLocation.Local(range) => boundRange(range)
      case _ => Set.empty

    expr match
      // Single cell reference
      case TExpr.Ref(at, _, _) => Set(at)

      // Polymorphic reference (type resolved at evaluation time)
      case TExpr.PolyRef(at, _) => Set(at)

      // Cross-sheet references return Set.empty in same-sheet dependency extraction.
      case TExpr.SheetRef(_, _, _, _) => Set.empty
      case TExpr.SheetPolyRef(_, _, _) => Set.empty
      case TExpr.SheetRange(_, _) => Set.empty
      // GH-353: external-workbook refs target cells OUTSIDE the workbook — no edges ever
      case TExpr.ExternalRef(_, _, _, _) => Set.empty
      case TExpr.ExternalRange(_, _, _) => Set.empty
      case TExpr.RangeRef(range) => boundRange(range)

      // Recursive cases (binary operators)
      case TExpr.Add(l, r) =>
        recurse(l) ++ recurse(r)
      case TExpr.Sub(l, r) =>
        recurse(l) ++ recurse(r)
      case TExpr.Mul(l, r) =>
        recurse(l) ++ recurse(r)
      case TExpr.Div(l, r) =>
        recurse(l) ++ recurse(r)
      case TExpr.Pow(l, r) =>
        recurse(l) ++ recurse(r)
      case TExpr.Concat(l, r) =>
        recurse(l) ++ recurse(r)
      case TExpr.Eq(l, r) =>
        recurse(l) ++ recurse(r)
      case TExpr.Neq(l, r) =>
        recurse(l) ++ recurse(r)
      case TExpr.Lt(l, r) =>
        recurse(l) ++ recurse(r)
      case TExpr.Lte(l, r) =>
        recurse(l) ++ recurse(r)
      case TExpr.Gt(l, r) =>
        recurse(l) ++ recurse(r)
      case TExpr.Gte(l, r) =>
        recurse(l) ++ recurse(r)
      case TExpr.ToInt(expr) => recurse(expr)
      case TExpr.UnaryPlus(expr) => recurse(expr)
      case TExpr.Percent(expr) => recurse(expr)
      case TExpr.Aggregate(_, location) => localCells(location)

      case call: TExpr.Call[?] =>
        depsFromArgValues(
          call.spec.argSpec.toValues(call.args),
          recurse,
          localCells,
          boundRange
        )

      // GH-193: LET — union of binding-value and body dependencies; BindingRef has none
      case TExpr.Let(bindings, body) =>
        bindings.foldLeft(recurse(body)) { case (acc, (_, value)) =>
          acc ++ recurse(value)
        }
      case TExpr.BindingRef(_) => Set.empty
      case TExpr.CoercedBindingRef(_, _) => Set.empty
      // GH-384: name targets live behind workbook metadata this sheet-level walk cannot see —
      // no intra-sheet edges (consistent with cross-sheet refs; the qualified extractor resolves).
      // GH-394: sheet-qualified names likewise — NOTE this means an edit to a name's TARGET does
      // not trigger sheet-level targeted recalc of name users; the workbook-level qualified
      // extractor below carries the real edges.
      case TExpr.NameRef(_) => Set.empty
      case TExpr.SheetNameRef(_, _) => Set.empty

      // GH-306: runtime coercion wrapper — transparent for analysis
      case TExpr.Coerced(inner, _) => recurse(inner)

      // Literals and nullary functions (no dependencies)
      case TExpr.Lit(_) => Set.empty
      case TExpr.DateToSerial(dateExpr) => recurse(dateExpr)
      case TExpr.DateTimeToSerial(dtExpr) => recurse(dtExpr)

  /**
   * Get cells this cell depends on (precedents).
   *
   * Returns the set of cells whose values are used in calculating this cell's value. If the cell
   * has no formula or is not in the graph, returns empty set.
   *
   * @param graph
   *   The dependency graph
   * @param ref
   *   The cell reference to query
   * @return
   *   Set of cells this cell depends on (may be empty)
   *
   * Example:
   * {{{
   * // A1 = "=B1+C1"
   * precedents(graph, ref"A1") // Set(B1, C1)
   * precedents(graph, ref"B1") // Set() - B1 is a constant
   * }}}
   */
  def precedents(graph: DependencyGraph, ref: ARef): Set[ARef] =
    graph.dependencies.getOrElse(ref, Set.empty)

  /**
   * Get cells that depend on this cell (dependents).
   *
   * Returns the set of cells that use this cell's value in their calculations. If no cells depend
   * on this cell, returns empty set.
   *
   * @param graph
   *   The dependency graph
   * @param ref
   *   The cell reference to query
   * @return
   *   Set of cells that depend on this cell (may be empty)
   *
   * Example:
   * {{{
   * // A1 = "=B1+C1", D1 = "=B1*2"
   * dependents(graph, ref"B1") // Set(A1, D1)
   * dependents(graph, ref"A1") // Set() - nothing depends on A1
   * }}}
   */
  def dependents(graph: DependencyGraph, ref: ARef): Set[ARef] =
    graph.dependents.getOrElse(ref, Set.empty)

  /**
   * Compute transitive dependencies for a set of cells.
   *
   * Given a set of starting cells, returns all cells that are directly or transitively depended
   * upon. This is useful for targeted evaluation - to evaluate only formulas in a range, we need to
   * also evaluate all cells they depend on (recursively).
   *
   * @param graph
   *   The dependency graph to traverse
   * @param refs
   *   The starting cell references
   * @return
   *   Set of all cells that the starting cells depend on (transitively)
   *
   * Example:
   * {{{
   * // A1="=B1+C1", B1="=D1", C1="=10", D1="=20"
   * // transitiveDependencies(graph, Set(A1))
   * // Returns: Set(B1, C1, D1) - all cells A1 depends on directly or indirectly
   * }}}
   */
  @scala.annotation.tailrec
  def transitiveDependencies(
    graph: DependencyGraph,
    refs: Set[ARef],
    visited: Set[ARef] = Set.empty
  ): Set[ARef] =
    val toVisit = refs -- visited
    if toVisit.isEmpty then visited
    else
      // Get direct dependencies of all cells in toVisit
      val directDeps = toVisit.flatMap(ref => graph.dependencies.getOrElse(ref, Set.empty))
      // Recurse with direct deps as new frontier
      transitiveDependencies(graph, directDeps, visited ++ toVisit)

  /**
   * Compute transitive dependents (reverse of transitiveDependencies).
   *
   * Returns all cells that depend on the given cells, directly or transitively. Used for eager
   * recalculation - when cell X changes, find all formulas to recalculate.
   *
   * @param graph
   *   The dependency graph to traverse
   * @param refs
   *   The starting cell references (modified cells)
   * @return
   *   Set of all cells that depend on the starting cells (transitively), excluding the starting
   *   refs themselves
   *
   * Example:
   * {{{
   * // A1="=10", B1="=A1*2", C1="=B1+5"
   * // transitiveDependents(graph, Set(A1))
   * // Returns: Set(B1, C1) - all cells affected when A1 changes
   * }}}
   */
  def transitiveDependents(
    graph: DependencyGraph,
    refs: Set[ARef]
  ): Set[ARef] =
    transitiveDependentsOf(graph.dependents, refs)

  /**
   * Reverse-reachability core shared by the sheet-level and workbook-level (qualified) variants —
   * generic in the node type (GH-346). Excludes the starting refs from the result.
   */
  @SuppressWarnings(Array("org.wartremover.warts.While"))
  private def transitiveDependentsOf[K](
    dependents: Map[K, Set[K]],
    originalRefs: Set[K]
  ): Set[K] =
    // A persistent-Set frontier looks elegant but is catastrophically expensive on deep graphs:
    // every round rebuilds `frontier -- visited` and `visited ++ frontier`. A 100k linear cone can
    // therefore perform billions of hash probes. Keep mutation invocation-local and freeze once.
    val visited = scala.collection.mutable.HashSet.empty[K]
    val pending = scala.collection.mutable.ArrayDeque.empty[K]
    originalRefs.foreach { ref =>
      visited += ref
      pending.append(ref)
    }
    while pending.nonEmpty do
      val ref = pending.removeHead()
      dependents.getOrElse(ref, Set.empty).foreach { dependent =>
        if visited.add(dependent) then pending.append(dependent)
      }
    visited --= originalRefs
    visited.toSet

  /**
   * Iterative Tarjan SCC engine shared by `cyclicNodes` and its qualified (workbook-level)
   * counterparts — generic in the node type (GH-346).
   *
   * Uses an explicit work stack instead of recursion: `recalculate` promises totality, and a deep
   * linear dependency chain (e.g. 100k sequential formulas) must not throw StackOverflowError.
   * Invokes `onScc` for every completed component in Tarjan pop order (deepest member first,
   * component root last).
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def foreachSccOf[K](dependencies: Map[K, Set[K]])(onScc: List[K] => Unit): Unit =
    final class Frame(val node: K, val successors: Iterator[K])

    var nextIndex = 0
    val indices = scala.collection.mutable.HashMap.empty[K, Int]
    val lowLinks = scala.collection.mutable.HashMap.empty[K, Int]
    val stack = scala.collection.mutable.ArrayBuffer.empty[K]
    val onStack = scala.collection.mutable.HashSet.empty[K]
    val frames = scala.collection.mutable.ArrayBuffer.empty[Frame]
    indices.sizeHint(dependencies.size)
    lowLinks.sizeHint(dependencies.size)
    stack.sizeHint(dependencies.size)
    onStack.sizeHint(dependencies.size)
    frames.sizeHint(dependencies.size)

    def push(v: K): Unit =
      indices(v) = nextIndex
      lowLinks(v) = nextIndex
      nextIndex += 1
      stack += v
      onStack += v
      frames += new Frame(v, dependencies.getOrElse(v, Set.empty).iterator)

    dependencies.keySet.foreach { root =>
      if !indices.contains(root) then
        push(root)
        while frames.nonEmpty do
          val frame = frames(frames.length - 1)
          if frame.successors.hasNext then
            val successor = frame.successors.next()
            if !indices.contains(successor) then push(successor)
            else if onStack.contains(successor) then
              lowLinks(frame.node) = math.min(lowLinks(frame.node), indices(successor))
          else
            frames.remove(frames.length - 1)
            val node = frame.node
            if lowLinks(node) == indices(node) then
              val members = List.newBuilder[K]
              var complete = false
              while !complete do
                val member = stack.remove(stack.length - 1)
                onStack -= member
                members += member
                complete = member == node
              onScc(members.result())
            if frames.nonEmpty then
              val parent = frames(frames.length - 1).node
              lowLinks(parent) = math.min(lowLinks(parent), lowLinks(node))
    }

  private def foreachScc(graph: DependencyGraph)(onScc: List[ARef] => Unit): Unit =
    foreachSccOf(graph.dependencies)(onScc)

  /**
   * Return the first actual directed cycle encountered inside `nodes`.
   *
   * SCC membership alone is insufficient for diagnostics: Tarjan pop order is not necessarily an
   * edge path when a component branches. Likewise, following one successor from a Kahn remainder
   * can begin downstream of a cycle, so blindly appending the first visited node may invent a
   * closing edge. This explicit-stack DFS records the active path and closes a cycle only on a back
   * edge to an active ancestor. Every adjacent pair in the result is therefore an edge in
   * `dependencies`, the first and last nodes are equal, and no interior node is repeated.
   *
   * Roots and successors are ordered by `key`, making the diagnostic a function of the graph's
   * value rather than Map/Set construction order. All mutation is invocation-local and the returned
   * path is immutable.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def firstCycleOf[K, O: Ordering](
    dependencies: Map[K, Set[K]],
    nodes: Set[K],
    key: K => O
  ): Option[List[K]] =
    // State: absent = unseen, 1 = active (gray), 2 = complete (black).
    val states = scala.collection.mutable.HashMap.empty[K, Int]
    val activePath = scala.collection.mutable.ArrayBuffer.empty[K]
    val activeIndices = scala.collection.mutable.HashMap.empty[K, Int]
    states.sizeHint(nodes.size)
    activePath.sizeHint(nodes.size)
    activeIndices.sizeHint(nodes.size)
    var frames = List.empty[(K, List[K])]
    var found = Option.empty[List[K]]

    def push(node: K): Unit =
      states(node) = 1
      activeIndices(node) = activePath.size
      activePath += node
      val successors =
        dependencies
          .getOrElse(node, Set.empty)
          .iterator
          .filter(nodes.contains)
          .toList
          .sortBy(key)
      frames = (node, successors) :: frames

    val roots = nodes.toList.sortBy(key).iterator
    while roots.hasNext && found.isEmpty do
      val root = roots.next()
      if !states.contains(root) then
        push(root)
        while frames.nonEmpty && found.isEmpty do
          frames match
            case (node, successor :: rest) :: tail =>
              frames = (node, rest) :: tail
              states.get(successor) match
                case None => push(successor)
                case Some(1) =>
                  activeIndices.get(successor) match
                    case Some(start) =>
                      found = Some(activePath.iterator.drop(start).toList :+ successor)
                    case None => () // impossible for an active node; remain total if state corrupts
                case Some(_) => ()
            case (node, Nil) :: tail =>
              frames = tail
              states(node) = 2
              activeIndices -= node
              if activePath.nonEmpty then activePath.remove(activePath.size - 1)
            case Nil => () // unreachable: guarded by frames.nonEmpty

    found

  /**
   * Detect circular references using an explicit-stack depth-first search.
   *
   * A circular reference occurs when a cell's formula depends (directly or transitively) on its own
   * value. For example: A1="=B1", B1="=C1", C1="=A1" forms a cycle.
   *
   * The explicit DFS is O(V + E), stops at the first back edge, and is stack-safe on deep
   * dependency chains. Canonical diagnostic ordering adds sorting proportional to the visited roots
   * and successor sets.
   *
   * @param graph
   *   The dependency graph to analyze
   * @return
   *   Left(CircularRef) if a cycle is detected (includes cycle path), Right(()) if acyclic
   *
   * Example:
   * {{{
   * // No cycle: A1="=10", B1="=A1+5"
   * detectCycles(graph) // Right(())
   *
   * // Cycle: A1="=B1", B1="=A1"
   * detectCycles(graph) // Left(EvalError.CircularRef(List(A1, B1, A1)))
   * }}}
   */
  def detectCycles(graph: DependencyGraph): Either[EvalError.CircularRef, Unit] =
    firstCycleOf(
      graph.dependencies,
      graph.dependencies.keySet,
      (ref: ARef) => (ref.row.index0, ref.col.index0)
    ) match
      case Some(cycle) => scala.util.Left(EvalError.CircularRef(cycle))
      case None => scala.util.Right(())

  /**
   * Collect every node that participates in a reference cycle.
   *
   * Unlike `detectCycles` (which fails fast on the first cycle), this identifies the complete set
   * of cells inside strongly connected components of size > 1, plus self-loops. Used by
   * `Workbook.recalculate` to isolate cyclic cells while still evaluating the acyclic remainder.
   *
   * Stack-safe: uses the iterative Tarjan engine shared with the qualified workbook-level path.
   *
   * @return
   *   The set of cycle participants (empty when the graph is acyclic)
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def cyclicNodes(graph: DependencyGraph): Set[ARef] =
    var cyclic = Set.empty[ARef]
    foreachScc(graph) { scc =>
      scc match
        case single :: Nil =>
          if graph.dependencies.get(single).exists(_.contains(single)) then cyclic = cyclic + single
        case multi => cyclic = cyclic ++ multi
    }
    cyclic

  /**
   * Kahn's-algorithm core shared by the sheet-level and workbook-level (qualified) sorts — generic
   * in the node type (GH-346). Only nodes present in `dependencies` (formula cells) are ordered;
   * edges to non-nodes (constants, empty cells) are ignored. Left carries the nodes that could not
   * be ordered (cycle participants and everything downstream of them).
   *
   * GH-518: implemented with LOCAL mutable state behind the pure signature (the same posture as
   * [[foreachSccOf]]). The immutable-List formulation allocated O(V²) cons cells — `acc :+ node`
   * copies the accumulator once per node and `rest ++ newlyZero` copies the pending queue — which
   * on a 50k-formula workbook was 81% of an entire recalculation's allocation. Iteration stays over
   * the same collections in the same order, so the emitted order is unchanged: FIFO queue seeded in
   * `allNodes` iteration order, newly-ready nodes appended in the iteration order of the filtered
   * dependents set. The FIFO discipline is additionally load-bearing for parallel recalculation: it
   * keeps longest-path depth classes contiguous, so folding completed waves preserves the exact
   * sequential error order. A replacement queue must retain that level-monotone property.
   */
  @SuppressWarnings(Array("org.wartremover.warts.While"))
  private def kahnOrder[K](
    dependencies: Map[K, Set[K]],
    dependents: Map[K, Set[K]]
  ): Either[Set[K], List[K]] =
    // All formula cells (only process formulas, not constants)
    val allNodes = dependencies.keySet

    // If no formula cells, early exit
    if allNodes.isEmpty then scala.util.Right(List.empty[K])
    else
      // In-degree = dependencies on other formula cells only, not constants
      val inDegree = new scala.collection.mutable.HashMap[K, Int](
        allNodes.size * 2,
        scala.collection.mutable.HashMap.defaultLoadFactor
      )
      allNodes.foreach { node =>
        inDegree(node) = dependencies.getOrElse(node, Set.empty).count(allNodes.contains)
      }

      // Start with nodes that have in-degree 0 (no dependencies)
      val queue = new scala.collection.mutable.ArrayDeque[K](allNodes.size)
      allNodes.foreach { node =>
        if inDegree(node) == 0 then queue.append(node)
      }

      val ordered = scala.collection.mutable.ListBuffer.empty[K]
      while queue.nonEmpty do
        val node = queue.removeHead()
        ordered += node
        // Test membership inline: materializing `filter(allNodes.contains)` allocated one throwaway
        // Set per processed formula. Iteration stays over the original Set in the same order; only
        // successors absent from the node set are skipped.
        dependents.getOrElse(node, Set.empty).foreach { dep =>
          if allNodes.contains(dep) then
            val remaining = inDegree(dep) - 1
            inDegree(dep) = remaining
            if remaining == 0 then queue.append(dep)
        }

      // If all nodes are processed, graph is acyclic
      if ordered.size == allNodes.size then scala.util.Right(ordered.toList)
      else scala.util.Left(allNodes -- ordered)

  /**
   * Topological sort using Kahn's algorithm.
   *
   * Returns a linear ordering of cells such that for every dependency A → B, cell B appears before
   * cell A in the ordering. This ensures formulas are evaluated in the correct order (dependencies
   * before dependents).
   *
   * Uses Kahn's algorithm which runs in O(V + E) time. The algorithm maintains a queue of nodes
   * with in-degree 0 and processes them in order, removing edges as it goes.
   *
   * @param graph
   *   The dependency graph to sort
   * @return
   *   Left(CircularRef) if a cycle is detected, Right(evaluation order) if acyclic
   *
   * Example:
   * {{{
   * // A1="=B1+C1", B1="=10", C1="=20"
   * topologicalSort(graph) // Right(List(B1, C1, A1))
   *
   * // A1="=B1", B1="=A1" (cycle)
   * topologicalSort(graph) // Left(EvalError.CircularRef(List(A1, B1, A1)))
   * }}}
   */
  def topologicalSort(graph: DependencyGraph): Either[EvalError.CircularRef, List[ARef]] =
    kahnOrder(graph.dependencies, graph.dependents) match
      case scala.util.Right(order) => scala.util.Right(order)
      case scala.util.Left(remainingNodes) =>
        // Kahn's remainder can include nodes merely downstream of a cycle. Reconstruct only from a
        // real back edge, so the diagnostic never invents an edge from the cycle to that tail.
        val cycle =
          firstCycleOf(
            graph.dependencies,
            remainingNodes,
            (ref: ARef) => (ref.row.index0, ref.col.index0)
          )
            .getOrElse(List.empty)
        scala.util.Left(EvalError.CircularRef(cycle))

  // ===== Cross-Sheet Dependency Tracking =====

  /**
   * Cell reference qualified with sheet name for cross-sheet tracking.
   *
   * Used to track dependencies across sheets within a workbook. Each QualifiedRef uniquely
   * identifies a cell in the workbook by combining the sheet name and cell reference.
   *
   * Example:
   * {{{
   * val ref = QualifiedRef(SheetName.unsafe("Sales"), ref"A1")
   * // Represents Sales!A1
   * }}}
   */
  final case class QualifiedRef(sheet: SheetName, ref: ARef):
    // GH-280: quote cell-ref-shaped sheet names so a sheet literally named "A1" renders
    // unambiguously in diagnostics ('A1'!B2, not A1!B2)
    override def toString: String = s"${SheetName.quoteForFormula(sheet.value)}!${ref.toA1}"

  /**
   * Build dependency graph from Workbook (cross-sheet aware).
   *
   * Iterates through all sheets and cells, extracting references from Formula cells and
   * constructing a workbook-level dependency graph. Cross-sheet references are properly tracked
   * using QualifiedRef.
   *
   * @param workbook
   *   The workbook to analyze
   * @return
   *   Dependency graph with QualifiedRef nodes covering all sheets
   *
   * Example:
   * {{{
   * // Sheet1!A1 = "=Sheet2!B1", Sheet2!B1 = 10
   * val graph = DependencyGraph.fromWorkbook(workbook)
   * // graph contains: QualifiedRef(Sheet1, A1) -> Set(QualifiedRef(Sheet2, B1))
   * }}}
   */
  def fromWorkbook(
    workbook: com.tjclp.xl.workbooks.Workbook
  ): Map[QualifiedRef, Set[QualifiedRef]] =
    // GH-522: expand any (sheet, range) once per build — see memoizedCells
    val cellsFor = memoizedCells(unboundedQualifiedCells)
    workbook.sheets.flatMap { sheet =>
      sheet.cells.flatMap { case (cellRef, cell) =>
        cell.value match
          // GH-430: dataTable-kind cells are pinned value sources, never computation nodes
          case CellValue.Formula(_, _, _: FormulaKind.DataTable) => None
          case CellValue.Formula(expression, _, _) =>
            val deps = FormulaParser.parse(expression) match
              case scala.util.Right(expr) =>
                extractQualifiedDependencies(expr, sheet.name, cellsFor, Some(workbook))
              case scala.util.Left(_) => Set.empty[QualifiedRef]
            Some(QualifiedRef(sheet.name, cellRef) -> deps)
          case _ => None
      }
    }.toMap

  /**
   * GH-346: workbook-level dependency graph with ranges bounded by each TARGET sheet's used range —
   * the qualified analog of `fromSheet`'s bounded extraction (a full-column reference must not
   * enumerate 1M+ cells). Single refs stay exact. Ranges over sheets absent from the workbook (or
   * empty ones) contribute no edges: they cannot affect evaluation order, and the referencing
   * formula still errors per cell on evaluation.
   *
   * Returns forward edges (ref → its dependencies) and reverse edges (ref → its dependents), which
   * together drive `qualifiedTopologicalSort`.
   */
  def fromWorkbookBounded(
    workbook: com.tjclp.xl.workbooks.Workbook
  ): (Map[QualifiedRef, Set[QualifiedRef]], Map[QualifiedRef, Set[QualifiedRef]]) =
    val bounds: Map[SheetName, Option[CellRange]] =
      workbook.sheets.map(s => s.name -> s.usedRange).toMap
    // GH-522: expand any (sheet, range) once per build — see memoizedCells
    val cellsFor: (SheetName, CellRange) => Set[QualifiedRef] = memoizedCells { (sheet, range) =>
      bounds.get(sheet) match
        case Some(Some(b)) =>
          range.intersect(b) match
            case Some(clipped) => clipped.cells.map(ref => QualifiedRef(sheet, ref)).toSet
            case None => Set.empty
        case _ => Set.empty // empty or missing sheet: nothing to depend on
    }

    val dependencies = workbook.sheets.flatMap { sheet =>
      sheet.cells.flatMap { case (cellRef, cell) =>
        cell.value match
          // GH-430: dataTable-kind cells are pinned value sources, never computation nodes
          case CellValue.Formula(_, _, _: FormulaKind.DataTable) => None
          case CellValue.Formula(expression, _, _) =>
            val deps = FormulaParser.parse(expression) match
              case scala.util.Right(expr) =>
                extractQualifiedDependencies(expr, sheet.name, cellsFor, Some(workbook))
              case scala.util.Left(_) => Set.empty[QualifiedRef]
            Some(QualifiedRef(sheet.name, cellRef) -> deps)
          case _ => None
      }
    }.toMap

    (dependencies, reverseEdges(dependencies))

  /**
   * Build the formula-to-formula graph needed by whole-workbook recalculation.
   *
   * The public dependency graph records every referenced cell because impact analysis must answer
   * "which formulas read this edited constant?" Whole-book topological evaluation asks a much
   * narrower question: "which FORMULAS must run before this formula?" Materializing constant cells
   * in that graph is pure waste. On a common financial-model shape — 9,900 formulas sharing
   * `SUM($A$1:$A$5000)` over a constant driver column — the value-preserving graph has no range
   * edges at all, while [[fromWorkbookBounded]] retains and repeatedly traverses 49.5 million.
   *
   * Range lookup scans only the target sheet's formula refs and is memoized by range. Single-cell
   * references are filtered against the same formula-node set after name resolution. Consequently
   * this graph is exactly `fromWorkbookBounded` with every non-formula successor removed, without
   * first allocating those successors. Formula evaluation itself is unchanged and still reads the
   * complete ranges from the workbook snapshot.
   *
   * This representation is intentionally internal: dirty-cone callers still need the public graph
   * (or, ultimately, a symbolic point/range index) to discover formulas affected by edited values.
   */
  private[formula] def fromWorkbookFormulaGraph(
    workbook: com.tjclp.xl.workbooks.Workbook
  ): (Map[QualifiedRef, Set[QualifiedRef]], Map[QualifiedRef, Set[QualifiedRef]]) =
    val formulas: Vector[(QualifiedRef, String)] = workbook.sheets.flatMap { sheet =>
      sheet.cells.flatMap { case (cellRef, cell) =>
        cell.value match
          case CellValue.Formula(_, _, _: FormulaKind.DataTable) => None
          case CellValue.Formula(expression, _, _) =>
            Some(QualifiedRef(sheet.name, cellRef) -> expression)
          case _ => None
      }
    }
    val formulaNodes = formulas.iterator.map(_._1).toSet
    val refsBySheet: Map[SheetName, Vector[ARef]] =
      formulas.groupMap(_._1.sheet)(_._1.ref)
    val cellsFor: (SheetName, CellRange) => Set[QualifiedRef] = memoizedCells { (sheet, range) =>
      refsBySheet
        .getOrElse(sheet, Vector.empty)
        .iterator
        .filter(range.contains)
        .map(ref => QualifiedRef(sheet, ref))
        .toSet
    }

    val dependencies = formulas.iterator.map { case (ref, expression) =>
      val deps = FormulaParser.parse(expression) match
        case scala.util.Right(expr) =>
          extractQualifiedDependencies(expr, ref.sheet, cellsFor, Some(workbook))
            .filter(formulaNodes.contains)
        case scala.util.Left(_) => Set.empty[QualifiedRef]
      ref -> deps
    }.toMap

    (dependencies, reverseEdges(dependencies))

  /** One declared range and the formulas that read it. */
  private[xl] final case class QualifiedRangeDependents(
    range: CellRange,
    dependents: Set[QualifiedRef]
  )

  /**
   * Symbolic reverse-dependency index for targeted recalculation.
   *
   * Point references retain the usual `cell -> formulas` map. Range references remain canonical
   * rectangles rather than expanding to one map entry per covered cell. This changes the common
   * N-formulas-by-M-range shape from O(N × M) retained memberships to O(N + unique ranges), and a
   * full-column reference never allocates one million `QualifiedRef`s merely to discover whether an
   * edited point lies inside it.
   */
  private[xl] final case class QualifiedDependencyIndex(
    pointDependents: Map[QualifiedRef, Set[QualifiedRef]],
    rangeDependents: Map[SheetName, Vector[QualifiedRangeDependents]]
  ):
    /**
     * Every formula directly or transitively affected by `originalRefs`, excluding the seeds.
     *
     * Local mutation makes a deep chain linear. Range containment is checked against declared
     * coordinates, not `usedRange`, so clearing the last occupied cell in a range cannot erase the
     * dependency before invalidation runs.
     */
    @SuppressWarnings(Array("org.wartremover.warts.While"))
    def transitiveDependents(originalRefs: Set[QualifiedRef]): Set[QualifiedRef] =
      val visited = scala.collection.mutable.HashSet.empty[QualifiedRef]
      val pending = scala.collection.mutable.ArrayDeque.empty[QualifiedRef]

      def enqueue(ref: QualifiedRef): Unit =
        if visited.add(ref) then pending.append(ref)

      originalRefs.foreach(enqueue)
      while pending.nonEmpty do
        val ref = pending.removeHead()
        pointDependents.getOrElse(ref, Set.empty).foreach(enqueue)
        rangeDependents.getOrElse(ref.sheet, Vector.empty).foreach { entry =>
          if entry.range.contains(ref.ref) then entry.dependents.foreach(enqueue)
        }

      visited --= originalRefs
      visited.toSet

  /**
   * Build a symbolic reverse-dependency index without materializing range cells.
   *
   * The existing qualified extractor already centralizes every AST shape and defined-name rule.
   * Supplying a range callback that records canonical rectangles and returns an empty set separates
   * point dependencies from range dependencies without duplicating that traversal. All mutation is
   * build-local; only immutable maps, vectors, ranges, and sets escape.
   */
  private[xl] def fromWorkbookDependencyIndex(
    workbook: com.tjclp.xl.workbooks.Workbook
  ): QualifiedDependencyIndex =
    val pointBuilders = scala.collection.mutable.HashMap.empty[
      QualifiedRef,
      scala.collection.mutable.HashSet[QualifiedRef]
    ]
    val rangeBuilders = scala.collection.mutable.HashMap.empty[
      (SheetName, ARef, ARef),
      scala.collection.mutable.HashSet[QualifiedRef]
    ]

    workbook.sheets.foreach { sheet =>
      sheet.cells.foreach { case (cellRef, cell) =>
        cell.value match
          case CellValue.Formula(_, _, _: FormulaKind.DataTable) => ()
          case CellValue.Formula(expression, _, _) =>
            FormulaParser.parse(expression) match
              case scala.util.Left(_) => ()
              case scala.util.Right(expr) =>
                val dependent = QualifiedRef(sheet.name, cellRef)
                val ranges = scala.collection.mutable.HashSet.empty[(SheetName, ARef, ARef)]
                val recordRange: (SheetName, CellRange) => Set[QualifiedRef] =
                  (targetSheet, range) =>
                    ranges += ((targetSheet, range.start, range.end))
                    Set.empty
                val points = extractQualifiedDependencies(
                  expr,
                  sheet.name,
                  recordRange,
                  Some(workbook)
                )
                points.foreach { point =>
                  pointBuilders.getOrElseUpdate(
                    point,
                    scala.collection.mutable.HashSet.empty
                  ) += dependent
                }
                ranges.foreach { key =>
                  rangeBuilders.getOrElseUpdate(
                    key,
                    scala.collection.mutable.HashSet.empty
                  ) += dependent
                }
          case _ => ()
      }
    }

    val points = pointBuilders.iterator.map((ref, refs) => ref -> refs.toSet).toMap
    val ranges = rangeBuilders.iterator
      .map { case ((sheet, start, end), refs) =>
        sheet -> QualifiedRangeDependents(CellRange(start, end), refs.toSet)
      }
      .toVector
      .groupMap(_._1)(_._2)
    QualifiedDependencyIndex(points, ranges)

  /**
   * GH-346: every node participating in a reference cycle of the workbook-level graph — the
   * qualified analog of `cyclicNodes`, sharing the iterative (stack-safe) Tarjan engine. Detects
   * cycles that span sheets, which per-sheet analysis cannot see.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def qualifiedCyclicNodes(
    dependencies: Map[QualifiedRef, Set[QualifiedRef]]
  ): Set[QualifiedRef] =
    var cyclic = Set.empty[QualifiedRef]
    foreachSccOf(dependencies) {
      case single :: Nil =>
        if dependencies.get(single).exists(_.contains(single)) then cyclic = cyclic + single
      case multi => cyclic = cyclic ++ multi
    }
    cyclic

  /**
   * GH-492: one node of the SCC condensation of the workbook-level graph.
   *
   * `members` is sorted by (sheet.value, ref.toA1) — the deterministic member order the Jacobi
   * fixpoint needs (values are order-independent under Jacobi, but error vectors, report members
   * and seeded-RNG draw order are not). `cyclic` is true for a multi-node component or a
   * self-looping singleton, matching [[qualifiedCyclicNodes]] exactly.
   *
   * No `derives CanEqual`: [[QualifiedRef]] does not derive it.
   */
  final case class Scc(members: Vector[QualifiedRef], cyclic: Boolean):
    /** Canonical key: the minimum member under (sheet.value, ref.toA1). Unique across nodes. */
    private[xl] def key: (String, String) =
      members.headOption.fold(("", ""))(q => (q.sheet.value, q.ref.toA1))

  /**
   * GH-492: the SCC condensation of the workbook-level graph, in the LEXICOGRAPHICALLY-MINIMUM
   * dependency-first topological order — every precedent component strictly precedes the components
   * that read it, ties broken by [[Scc.key]].
   *
   * This is what lets one iterative recalculation reach the workbook's GLOBAL fixpoint: walking the
   * condensation once (acyclic components evaluate, cyclic components fixpoint) guarantees every
   * precedent is a freshly computed value before anything reads it, so no read can fall back to a
   * loaded cache. `qualifiedCyclicNodes` flattens the same Tarjan run into an undifferentiated
   * `Set` and is kept for the callers that only need the membership test.
   *
   * Laws (property-tested in DependencyGraphSpec):
   *   - partition — `flatMap(_.members).toSet == dependencies.keySet`, no duplicates
   *   - order — for every edge u → v with different components, index(scc(v)) < index(scc(u))
   *   - cyclic-view — `filter(_.cyclic).flatMap(_.members).toSet == qualifiedCyclicNodes(deps)`
   *   - determinism — depends only on the graph's VALUE, never on Map/Set iteration order
   *
   * Runs in O(V + E + C log C). Only nodes present in `dependencies` (formula cells) are ordered:
   * Tarjan also visits constants reached as successors, and those are always singletons with no
   * outgoing edges, so filtering them out can never split a real component.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def qualifiedSccOrder(
    dependencies: Map[QualifiedRef, Set[QualifiedRef]]
  ): Vector[Scc] =
    var collected = Vector.empty[Scc]
    foreachSccOf(dependencies) { scc =>
      val cyclic = scc match
        case single :: Nil => dependencies.get(single).exists(_.contains(single))
        case _ => true
      val kept = scc.filter(dependencies.contains)
      if kept.nonEmpty then
        collected = collected :+ Scc(kept.toVector.sortBy(q => (q.sheet.value, q.ref.toA1)), cyclic)
    }
    val comps = collected
    if comps.isEmpty then Vector.empty
    else
      val compOf: Map[QualifiedRef, Int] =
        comps.zipWithIndex.flatMap((c, i) => c.members.map(_ -> i)).toMap
      // Condensation edges, deduplicated, self-edges dropped: `precedentsOf(i)` are the components
      // i reads, `dependentsOf(i)` the components that read i.
      val (precedentsOf, dependentsOf) =
        comps.zipWithIndex.foldLeft((Map.empty[Int, Set[Int]], Map.empty[Int, Set[Int]])) {
          case (acc, (c, i)) =>
            c.members.foldLeft(acc) { case (acc2, u) =>
              dependencies.getOrElse(u, Set.empty).foldLeft(acc2) { case ((pre, dep), v) =>
                compOf.get(v) match
                  case Some(j) if j != i =>
                    (
                      pre.updated(i, pre.getOrElse(i, Set.empty) + j),
                      dep.updated(j, dep.getOrElse(j, Set.empty) + i)
                    )
                  case _ => (pre, dep)
              }
            }
        }
      // Kahn over the condensation with a key-ordered frontier: the topological order is then a
      // pure function of the graph's value, not of Map/Set iteration order (small Scala Maps and
      // Sets are insertion-ordered, which is exactly where test fixtures live).
      val byKey: Ordering[Int] = Ordering.by(i => (comps(i).key, i))
      val emptyFrontier = scala.collection.immutable.TreeSet.empty[Int](using byKey)
      val degree0 = comps.indices.map(i => i -> precedentsOf.getOrElse(i, Set.empty).size).toMap
      val initial =
        emptyFrontier ++ comps.indices.filter(i => precedentsOf.getOrElse(i, Set.empty).isEmpty)

      @tailrec
      def drain(
        frontier: scala.collection.immutable.TreeSet[Int],
        degree: Map[Int, Int],
        acc: Vector[Int]
      ): Vector[Int] =
        frontier.headOption match
          case None => acc
          case Some(i) =>
            val (nextDegree, newlyReady) =
              dependentsOf.getOrElse(i, Set.empty).foldLeft((degree, emptyFrontier)) {
                case ((d, ready), j) =>
                  val remaining = d.getOrElse(j, 0) - 1
                  (d.updated(j, remaining), if remaining == 0 then ready + j else ready)
              }
            drain((frontier - i) ++ newlyReady, nextDegree, acc :+ i)

      drain(initial, degree0, Vector.empty).map(comps)

  /**
   * GH-346: transitive dependents over the workbook-level reverse edges (excludes the starting
   * refs) — the qualified analog of `transitiveDependents`.
   */
  def qualifiedTransitiveDependents(
    dependents: Map[QualifiedRef, Set[QualifiedRef]],
    refs: Set[QualifiedRef]
  ): Set[QualifiedRef] =
    transitiveDependentsOf(dependents, refs)

  /**
   * GH-453: transitive PRECEDENTS over the workbook-level graph (excludes the starting refs) — the
   * same generic reachability BFS run over the FORWARD edge map instead of the reverse one. Used by
   * the data-table seeder to gate a table's source formula on the cyclic core.
   */
  def qualifiedTransitivePrecedents(
    dependencies: Map[QualifiedRef, Set[QualifiedRef]],
    refs: Set[QualifiedRef]
  ): Set[QualifiedRef] =
    transitiveDependentsOf(dependencies, refs)

  /**
   * GH-346: Kahn order over the workbook-level graph. Left carries the nodes that could not be
   * ordered (cycle participants and their downstream); callers that prune cycles first only reach
   * Left as a defensive-totality path.
   */
  def qualifiedTopologicalSort(
    dependencies: Map[QualifiedRef, Set[QualifiedRef]],
    dependents: Map[QualifiedRef, Set[QualifiedRef]]
  ): Either[Set[QualifiedRef], List[QualifiedRef]] =
    kahnOrder(dependencies, dependents)

  /** Unbounded range expansion — the original `fromWorkbook` semantics. */
  private val unboundedQualifiedCells: (SheetName, CellRange) => Set[QualifiedRef] =
    (sheet, range) => range.cells.map(ref => QualifiedRef(sheet, ref)).toSet

  /**
   * GH-522: expand any given (sheet, range) exactly once per graph build.
   *
   * A financial-model shape — many formulas each aggregating the same driver column — expanded the
   * SAME range once per referencing formula: 9,900 × `SUM($A$1:$A$5000)` built ~50M QualifiedRefs
   * (54 GB of a 135 GB recalc's allocation) where 5,000 suffice. The memo makes every referencing
   * formula share one immutable Set instance; correctness is untouched because expansion is a pure
   * function of (sheet, range) within one build (bounds are computed once up front and never change
   * mid-build). The captured mutable map is intentionally build-local and unsynchronized: callers
   * must invoke the returned closure only from the single graph-building thread.
   */
  private def memoizedCells(
    expand: (SheetName, CellRange) => Set[QualifiedRef]
  ): (SheetName, CellRange) => Set[QualifiedRef] =
    // Anchors affect formula dragging, never the cells a dependency range denotes. Keying by the
    // normalized endpoints lets `$A$1:$A$5000`, `A1:A5000`, and mixed-anchor spellings share the
    // same expansion within one graph build.
    val cache =
      scala.collection.mutable.HashMap.empty[(SheetName, ARef, ARef), Set[QualifiedRef]]
    (sheet, range) => cache.getOrElseUpdate((sheet, range.start, range.end), expand(sheet, range))

  /**
   * GH-522: set union that iterates the smaller side into the larger — `Set(oneRef) ++ rangeSet`
   * walks all 5,000 range elements today, once per operator node above a range reference. Union is
   * commutative, and any result with more than 4 elements is a CHAMP HashSet whose iteration order
   * is a function of its contents (apart from full hash-collision nodes, whose insertion sequence
   * is unchanged here), so redirecting the merge cannot move downstream order. When the larger side
   * has 4 or fewer elements the left-to-right build is kept: small sets are insertion-ordered and
   * Kahn's emitted order feeds off their iteration.
   */
  private def union(a: Set[QualifiedRef], b: Set[QualifiedRef]): Set[QualifiedRef] =
    if a.size < b.size && b.size > 4 then b ++ a else a ++ b

  /**
   * GH-522: reverse-edge map built with a local mutable accumulator. The immutable foldLeft
   * allocated a fresh outer-Map node chain per edge — 50M times on the shape above (8.5 GB). The A
   * mutable insertion-ordered set per precedent also avoids building one intermediate immutable Set
   * per edge. Each set receives refs in the identical sequence as the fold it replaces, then
   * freezes through the ordinary immutable Set builder, so its observable iteration order is
   * unchanged. The outer map's iteration order is deliberately not a contract; every consumer
   * performs keyed lookup or set membership.
   */
  private[formula] def reverseEdges(
    dependencies: Map[QualifiedRef, Set[QualifiedRef]]
  ): Map[QualifiedRef, Set[QualifiedRef]] =
    val acc = scala.collection.mutable.HashMap.empty[
      QualifiedRef,
      scala.collection.mutable.LinkedHashSet[QualifiedRef]
    ]
    dependencies.foreach { case (ref, deps) =>
      deps.foreach { dep =>
        acc.getOrElseUpdate(dep, scala.collection.mutable.LinkedHashSet.empty) += ref
      }
    }
    acc.iterator.map((ref, refs) => ref -> refs.toSet).toMap

  /**
   * Extract all qualified cell references from TExpr.
   *
   * Similar to extractDependencies but returns QualifiedRef to track cross-sheet references.
   * Same-sheet references are qualified with the current sheet name. Range expansion is delegated
   * to `cellsFor` so workbook-level callers can bound ranges by the TARGET sheet's used range
   * (GH-346) — the qualified analog of `extractDependenciesBounded`; single refs stay exact.
   *
   * @param expr
   *   The expression to analyze
   * @param currentSheet
   *   The sheet containing the formula (used for same-sheet ref qualification)
   * @param cellsFor
   *   Expands a range on a named sheet to qualified cells (possibly bounded)
   * @param workbook
   *   GH-384: name table for defined-name resolution (None disables name edges)
   * @param visitingNames
   *   GH-384: UPPERCASED names on the current resolution path — the name→name cycle guard
   * @return
   *   Set of qualified cell references used in the expression
   */
  @nowarn("msg=Unreachable case")
  private def extractQualifiedDependencies[A](
    expr: TExpr[A],
    currentSheet: SheetName,
    cellsFor: (SheetName, CellRange) => Set[QualifiedRef] = unboundedQualifiedCells,
    workbook: Option[Workbook] = None,
    visitingNames: Set[String] = Set.empty
  ): Set[QualifiedRef] =
    def locCells(location: TExpr.RangeLocation): Set[QualifiedRef] =
      location match
        case TExpr.RangeLocation.Local(range) => cellsFor(currentSheet, range)
        case TExpr.RangeLocation.CrossSheet(sheet, range) => cellsFor(sheet, range)
        // GH-353: external-workbook ranges target cells OUTSIDE the workbook — no edges ever
        case TExpr.RangeLocation.External(_, _, _) => Set.empty
        // GH-394: a name in a range slot resolves like the equivalent name EXPRESSION — its
        // target's cells contribute edges (qualified to the DEFINING sheet) so recalc orders
        // name-gated dependents correctly; a sheet-qualified name looks up relative to its
        // qualifier. Unresolvable names contribute no edges (evaluation reports the error).
        case TExpr.RangeLocation.Name(name, scope) =>
          scope match
            case None => go(TExpr.NameRef(name))
            case Some(qualifier) => go(TExpr.SheetNameRef(qualifier, name))

    def go(e: TExpr[?]): Set[QualifiedRef] =
      e match
        // Same-sheet references - qualify with current sheet
        case TExpr.Ref(at, _, _) => Set(QualifiedRef(currentSheet, at))
        case TExpr.PolyRef(at, _) => Set(QualifiedRef(currentSheet, at))
        case TExpr.RangeRef(range) => cellsFor(currentSheet, range)

        // Cross-sheet references - use target sheet
        case TExpr.SheetRef(sheet, at, _, _) => Set(QualifiedRef(sheet, at))
        case TExpr.SheetPolyRef(sheet, at, _) => Set(QualifiedRef(sheet, at))
        case TExpr.SheetRange(sheet, range) => cellsFor(sheet, range)
        // GH-353: external-workbook refs target cells OUTSIDE the workbook — no edges ever
        case TExpr.ExternalRef(_, _, _, _) => Set.empty
        case TExpr.ExternalRange(_, _, _) => Set.empty
        case TExpr.Add(l, r) => union(go(l), go(r))
        case TExpr.Sub(l, r) => union(go(l), go(r))
        case TExpr.Mul(l, r) => union(go(l), go(r))
        case TExpr.Div(l, r) => union(go(l), go(r))
        case TExpr.Pow(l, r) => union(go(l), go(r))
        case TExpr.Concat(l, r) => union(go(l), go(r))
        case TExpr.Eq(l, r) => union(go(l), go(r))
        case TExpr.Neq(l, r) => union(go(l), go(r))
        case TExpr.Lt(l, r) => union(go(l), go(r))
        case TExpr.Lte(l, r) => union(go(l), go(r))
        case TExpr.Gt(l, r) => union(go(l), go(r))
        case TExpr.Gte(l, r) => union(go(l), go(r))
        // Unary operators
        case TExpr.ToInt(x) => go(x)
        case TExpr.UnaryPlus(x) => go(x)
        case TExpr.Percent(x) => go(x)

        // Reference functions
        case TExpr.Aggregate(_, location) => locCells(location)

        case call: TExpr.Call[?] =>
          val values = call.spec.argSpec.toValues(call.args)
          values.foldLeft(Set.empty[QualifiedRef]) { (acc, value) =>
            value match
              case ArgValue.Expr(expr) => union(acc, go(expr))
              case ArgValue.Range(range) => union(acc, locCells(range))
              case ArgValue.Cells(range) => union(acc, cellsFor(currentSheet, range))
          }

        // GH-193: LET — union of binding-value and body dependencies; BindingRef has none
        case TExpr.Let(bindings, body) =>
          bindings.foldLeft(go(body)) { case (acc, (_, value)) => union(acc, go(value)) }
        case TExpr.BindingRef(_) => Set.empty
        case TExpr.CoercedBindingRef(_, _) => Set.empty

        // GH-384: resolve the defined name and recurse into its refersTo, qualifying its
        // same-sheet refs to the DEFINING sheet — =IF(case=2, ...) contributes an edge to the
        // cells 'case' targets, so Kahn orders a computed toggle before its name-gated
        // dependents and Tarjan sees cycles routed through names. Name→name chains carry a
        // visited guard; unresolvable/unparseable names contribute no edges (evaluation
        // reports the per-cell error).
        case TExpr.NameRef(name) =>
          val key = name.toUpperCase
          if visitingNames.contains(key) then Set.empty
          else
            (for
              wb <- workbook
              dn <- Evaluator.lookupDefinedName(wb, currentSheet, name)
              target <- FormulaParser.parse(dn.formula).toOption
            yield
              val definingSheet =
                Evaluator.definedNameScope(wb, dn).map(_.name).getOrElse(currentSheet)
              extractQualifiedDependencies(
                target,
                definingSheet,
                cellsFor,
                workbook,
                visitingNames + key
              )
            ).getOrElse(Set.empty)

        // GH-394: a sheet-qualified name resolves like NameRef, but the lookup runs as seen
        // from its QUALIFIER (sheet-scoped names on that sheet shadow workbook-scoped ones)
        case TExpr.SheetNameRef(qualifier, name) =>
          val key = name.toUpperCase
          if visitingNames.contains(key) then Set.empty
          else
            (for
              wb <- workbook
              dn <- Evaluator.lookupDefinedName(wb, qualifier, name)
              target <- FormulaParser.parse(dn.formula).toOption
            yield
              val definingSheet =
                Evaluator.definedNameScope(wb, dn).map(_.name).getOrElse(currentSheet)
              extractQualifiedDependencies(
                target,
                definingSheet,
                cellsFor,
                workbook,
                visitingNames + key
              )
            ).getOrElse(Set.empty)

        // GH-306: runtime coercion wrapper — transparent for analysis
        case TExpr.Coerced(inner, _) => go(inner)

        // Literals and nullary functions (no dependencies)
        case TExpr.Lit(_) => Set.empty
        case TExpr.DateToSerial(dateExpr) => go(dateExpr)
        case TExpr.DateTimeToSerial(dtExpr) => go(dtExpr)

    go(expr)

  /**
   * Detect circular references across sheets using an explicit-stack depth-first search.
   *
   * Similar to detectCycles but works with QualifiedRef to detect cycles that span multiple sheets.
   * A cross-sheet cycle occurs when cells across different sheets form a circular dependency (e.g.,
   * Sheet1!A1 = Sheet2!B1, Sheet2!B1 = Sheet1!A1).
   *
   * @param graph
   *   The cross-sheet dependency graph from fromWorkbook
   * @return
   *   Left(CircularRef) if cycle detected, Right(()) if acyclic
   *
   * Example:
   * {{{
   * val graph = DependencyGraph.fromWorkbook(workbook)
   * DependencyGraph.detectCrossSheetCycles(graph) match
   *   case Left(err) => println(s"Circular reference: $err")
   *   case Right(_) => println("No cycles")
   * }}}
   */
  def detectCrossSheetCycles(
    graph: Map[QualifiedRef, Set[QualifiedRef]]
  ): Either[EvalError.CircularRef, Unit] =
    firstCycleOf(
      graph,
      graph.keySet,
      (q: QualifiedRef) => (q.sheet.value, q.ref.row.index0, q.ref.col.index0)
    ) match
      // CircularRef predates QualifiedRef and therefore retains its sheetless public payload.
      case Some(cycle) => scala.util.Left(EvalError.CircularRef(cycle.map(_.ref)))
      case None => scala.util.Right(())
