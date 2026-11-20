package com.tjclp.xl.formula.graph

import com.tjclp.xl.formula.ast.{BinarySpine, BindingCoercion, TExpr}
import com.tjclp.xl.formula.functions.{
  ArgValue,
  CriteriaRangeResize,
  FunctionRegistry,
  FunctionSpecs,
  ReferenceOperators,
  TableColumnReferences
}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.formula.eval.{EvalError, Evaluator, NameWalk, RecalcCounters}

import com.tjclp.xl.workbooks.Workbook

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.CellRange
import com.tjclp.xl.cells.{Cell, CellValue, FormulaKind}
import com.tjclp.xl.sheets.Sheet
import scala.annotation.{nowarn, tailrec}

/**
 * Formula dependency analysis: the reference extractors every graph builder shares
 * ([[extractQualifiedDependencies]], [[extractDependencies]]), the dynamic-reference and
 * unresolved-name classifications, the workbook-level condensation ([[qualifiedSccOrder]],
 * [[withinComponentOrder]]) over dense ids, and the declared-reads index behind `QualifiedGraph`.
 * Recalculation orders with the compact graph (`CompactGraph`); #891 removed the per-sheet graph
 * (`fromSheet`, `topologicalSort`, `detectCycles`) and its hash-map engines.
 */
object DependencyGraph:
  private def depsFromArgValues(
    values: List[ArgValue],
    exprDeps: TExpr[?] => Set[ARef],
    rangeDeps: TExpr.RangeLocation => Set[ARef]
  ): Set[ARef] =
    values.foldLeft(Set.empty[ARef]) { (acc, value) =>
      val next = value match
        case ArgValue.Expr(expr) => exprDeps(expr)
        case ArgValue.Range(TExpr.RangeLocation.Computed(expr)) => exprDeps(expr)
        case ArgValue.Range(range) => rangeDeps(range)
      // Preserve a memoized range Set when it is the first argument. Besides avoiding a needless
      // rebuild, this lets every one-range aggregate in a graph share the exact cached expansion.
      if acc.isEmpty then next else if next.isEmpty then acc else acc ++ next
    }

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
      case TExpr.RangeRef(_, _) => true
      case TExpr.SheetRef(_, _, _, _) => true
      case TExpr.SheetPolyRef(_, _, _) => true
      case TExpr.SheetRange(_, _, _) => true
      // GH-353: external-workbook refs ARE cell references (they just live outside the workbook)
      case TExpr.ExternalRef(_, _, _, _) => true
      case TExpr.ExternalRange(_, _, _, _) => true
      case TExpr.Aggregate(_, _) => true

      case TableColumnReferences.Ref(_) => true

      // Function calls - check arguments
      case call: TExpr.Call[?] =>
        call.spec.argSpec
          .toValues(call.args)
          .exists {
            case ArgValue.Expr(e) => containsCellReferences(e)
            case ArgValue.Range(_) => true
          }

      // Binary operators - check both sides
      // GH-680: a chain's left spine in one loop, not one recursion per operator
      case chain @ (_: TExpr.Add | _: TExpr.Sub | _: TExpr.Mul | _: TExpr.Div | _: TExpr.Pow |
          _: TExpr.Concat | _: TExpr.Eq[?] | _: TExpr.Neq[?] | _: TExpr.Lt[?] | _: TExpr.Lte[?] |
          _: TExpr.Gt[?] | _: TExpr.Gte[?]) =>
        BinarySpine.existsOperand(chain)(containsCellReferences(_))

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

      case TExpr.ErrorLit(_, _) => false
      case TExpr.Missing => false

  /**
   * GH-274: Check whether an expression contains a dynamic-reference function call.
   *
   * A dynamic reference is a `TExpr.Call` whose spec is flagged `FunctionFlags.dynamicDeps` (e.g.
   * INDIRECT): its data dependencies are decided by evaluated text, so the static graph cannot see
   * them. The walk recurses into Call arguments; `ArgValue.Range`/`Cells` arguments cannot contain
   * calls. A resized SUMIF/AVERAGEIF whose shape or origin only evaluation knows — a LET name, or a
   * defined name computing its reference (`nameIsComputed`) — also has dynamic dependencies: the
   * graph cannot record the resized target (#733: one bound to a plain range resizes statically).
   *
   * @param nameIsComputed
   *   whether a defined name in a range slot computes its reference; the default, false, is the
   *   answer without a workbook, where a name resolves to nothing and so reads nothing
   */
  @nowarn("msg=Unreachable case")
  private[graph] def referencesMatching[A](
    expr: TExpr[A],
    resolveName: (String, Option[SheetName]) => Boolean,
    includeDynamicCalls: Boolean,
    nameIsComputed: (String, Option[SheetName]) => Boolean = (_, _) => false
  ): Boolean =
    def recurse(e: TExpr[?]): Boolean =
      referencesMatching(e, resolveName, includeDynamicCalls, nameIsComputed)
    expr match
      case call: TExpr.Call[?] =>
        val values = call.spec.argSpec.toValues(call.args)
        // GH-713: a computed range its own edges do not cover reads cells the graph cannot see
        (includeDynamicCalls && (call.spec.flags.dynamicDeps ||
          ReferenceOperators.isComputedRangeDynamic(call, ReferenceOperators.staticShape) ||
          CriteriaRangeResize.hasDynamicResize(
            call.spec.name,
            values,
            named => nameIsComputed(named.name, named.scope)
          ))) || values.exists {
          case ArgValue.Expr(e) => recurse(e)
          // GH-669: a table column's cells come from table metadata, which both graphs resolve
          // exactly, so it is static like a plain range
          case ArgValue.Range(TExpr.RangeLocation.Computed(TableColumnReferences.Ref(_))) => false
          case ArgValue.Range(TExpr.RangeLocation.Computed(e)) =>
            includeDynamicCalls || recurse(e)
          case ArgValue.Range(TExpr.RangeLocation.Name(name, scope)) =>
            resolveName(name, scope)
          case ArgValue.Range(_) => false
        }
      // GH-680: a chain's left spine in one loop, not one recursion per operator
      case chain @ (_: TExpr.Add | _: TExpr.Sub | _: TExpr.Mul | _: TExpr.Div | _: TExpr.Pow |
          _: TExpr.Concat | _: TExpr.Eq[?] | _: TExpr.Neq[?] | _: TExpr.Lt[?] | _: TExpr.Lte[?] |
          _: TExpr.Gt[?] | _: TExpr.Gte[?]) =>
        BinarySpine.existsOperand(chain)(recurse)
      case TExpr.ToInt(e) => recurse(e)
      case TExpr.UnaryPlus(e) => recurse(e)
      case TExpr.Percent(e) => recurse(e)
      case TExpr.DateToSerial(e) => recurse(e)
      case TExpr.DateTimeToSerial(e) => recurse(e)
      // GH-306: runtime coercion wrapper — a coerced INDIRECT/OFFSET is still dynamic
      case TExpr.Coerced(inner, _) => recurse(inner)
      // GH-193: LET — binding values and the body may carry dynamic calls
      case TExpr.Let(bindings, body) =>
        bindings.exists((_, value) => recurse(value)) || recurse(body)
      case TExpr.NameRef(name) => resolveName(name, None)
      case TExpr.SheetNameRef(sheet, name) => resolveName(name, Some(sheet))
      case TExpr.Aggregate(_, location) =>
        location match
          case TExpr.RangeLocation.Name(name, scope) => resolveName(name, scope)
          case _ => false
      // Leaves, listed rather than a wildcard: a new TExpr case is a "match may not be exhaustive"
      // warning here until it is classified as a wrapper to descend through or a leaf
      case _: TExpr.Lit[?] | _: TExpr.Ref[?] | _: TExpr.PolyRef | _: TExpr.SheetRef[?] |
          _: TExpr.SheetPolyRef | _: TExpr.ExternalRef | _: TExpr.ExternalRange |
          _: TExpr.RangeRef | _: TExpr.SheetRange | _: TExpr.ErrorLit | TExpr.Missing |
          _: TExpr.BindingRef | _: TExpr.CoercedBindingRef[?] =>
        false

  private def containsDynamicReferenceResolved[A](
    expr: TExpr[A],
    resolveName: (String, Option[SheetName]) => Boolean,
    nameIsComputed: (String, Option[SheetName]) => Boolean = (_, _) => false
  ): Boolean = referencesMatching(expr, resolveName, includeDynamicCalls = true, nameIsComputed)

  /**
   * GH-507: readers whose dependencies cannot be proved by the static graph. Name resolution is
   * scoped exactly as evaluation, and only names actually reached by formulas are inspected. The
   * local memos avoid repeating definition parsing and name-chain walks across readers.
   *
   * A name reference is unknown when its qualifier names no sheet, the name is not defined, or the
   * names it reaches include an unparseable definition, an unknown reference or a cycle. #691: the
   * walk is a [[NameWalk.Walk]] over resolved nodes, so a name with a chain more than
   * [[NameWalk.MaxDepth]] names long is unknown (`tooDeep`) whichever reader reaches it first — the
   * verdict is the node's own, never one computed partway down another reader's path — and every
   * name costs one step however many readers and paths reach it.
   */
  private[xl] def unresolvedReaders(
    workbook: Workbook,
    parses: FormulaParses = FormulaParses.none
  ): Set[QualifiedRef] =
    val positions =
      workbook.sheets.zipWithIndex.reverseIterator.map((sheet, i) => sheet.name -> i).toMap
    val resolve = NameWalk.resolver(workbook)
    val parsed = scala.collection.mutable.HashMap.empty[String, Option[TExpr[?]]]

    def parse(text: String): Option[TExpr[?]] =
      parsed.getOrElseUpdate(text, FormulaParser.parse(text).toOption)

    // The node a reference resolves to; None when its qualifier names no sheet or no such name.
    def target(qualifier: Option[SheetName], name: String, current: SheetName) =
      if qualifier.exists(!positions.contains(_)) then None else resolve(qualifier, name, current)

    // A node's hit: an unparseable definition, or a reference that resolves to nothing. Stepped
    // once per node by the walk.
    def step(node: NameWalk.Node): NameWalk.Step[NameWalk.Node] =
      parse(node.definition.formula) match
        case None => NameWalk.Step(hit = true, Vector.empty)
        case Some(expr) =>
          val refs = Vector.newBuilder[NameWalk.Node]
          val hit = referencesMatching(
            expr,
            (name, scope) =>
              target(scope, name, node.sheet) match
                case None => true
                case Some(child) =>
                  refs += child
                  false
            ,
            includeDynamicCalls = false
          )
          val next = refs.result()
          NameWalk.Step(hit, if next.sizeIs > 1 then next.distinct else next)

    val walk = NameWalk.Walk(step)

    def nameUnknown(qualifier: Option[SheetName], name: String, current: SheetName): Boolean =
      target(qualifier, name, current).fold(true) { node =>
        val verdict = walk(node)
        verdict.found || verdict.cyclic || verdict.tooDeep
      }

    def expressionUnknown(expr: TExpr[?], current: SheetName): Boolean =
      referencesMatching(
        expr,
        (name, scope) => nameUnknown(scope, name, current),
        includeDynamicCalls = false
      )

    workbook.sheets.iterator.flatMap { sheet =>
      sheet.cells.iterator.collect {
        case cell @ Cell(ref, _, _, _) if (cell.value match
              case CellValue.Formula(_, _, _: FormulaKind.DataTable) => false
              case CellValue.Formula(text, _, _) =>
                // Closed-workbook external caches are pinned by the explicit evaluation contract.
                val external = text.contains('[') &&
                  com.tjclp.xl.formula.eval.SheetEvaluator.pinnedExternalCache(cell.value).isDefined
                !external && parses
                  .parse(sheet.name, ref, text)
                  .toOption
                  .fold(true)(expressionUnknown(_, sheet.name))
              case _ => false
            ) =>
          QualifiedRef(sheet.name, ref)
      }
    }.toSet

  def containsDynamicReference[A](expr: TExpr[A]): Boolean =
    containsDynamicReferenceResolved(expr, (_, _) => false)

  /**
   * GH-520: workbook-aware dynamic classification, including parseable defined-name chains.
   *
   * A name is resolved with the same case-insensitive, sheet-scoped-shadowing rules as evaluation:
   * for `Other!GlobalName`, lookup occurs as seen from `Other`, while an unqualified name inside a
   * workbook-scoped definition still resolves from the original formula's sheet. #691: a name's
   * verdict is a [[NameWalk.Walk]] from the node it resolves to (the definition and the sheet its
   * body resolves from), so a name cycle stops cleanly, every node costs one step however many
   * names reach it, and chains of any length leave the stack alone. A name reaching a dynamic call
   * remains dynamic even when another chain is too deep: lazy evaluation can skip the deep branch
   * or recover from its error. Its readers may also be `unresolvedReaders`; that does not replace
   * the dynamic scheduling which keeps their otherwise invisible targets' caches fresh.
   *
   * Ordinary names do not make their users dynamic. Definitions are parsed and memoized first; only
   * names whose parseable chains actually reach a dynamic function join the cheap substring
   * pre-filter used for cell formulas. This matters for financial models with tens of thousands of
   * formulas referring to a static scenario name such as `Case`.
   *
   * GH-537: the work is linear in the names, not sheets × names. Each definition parses once per
   * distinct text; case-variant spellings of one name (`case`/`CASE` — one name to every reader)
   * classify once; and a name with no sheet-scoped variant whose definition mentions no other name
   * is classified ONCE for the workbook — its verdict cannot depend on the reading sheet. Only
   * names that nest another name (which an unqualified reference resolves from the reader's sheet,
   * where a sheet-scoped variant may shadow it) keep the per-sheet classification.
   */
  def dynamicCells(workbook: Workbook): Set[QualifiedRef] = dynamicCellsTraced(workbook).cells

  /**
   * [[dynamicCells]] reusing `parses` (the pass's graph build, #790) for the candidate cells it
   * holds; the result is the same.
   */
  def dynamicCells(workbook: Workbook, parses: FormulaParses): Set[QualifiedRef] =
    dynamicCellsTraced(workbook, parses).cells

  /**
   * GH-537: [[dynamicCells]] with the work it performed — the structural pin RecalcPerfSpec asserts
   * (parses and classifications linear in the NAMES, not sheets × names), since no wall-clock
   * budget can distinguish the two algorithms on a fixture that runs in under a second either way.
   *
   * @param parses
   *   `FormulaParser.parse` calls: one per distinct definition or candidate cell text (memo misses)
   * @param classifications
   *   defined-name classifications (a node's own verdict, stepped once by the [[NameWalk.Walk]]):
   *   one per sheet-independent name, one per resolved node (definition × the sheet its body
   *   resolves from) for names that nest another name
   */
  private[formula] final case class DynamicCellsTrace(
    cells: Set[QualifiedRef],
    parses: Int,
    classifications: Int
  )

  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private[formula] def dynamicCellsTraced(
    workbook: Workbook,
    buildParses: FormulaParses = FormulaParses.none
  ): DynamicCellsTrace =
    var parses = 0
    var classifications = 0

    val dynamicFunctions = FunctionRegistry.dynamicFunctionNames
    // Token form for the cheap substring pre-filter over cell text (function names are upper).
    def upper(text: String): String = text.toUpperCase(java.util.Locale.ROOT)
    // GH-537: definitions and candidate cell formulas parse once per distinct TEXT (the memo
    // `unresolvedReaders` keeps); the per-(sheet, name) verdict is memoized separately, so a
    // workbook-scoped name read from 40 sheets is parsed once, not 40 times.
    val parsed = scala.collection.mutable.HashMap.empty[String, Option[TExpr[?]]]
    def parse(text: String): Option[TExpr[?]] =
      parsed.getOrElseUpdate(
        text, {
          parses += 1
          FormulaParser.parse(text).toOption
        }
      )
    // Positions computed once (inside the resolver): this function makes sheets × names lookups,
    // so the per-lookup O(sheets) indexWhere inside lookupDefinedName would add an
    // O(sheets² × names) term.
    val resolve = NameWalk.resolver(workbook)

    // GH-537: one classification per resolution key — `case` and `CASE` are one name to every
    // reader — under the first-declared spelling.
    val definedNames = workbook.metadata.definedNames
    val distinctNames: Vector[(String, String)] =
      definedNames.iterator.map(dn => (resolutionKey(dn.name), dn.name)).distinctBy(_._1).toVector
    // GH-537: a name is SHEET-INDEPENDENT when no sheet-scoped variant can shadow it on any sheet
    // and its workbook-scoped definition mentions no other name: lookup then returns the same
    // definition from every sheet and classification never consults the reading sheet, so ONE
    // verdict serves the whole workbook (a definition that fails to parse is not dynamic anywhere
    // and qualifies too). Anything nesting a name stays per sheet, conservatively: an unqualified
    // nested name resolves from the READER's sheet (`definingSheet` falls back to it) and may be
    // shadowed there — the GH-520 fixture (`Outer = Driver`, Calc's local `Driver` dynamic).
    val sheetScopedKeys: Set[String] =
      definedNames.iterator
        .filter(_.localSheetId.isDefined)
        .map(dn => resolutionKey(dn.name))
        .toSet
    def mentionsName(expr: TExpr[?]): Boolean =
      referencesMatching(expr, (_, _) => true, includeDynamicCalls = false)
    def sheetIndependent(key: String, name: String): Boolean =
      !sheetScopedKeys.contains(key) &&
        Evaluator
          .lookupDefinedNameAt(workbook, None, name)
          .forall(dn => parse(dn.formula).forall(target => !mentionsName(target)))
    val (independentNames, dependentNames) =
      distinctNames.partition((key, name) => sheetIndependent(key, name))
    val independentKeys: Set[String] = independentNames.iterator.map(_._1).toSet
    val globalVerdict = scala.collection.mutable.HashMap.empty[String, Boolean]
    val keys = scala.collection.mutable.HashMap.empty[String, String]

    // #733: whether a name in a resizing call's range slot computes its reference, as seen from
    // `current` — memoized per (name, qualifier, sheet): a book's SUMIFs share a few names. NOT
    // thread-safe: a mutable memo for this single-threaded pass; never share it with a parallel
    // per-sheet loop
    val canonical = sheetCanonicaliser(workbook)
    val computedVerdicts =
      scala.collection.mutable.HashMap.empty[(String, Option[SheetName], SheetName), Boolean]
    def nameIsComputed(current: SheetName)(name: String, qualifier: Option[SheetName]): Boolean =
      computedVerdicts.getOrElseUpdate(
        (resolutionKey(name), qualifier.map(canonical), current),
        workbook(current).toOption
          .exists(Evaluator.nameComputesRange(name, qualifier, _, workbook))
      )

    // A classification: one node's own verdict and references, stepped once by the walk.
    def step(node: NameWalk.Node): NameWalk.Step[NameWalk.Node] =
      classifications += 1
      parse(node.definition.formula).fold(NameWalk.Step.empty)(
        NameWalk.Step.of(
          _,
          node.sheet,
          resolve,
          (expr, viaName) =>
            containsDynamicReferenceResolved(
              expr,
              (name, scope) => viaName(scope, name),
              nameIsComputed(node.sheet)
            )
        )
      )

    val walk = NameWalk.Walk(step)

    def classify(node: NameWalk.Node): Boolean = walk(node).reach == NameWalk.Reach.Found

    def expressionIsDynamic(expr: TExpr[?], currentSheet: SheetName): Boolean =
      containsDynamicReferenceResolved(
        expr,
        (name, scope) => nameIsDynamic(name, scope, currentSheet),
        nameIsComputed(currentSheet)
      )

    def nameIsDynamic(name: String, qualifier: Option[SheetName], current: SheetName): Boolean =
      resolve(qualifier, name, current).exists { node =>
        val nameKey = keys.getOrElseUpdate(name, resolutionKey(name))
        // Sheet-invariant by construction (see above): its definition nests no name, so one
        // verdict serves every reading sheet.
        if independentKeys.contains(nameKey) then
          globalVerdict.getOrElseUpdate(nameKey, classify(node))
        else classify(node)
      }

    // Sheet-independent names are classified once (from any sheet — the first will do);
    // dependent names once per sheet, as before. A group is dynamic when its verdict is true from
    // any sheet.
    val dynamicKeys: Set[String] =
      workbook.sheets.headOption.fold(Set.empty[String]) { first =>
        independentNames.iterator.collect {
          case (key, name) if nameIsDynamic(name, None, first.name) => key
        }.toSet
      } ++
        workbook.sheets.iterator.flatMap { sheet =>
          dependentNames.iterator.collect {
            case (key, name) if nameIsDynamic(name, None, sheet.name) => key
          }
        }
    // Tokens are the pre-filter's upper form of EVERY declared spelling of a dynamic group, not
    // only the first: `toUpperCase` is not the resolution relation (the Kelvin sign U+212A
    // upper-cases to itself yet resolves as `k`), so with `K` declared first a reader spelled `k`
    // would otherwise slip past the substring test — pre-GH-537 every spelling had its own token.
    val nameTokens: Set[String] =
      definedNames.iterator
        .filter(dn => dynamicKeys.contains(resolutionKey(dn.name)))
        .map(dn => upper(dn.name))
        .toSet
    val candidateTokens = dynamicFunctions ++ nameTokens
    // A SUMIF/AVERAGEIF resized through a LET name or a computed defined name is dynamic too
    // (hasDynamicResize; #733: a name bound to a range resizes statically). Its own text pre-filter
    // keeps SUMIFS and SUMIF over written ranges unparsed. GH-713: a computed range its edges do
    // not cover is dynamic (see dynamicCells(sheet)). The pre-filter's verdict depends on the text
    // alone: each distinct text is scanned once
    val candidates = scala.collection.mutable.HashMap.empty[String, Boolean]
    def candidate(expression: String): Boolean =
      candidates.getOrElseUpdate(
        expression, {
          val text = upper(expression)
          candidateTokens.exists(text.contains) ||
          CriteriaRangeResize.mayResizeThroughName(expression, text) ||
          TableColumnReferences.mayContain(expression) || ReferenceOperators
            .mayContainComputedRange(expression) ||
          com.tjclp.xl.formula.functions.ReferenceFunctionNames
            .mayContainComputingCall(expression, text)
        }
      )
    val cells =
      workbook.sheets.iterator.flatMap { sheet =>
        sheet.cells.iterator.flatMap { case cell @ Cell(ref, _, _, _) =>
          cell.value match
            case CellValue.Formula(expression, _, _) if candidate(expression) =>
              buildParses.get(sheet.name, ref, expression).orElse(parse(expression)) match
                case Some(expr) if expressionIsDynamic(expr, sheet.name) =>
                  Some(QualifiedRef(sheet.name, ref))
                case _ => None
            case _ => None
        }
      }.toSet
    DynamicCellsTrace(cells, parses, classifications)

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
      case TExpr.RangeRef(_, _) => true
      case TExpr.Aggregate(_, TExpr.RangeLocation.Local(_, _)) => true

      // Qualified cell references - sheet already specified
      case TExpr.SheetRef(_, _, _, _) => false
      case TExpr.SheetPolyRef(_, _, _) => false
      case TExpr.SheetRange(_, _, _) => false
      // GH-353: external refs are fully qualified (workbook + sheet) — no ambient sheet needed
      case TExpr.ExternalRef(_, _, _, _) => false
      case TExpr.ExternalRange(_, _, _, _) => false
      case TExpr.Aggregate(_, TExpr.RangeLocation.CrossSheet(_, _, _)) => false
      case TExpr.Aggregate(_, TExpr.RangeLocation.External(_, _, _, _)) => false
      // GH-394: an unqualified name's lookup depends on the ambient sheet (sheet-scoped names
      // shadow workbook-scoped ones); a sheet-qualified name carries its own context
      case TExpr.Aggregate(_, TExpr.RangeLocation.Name(_, scope)) => scope.isEmpty
      case TExpr.Aggregate(_, TExpr.RangeLocation.Error(_, _)) => false
      // GH-710: a LET name's binding value carries its own references (the Let case)
      case TExpr.Aggregate(_, TExpr.RangeLocation.Binding(_)) => false
      case TExpr.Aggregate(_, TExpr.RangeLocation.Computed(e)) =>
        containsUnqualifiedCellReferences(e)

      // Function calls - check arguments
      case call: TExpr.Call[?] =>
        call.spec.argSpec
          .toValues(call.args)
          .exists {
            case ArgValue.Expr(e) => containsUnqualifiedCellReferences(e)
            case ArgValue.Range(loc) =>
              loc match
                case TExpr.RangeLocation.Local(_, _) => true
                case TExpr.RangeLocation.CrossSheet(_, _, _) => false
                // GH-353: fully qualified (workbook + sheet) — no ambient sheet needed
                case TExpr.RangeLocation.External(_, _, _, _) => false
                // GH-394: unqualified name lookup depends on the ambient sheet
                case TExpr.RangeLocation.Name(_, scope) => scope.isEmpty
                case TExpr.RangeLocation.Binding(_) => false
                case TExpr.RangeLocation.Computed(e) => containsUnqualifiedCellReferences(e)
                case TExpr.RangeLocation.Error(_, _) => false
          }

      // Binary operators - check both sides
      // GH-680: a chain's left spine in one loop, not one recursion per operator
      case chain @ (_: TExpr.Add | _: TExpr.Sub | _: TExpr.Mul | _: TExpr.Div | _: TExpr.Pow |
          _: TExpr.Concat | _: TExpr.Eq[?] | _: TExpr.Neq[?] | _: TExpr.Lt[?] | _: TExpr.Lte[?] |
          _: TExpr.Gt[?] | _: TExpr.Gte[?]) =>
        BinarySpine.existsOperand(chain)(containsUnqualifiedCellReferences(_))

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

      case TExpr.ErrorLit(_, _) => false
      case TExpr.Missing => false

  /**
   * Extract all cell references from TExpr.
   *
   * Recursively traverses the expression AST and collects all cell references, including:
   *   - Single cell references (Ref)
   *   - Range references (RangeRef) expanded to all cells in range
   *
   * @param expr
   *   The expression to analyze
   * @param reader
   *   #743: the cell holding the formula, when known: INDEX at a position written with ROW() or
   *   COLUMN() then depends on the cells it selects, not on its whole array
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
  def extractDependencies[A](expr: TExpr[A], reader: Option[ARef] = None): Set[ARef] =
    def go(child: TExpr[?], inner: Option[ARef] = reader): Set[ARef] =
      extractDependencies(child, inner)
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
      case TExpr.SheetRange(_, _, _) => Set.empty
      // GH-353: external-workbook refs target cells OUTSIDE the workbook — no edges ever
      case TExpr.ExternalRef(_, _, _, _) => Set.empty
      case TExpr.ExternalRange(_, _, _, _) => Set.empty
      case TExpr.RangeRef(range, _) =>
        range.cells.toSet

      case call: TExpr.Call[?] =>
        // GH-631: SUMIF/AVERAGEIF read their third argument resized to their first
        // GH-669: an intersection of static locations depends only on the cells they share
        // #743: INDEX at a position folded for the reader reads its selected cells alone
        val inner = ReferenceOperators.operandReader(call, ReferenceOperators.staticShape, reader)
        depsFromArgValues(
          ReferenceOperators.callDependencyValues(
            call,
            CriteriaRangeResize.resizedArgs(call.spec.name, call.spec.argSpec.toValues(call.args)),
            ReferenceOperators.staticShape,
            reader
          ),
          go(_, inner),
          _.localCells
        )

      // Recursive cases (binary operators)
      // GH-680: a chain's left spine in one loop, not one recursion per operator
      case chain @ (_: TExpr.Add | _: TExpr.Sub | _: TExpr.Mul | _: TExpr.Div | _: TExpr.Pow |
          _: TExpr.Concat | _: TExpr.Eq[?] | _: TExpr.Neq[?] | _: TExpr.Lt[?] | _: TExpr.Lte[?] |
          _: TExpr.Gt[?] | _: TExpr.Gte[?]) =>
        BinarySpine.operandList(chain).foldLeft(Set.empty[ARef])(_ ++ go(_))
      case TExpr.ToInt(expr) =>
        go(expr) // Type conversion - extract from wrapped expr
      // GH-374: unary plus is transparent — dependencies under it feed recalc edges
      case TExpr.UnaryPlus(expr) => go(expr)
      // GH-355: postfix percent — dependencies under it feed recalc edges
      case TExpr.Percent(expr) => go(expr)
      case TExpr.Aggregate(_, location) => location.localCells

      // GH-193: LET — union of binding-value and body dependencies; BindingRef has none
      case TExpr.Let(bindings, body) =>
        bindings.foldLeft(go(body)) { case (acc, (_, value)) =>
          acc ++ go(value)
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
      case TExpr.Coerced(inner, _) => go(inner)

      // Literals and nullary functions (no dependencies)
      case TExpr.Lit(_) => Set.empty

      case TExpr.ErrorLit(_, _) => Set.empty
      case TExpr.Missing => Set.empty
      case TExpr.DateToSerial(dateExpr) => go(dateExpr)
      case TExpr.DateTimeToSerial(dtExpr) => go(dtExpr)

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

    // #891: mix the row and the column as separate words. The derived hash mixed `ref` as a boxed
    // Long, whose hash was row ^ column (one hash per anti-diagonal) until ARef folded the row
    // into its low word; mixing the two Ints directly stays spread whatever the packing and boxes
    // nothing.
    override def hashCode: Int =
      import scala.util.hashing.MurmurHash3
      val withSheet = MurmurHash3.mix(MurmurHash3.productSeed, sheet.hashCode)
      val withRow = MurmurHash3.mix(withSheet, ref.row.index0)
      MurmurHash3.finalizeHash(MurmurHash3.mixLast(withRow, ref.col.index0), 3)

  object QualifiedRef:
    /**
     * The GRID order: sheet name, then row, then column — row-major within a sheet, Excel's
     * documented left-to-right, top-to-bottom sweep. Every positional ordering in this graph
     * (component listings, the condensation's tie-break, the Gauss–Seidel cut) keys on it.
     *
     * Never the A1 TEXT: `"A10" < "A2"` and `"AA1" < "B1"` as strings, so a text key would sweep
     * row 10 before row 2 and column AA before column B — and, under Gauss–Seidel, hand isomorphic
     * models different kept values by the digit count of a row number.
     */
    def gridKey(q: QualifiedRef): (String, Int, Int) =
      (q.sheet.value, q.ref.row.index0, q.ref.col.index0)

    /**
     * [[gridKey]] as an `Ordering` — [[QualifiedGraph.byPosition]], the one order every layer and
     * every CLI listing already uses.
     */
    val gridOrdering: Ordering[QualifiedRef] = QualifiedGraph.byPosition

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

    val canonicalSheet = sheetCanonicaliser(workbook)
    val nameRanges = NameRanges.empty
    workbook.sheets.foreach { sheet =>
      sheet.cells.iterator.foreach { case cell @ Cell(cellRef, _, _, _) =>
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
                  Some(workbook),
                  canonicalSheet = canonicalSheet,
                  nameRanges = nameRanges,
                  reader = Some(cellRef)
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
   * GH-492: one node of the SCC condensation of the workbook-level graph.
   *
   * `members` is sorted by [[QualifiedRef.gridKey]] (sheet name, row, column) — the deterministic
   * member order the Jacobi fixpoint needs (values are order-independent under Jacobi, but error
   * vectors, report members and seeded-RNG draw order are not). `cyclic` is true for a multi-node
   * component or a self-looping singleton.
   *
   * No `derives CanEqual`: [[QualifiedRef]] does not derive it.
   */
  final case class Scc(members: Vector[QualifiedRef], cyclic: Boolean):
    /** Canonical key: the minimum member under [[QualifiedRef.gridKey]]. Unique across nodes. */
    private[xl] def key: (String, Int, Int) =
      members.headOption.fold(("", 0, 0))(QualifiedRef.gridKey)

  /**
   * GH-492: the SCC condensation of the workbook-level graph, in the KEY-MINIMUM dependency-first
   * topological order — every precedent component strictly precedes the components that read it,
   * ties broken by [[Scc.key]] (the grid order of each component's top-left member).
   *
   * This is what lets one iterative recalculation reach the workbook's GLOBAL fixpoint: walking the
   * condensation once (acyclic components evaluate, cyclic components fixpoint) guarantees every
   * precedent is a freshly computed value before anything reads it, so no read can fall back to a
   * loaded cache.
   *
   * Laws (property-tested in DependencyGraphSpec and, against a hash-map Tarjan oracle,
   * QualifiedSccOracleSpec):
   *   - partition — `flatMap(_.members).toSet == dependencies.keySet`, no duplicates
   *   - order — for every edge u → v with different components, index(scc(v)) < index(scc(u))
   *   - cyclic — a component is cyclic exactly when it has two members or one reads itself
   *   - determinism — depends only on the graph's VALUE, never on Map/Set iteration order
   *
   * Runs in O(V + E + C log C). Only nodes present in `dependencies` (formula cells) are ordered:
   * Tarjan also visits constants reached as successors, and those are always singletons with no
   * outgoing edges, so filtering them out can never split a real component.
   */
  def qualifiedSccOrder(
    dependencies: Map[QualifiedRef, Set[QualifiedRef]]
  ): Vector[Scc] =
    val csr = QualifiedCsr.of(dependencies)
    SccKernel
      .condensation(csr.size, csr.size, csr.succOff, csr.succDst, csr.components)
      .map((members, cyclic) => Scc(members.iterator.map(csr.nodes.qualified).toVector, cyclic))

  /**
   * #891: the formulas of one sheet's read map `edges` (each formula and the cells it reads) in a
   * dependency-first order — the condensation of [[qualifiedSccOrder]], so ties break in grid order
   * — or `CircularRef` naming the first cyclic component's members when one exists. Cells that are
   * not keys are leaves. For a small closure (the CLI's `eval`), over dense ids and arrays.
   */
  private[xl] def orderOf(
    sheet: SheetName,
    edges: Map[ARef, Set[ARef]]
  ): Either[EvalError.CircularRef, List[ARef]] =
    val qualified =
      edges.map((ref, reads) => QualifiedRef(sheet, ref) -> reads.map(QualifiedRef(sheet, _)))
    val components = qualifiedSccOrder(qualified)
    components.find(_.cyclic) match
      case Some(cycle) =>
        val members = cycle.members.iterator.map(_.ref).toList
        scala.util.Left(EvalError.CircularRef(members ++ members.headOption.toList))
      case None => scala.util.Right(components.iterator.flatMap(_.members).map(_.ref).toList)

  /**
   * #891: a workbook-level dependency map as dense ids and one CSR edge array, the input of
   * [[SccKernel]].
   *
   * The nodes are the map's KEYS (the formula cells), numbered in grid order ([[NodeTable.of]]), so
   * comparing ids compares [[QualifiedRef.gridKey]]s. An edge runs from each precedent to its
   * reader; an edge to a cell that is not a key is dropped (a constant reads nothing, so it is a
   * singleton component that can never join a cycle), and a key that reads itself is a self-loop
   * bit rather than an edge.
   */
  private final class QualifiedCsr(
    val nodes: NodeTable,
    val succOff: IArray[Int],
    val succDst: IArray[Int],
    val selfLoops: scala.collection.immutable.BitSet
  ):
    def size: Int = nodes.formulaCount
    lazy val components: CompactGraph.Components =
      SccKernel.tarjan(size, size, succOff, succDst, selfLoops)

  private object QualifiedCsr:
    @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
    def of(dependencies: Map[QualifiedRef, Set[QualifiedRef]]): QualifiedCsr =
      val nodes = NodeTable.of(dependencies.keySet)
      val n = nodes.formulaCount
      val src = new IntBuf(math.max(16, n))
      val dst = new IntBuf(math.max(16, n))
      val selfLoops = scala.collection.mutable.BitSet.empty
      val outCount = new Array[Int](n + 1)
      dependencies.foreach { (reader, reads) =>
        val r = nodes.lookup(reader).getOrElse(-1)
        reads.foreach { precedent =>
          val p = if precedent == reader then r else nodes.lookup(precedent).getOrElse(-1)
          if p >= 0 then
            if p == r then selfLoops += r
            else
              src += p
              dst += r
              outCount(p + 1) += 1
        }
      }
      var k = 0
      while k < n do
        outCount(k + 1) += outCount(k)
        k += 1
      val succDst = new Array[Int](src.length)
      val fill = java.util.Arrays.copyOf(outCount, n)
      var e = 0
      while e < src.length do
        val p = src(e)
        succDst(fill(p)) = dst(e)
        fill(p) += 1
        e += 1
      new QualifiedCsr(
        nodes,
        IArray.unsafeFromArray(outCount),
        IArray.unsafeFromArray(succDst),
        selfLoops.toImmutable
      )

  /**
   * GH-482: the order in which ONE cyclic component's members evaluate inside a Gauss–Seidel round
   * — Kahn's algorithm on the subgraph the component induces, with a grid-ordered frontier
   * ([[QualifiedRef.gridKey]]: sheet name, then row, then column), CUTTING at the top-left
   * remaining node whenever the frontier runs dry (inside a strongly-connected component it is dry
   * from the start: every member reads another). The edges into a cut node from members not yet
   * emitted are the round's back edges — the reads that see the PREVIOUS round's value; every other
   * read sees a value already refreshed this round. Self-edges are back edges by definition and
   * never block.
   *
   * The cut is positional — Excel's documented sweep runs left to right, top to bottom — never the
   * A1 text: a text key would cut at A10 before A2 and at AA1 before B1, so the kept values of a
   * non-convergent cycle would flip on the digit count of a row number. Whether Excel's own calc
   * chain lands on the same member for a given book has not been verified against Excel here.
   *
   * A pure function of the graph's VALUE (the discipline of [[qualifiedSccOrder]]): the same
   * component yields the same order whatever the insertion order of `members`, `dependencies` or
   * their sets. Laws (DependencyGraphSpec): a permutation of `members`; determinism under shuffle;
   * on an acyclic member set it IS the key-ordered Kahn order (no cut happens).
   *
   * `dependencies` is the workbook graph (`u -> the cells u reads`); edges leaving the component
   * are ignored. Runs in O((V + E) log V).
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def withinComponentOrder(
    members: Vector[QualifiedRef],
    dependencies: Map[QualifiedRef, Set[QualifiedRef]]
  ): Vector[QualifiedRef] =
    // #891: members as grid-ordered ids (comparing ids compares gridKeys), member-to-member reads
    // as edge lists, ordered by the shared kernel
    val nodes = NodeTable.of(members)
    val n = nodes.formulaCount
    val src = new IntBuf(math.max(16, n))
    val dst = new IntBuf(math.max(16, n))
    var u = 0
    while u < n do
      val reader = nodes.qualified(u)
      dependencies.getOrElse(reader, Set.empty).foreach { precedent =>
        if precedent != reader then
          val p = nodes.lookup(precedent).getOrElse(-1)
          if p >= 0 then
            src += p
            dst += u
      }
      u += 1
    val order = cutKahn(n, src, dst)
    Vector.tabulate(n)(i => nodes.qualified(order(i)))

  /**
   * #891: [[withinComponentOrder]]'s Kahn-with-cut over `n` grid-ordered local ids and the distinct
   * member-to-member edges `src(e) -> dst(e)` (precedent to reader, no self-edges): the min-heap
   * frontier, and the top-left remaining id when it runs dry. Returns the local ids in order.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private[graph] def cutKahn(n: Int, src: IntBuf, dst: IntBuf): Array[Int] =
    // member-to-member reads as one CSR array, the frontier an Int min-heap. The top-left
    // REMAINING member only ever moves right: it is the first id not yet emitted, so a cursor that
    // skips emitted ids finds it.
    val outCount = new Array[Int](n + 1)
    val degree = new Array[Int](n)
    var e = 0
    while e < src.length do
      outCount(src(e) + 1) += 1
      degree(dst(e)) += 1
      e += 1
    var k = 0
    while k < n do
      outCount(k + 1) += outCount(k)
      k += 1
    val succDst = new Array[Int](src.length)
    val fill = java.util.Arrays.copyOf(outCount, n)
    e = 0
    while e < src.length do
      succDst(fill(src(e))) = dst(e)
      fill(src(e)) += 1
      e += 1
    val emitted = new Array[Boolean](n)
    val frontier = new IntMinHeap(n)
    k = 0
    while k < n do
      if degree(k) == 0 then frontier.push(k)
      k += 1
    var cursor = 0
    val out = new Array[Int](n)
    var count = 0
    while count < n do
      // Kahn's next node or — the frontier dry with members remaining — the top-left remaining
      // member: the CUT, whose not-yet-emitted precedents become this round's back edges.
      val node =
        if frontier.nonEmpty then frontier.pop()
        else
          while emitted(cursor) do cursor += 1
          cursor
      emitted(node) = true
      out(count) = node
      count += 1
      var d = outCount(node)
      val end = outCount(node + 1)
      while d < end do
        val dependent = succDst(d)
        degree(dependent) -= 1
        // a dependent already emitted by a cut must not re-enter the frontier
        if degree(dependent) == 0 && !emitted(dependent) then frontier.push(dependent)
        d += 1
    out

  /** Unbounded range expansion — the original `fromWorkbook` semantics. */
  private val unboundedQualifiedCells: (SheetName, CellRange) => Set[QualifiedRef] =
    (sheet, range) => range.cells.map(ref => QualifiedRef(sheet, ref)).toSet

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
   * Extract all qualified cell references from TExpr.
   *
   * Similar to extractDependencies but returns QualifiedRef to track cross-sheet references.
   * Same-sheet references are qualified with the current sheet name. Range expansion is delegated
   * to `cellsFor` so workbook-level callers can bound ranges by the TARGET sheet's used range
   * (GH-346); single refs stay exact.
   *
   * @param expr
   *   The expression to analyze
   * @param currentSheet
   *   The sheet containing the formula (used for same-sheet ref qualification)
   * @param cellsFor
   *   Expands a range on a named sheet to qualified cells (possibly bounded)
   * @param workbook
   *   GH-384: name table for defined-name resolution (None disables name edges)
   * @return
   *   Set of qualified cell references used in the expression
   */
  /**
   * The workbook's spelling of every sheet name, keyed case-insensitively: a formula's qualifier
   * (`sheet1!A1` on a book whose tab is `Sheet1`) is folded onto the tab it resolves to in Excel,
   * so the graph has one node per cell rather than a phantom per spelling. A qualifier naming no
   * sheet is returned as written. The first sheet wins for a duplicated name, like `indexWhere`.
   */
  private[xl] def sheetCanonicaliser(workbook: Workbook): SheetName => SheetName =
    val byFold: Map[String, SheetName] =
      workbook.sheets.reverseIterator
        .map(sheet => sheet.name.value.toLowerCase(java.util.Locale.ROOT) -> sheet.name)
        .toMap
    sheet => byFold.getOrElse(sheet.value.toLowerCase(java.util.Locale.ROOT), sheet)

  /**
   * What `expr`, a formula on `sheet`, reads on that sheet, never expanded: each cell it names and
   * each range it reads (a whole column stays one range). With a workbook, defined names resolve
   * through the same scoped, cycle-guarded walk as the workbook graph. References to other sheets
   * or books are left out.
   */
  private[formula] def localReads[A](
    expr: TExpr[A],
    sheet: SheetName,
    workbook: Option[Workbook] = None
  ): (Set[ARef], Vector[CellRange]) =
    val ranges = Vector.newBuilder[CellRange]
    val canonical = workbook.fold[SheetName => SheetName](identity)(sheetCanonicaliser)
    val local = canonical(sheet)
    val cells = extractQualifiedDependencies(
      expr,
      local,
      cellsFor = (target, range) =>
        if target == local then ranges += range
        Set.empty
      ,
      workbook = workbook,
      canonicalSheet = canonical
    )
    (cells.collect { case QualifiedRef(`local`, ref) => ref }, ranges.result())

  private[graph] def extractQualifiedDependencies[A](
    expr: TExpr[A],
    currentSheet: SheetName,
    cellsFor: (SheetName, CellRange) => Set[QualifiedRef] = unboundedQualifiedCells,
    workbook: Option[Workbook] = None,
    preciseLookups: Boolean = false,
    whatIfSelector: Option[QualifiedRef => Boolean] = None,
    canonicalSheet: SheetName => SheetName = identity,
    nameRanges: NameRanges = NameRanges.empty,
    parseName: String => Option[TExpr[?]] = text => FormulaParser.parse(text).toOption,
    reader: Option[ARef] = None
  ): Set[QualifiedRef] =
    // #691: defined names expand from a queue, breadth-first, never by recursion: the formula is
    // scanned, then each resolved name it reaches — keyed on the node (the definition and the sheet
    // its body resolves from), never on the name's text — is expanded once, at its shallowest
    // depth, and adds what its body reads. A diamond of names expands each once, a cycle stops, and
    // the stack holds one expression at a time however long the chain. A name deeper than
    // NameWalk.MaxDepth contributes no edges, like an unresolvable name (evaluation fails it too).
    val names = new NameQueue
    // #743: `reader` positions the formula's own expression; a name body is read without one
    def extract(e: TExpr[?], sheet: SheetName, depth: Int): Set[QualifiedRef] =
      extractAt(
        e,
        sheet,
        depth,
        names,
        cellsFor,
        workbook,
        preciseLookups,
        whatIfSelector,
        canonicalSheet,
        nameRanges,
        if depth == 0 then reader else None
      )
    @tailrec
    def drain(acc: Set[QualifiedRef]): Set[QualifiedRef] =
      names.next() match
        case None => acc
        case Some((node, depth)) =>
          // #767: `parseName` lets a whole-graph build parse each name body once
          val body = parseName(node.definition.formula)
          drain(body.fold(acc)(target => union(acc, extract(target, node.sheet, depth))))
    drain(extract(expr, currentSheet, 0))

  /**
   * The RESOLUTION relation: two spellings are one name to every reader exactly when
   * `DefinedNameIndex` resolves them as one, i.e. under `String.equalsIgnoreCase` — per code point
   * `toLowerCase(toUpperCase(cp))`. `String.toUpperCase` is not that relation (it would split a
   * Kelvin-sign `K` from `k`), so shadowing, dedup and the name memos are keyed on this form.
   */
  private[graph] def resolutionKey(name: String): String =
    val folded = new java.lang.StringBuilder(name.length)
    name
      .codePoints()
      .forEach(cp => folded.appendCodePoint(Character.toLowerCase(Character.toUpperCase(cp))))
    folded.toString

  /**
   * #733: a defined name's static (sheet, range), keyed on its [[resolutionKey]], its qualifier and
   * the sheet it is read from — shared across one graph build, so a book whose SUMIFs read the same
   * few names resolves each once per sheet. Holds `None` for a name that does not resolve to one.
   * NOT thread-safe: a mutable map for one single-threaded build; never share it across a parallel
   * per-sheet loop.
   */
  private[graph] type NameRanges =
    scala.collection.mutable.HashMap[(String, Option[SheetName], SheetName), Option[
      (Sheet, CellRange)
    ]]

  private[graph] object NameRanges:
    def empty: NameRanges = scala.collection.mutable.HashMap.empty

  /**
   * #691: the defined names one [[extractQualifiedDependencies]] has reached and not yet expanded,
   * first in first out, each with its depth (the names on its path, itself included). Starts as two
   * empty immutable collections, so a formula reading no name allocates nothing more.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private final class NameQueue:
    private var pending = scala.collection.immutable.Queue.empty[(NameWalk.Node, Int)]
    private var seen = Set.empty[NameWalk.Node]

    /** Queue `node`, unless it was reached before (no shallower: the queue is breadth-first). */
    def reach(node: NameWalk.Node, depth: Int): Unit =
      if depth <= NameWalk.MaxDepth && !seen.contains(node) then
        seen += node
        pending = pending.enqueue((node, depth))

    def next(): Option[(NameWalk.Node, Int)] =
      pending.dequeueOption.map { (head, rest) =>
        pending = rest
        head
      }

  /**
   * One expression's references, as [[extractQualifiedDependencies]] reads them, at `depth` (the
   * names on its path, 0 for the formula's own): a defined name is resolved and queued in `names`,
   * never expanded here.
   */
  @nowarn("msg=Unreachable case")
  private def extractAt(
    expr: TExpr[?],
    currentSheet: SheetName,
    depth: Int,
    names: NameQueue,
    cellsFor: (SheetName, CellRange) => Set[QualifiedRef],
    workbook: Option[Workbook],
    preciseLookups: Boolean,
    whatIfSelector: Option[QualifiedRef => Boolean],
    canonicalSheet: SheetName => SheetName,
    nameRanges: NameRanges,
    reader: Option[ARef]
  ): Set[QualifiedRef] =
    // the formula's own sheet, looked up once per expression (SUMIF/AVERAGEIF name resolution)
    lazy val source: Option[Sheet] = workbook.flatMap(_(currentSheet).toOption)
    // #733: a range slot's location as a (sheet, range), when the workbook can say statically —
    // a defined name's answer memoized per build, since a book's SUMIFs share a few names
    def resolved(location: TExpr.RangeLocation): Option[(Sheet, CellRange)] =
      def resolve: Option[(Sheet, CellRange)] =
        source.flatMap(Evaluator.resolveRangeLocation(location, _, workbook).toOption)
      location match
        case TExpr.RangeLocation.Name(name, scope) =>
          nameRanges.getOrElseUpdate(
            (resolutionKey(name), scope.map(canonicalSheet), currentSheet),
            resolve
          )
        case _ => resolve
    def locCells(location: TExpr.RangeLocation): Set[QualifiedRef] =
      location match
        case TExpr.RangeLocation.Local(range, _) => cellsFor(currentSheet, range)
        case TExpr.RangeLocation.CrossSheet(sheet, range, _) =>
          cellsFor(canonicalSheet(sheet), range)
        // GH-353: external-workbook ranges target cells OUTSIDE the workbook — no edges ever
        case TExpr.RangeLocation.External(_, _, _, _) => Set.empty
        // GH-394: a name in a range slot resolves like the equivalent name EXPRESSION — its
        // target's cells contribute edges (qualified to the DEFINING sheet) so recalc orders
        // name-gated dependents correctly; a sheet-qualified name looks up relative to its
        // qualifier. Unresolvable names contribute no edges (evaluation reports the error).
        case TExpr.RangeLocation.Name(name, scope) =>
          scope match
            case None => go(TExpr.NameRef(name))
            case Some(qualifier) => go(TExpr.SheetNameRef(canonicalSheet(qualifier), name))
        // GH-612: an error in a range slot has no cells
        case TExpr.RangeLocation.Error(_, _) => Set.empty
        // GH-710: a LET name — its binding's value expression contributes the edges
        case TExpr.RangeLocation.Binding(_) => Set.empty
        case TExpr.RangeLocation.Computed(expr) => go(expr)

    def fixedIndex(expr: TExpr[?]): Option[Int] =
      def number(value: Any): Option[Int] = value match
        case n: Int => Some(n)
        case n: BigDecimal if n.isValidInt => Some(n.toInt)
        case _ => None
      def cell(sheet: SheetName, ref: ARef): Option[Int] =
        // `whatIfSelector` answers for a selector cell a data table substitutes: true when the
        // caller's graph must not fold it (the what-if pass would select a strip the graph never
        // ordered or recomputed)
        if whatIfSelector.exists(_(QualifiedRef(sheet, ref))) then None
        else
          workbook.flatMap(_(sheet).toOption).flatMap { target =>
            target(ref).value match
              case CellValue.Number(n) => number(n)
              case _ => None // a cached formula may change during this generation
          }
      expr match
        case TExpr.Lit(value) => number(value)
        case TExpr.ToInt(inner) => fixedIndex(inner)
        case TExpr.UnaryPlus(inner) => fixedIndex(inner)
        case TExpr.Coerced(inner, BindingCoercion.Integer) => fixedIndex(inner)
        case TExpr.Ref(ref, _, _) => cell(currentSheet, ref)
        case TExpr.SheetRef(sheet, ref, _, _) => cell(canonicalSheet(sheet), ref)
        case _ => None

    // Ordering needs only the lookup's key and selected result strip. Impact analysis retains
    // the full declared range, so changing the selector still invalidates all relevant readers.
    def lookupCells(name: String, args: List[ArgValue]): Option[Set[QualifiedRef]] =
      if !preciseLookups || (name != "VLOOKUP" && name != "HLOOKUP") then None
      else
        args match
          case ArgValue.Expr(_) :: ArgValue.Range(location) :: ArgValue.Expr(selector) :: _ =>
            for
              index <- fixedIndex(selector)
              source <- workbook.flatMap(_(currentSheet).toOption)
              resolved <- Evaluator.resolveRangeLocation(location, source, workbook).toOption
            yield
              val (target, range) = resolved
              val vertical = name == "VLOOKUP"
              val extent = if vertical then range.width else range.height
              if index < 1 || index > extent then Set.empty[QualifiedRef]
              else
                def strip(offset: Int): CellRange =
                  if vertical then
                    val col = range.colStart.index0 + offset
                    CellRange(
                      ARef.from0(col, range.rowStart.index0),
                      ARef.from0(col, range.rowEnd.index0)
                    )
                  else
                    val row = range.rowStart.index0 + offset
                    CellRange(
                      ARef.from0(range.colStart.index0, row),
                      ARef.from0(range.colEnd.index0, row)
                    )
                union(cellsFor(target.name, strip(0)), cellsFor(target.name, strip(index - 1)))
          case _ => None

    // #691: a name's body is the caller's to expand (`names`); here it reads nothing
    def expandName(qualifier: Option[SheetName], name: String): Set[QualifiedRef] =
      for
        wb <- workbook
        dn <- Evaluator.lookupDefinedName(wb, qualifier.getOrElse(currentSheet), name)
      do
        names.reach(
          NameWalk.Node(dn, Evaluator.definedNameScope(wb, dn).fold(currentSheet)(_.name)),
          depth + 1
        )
      Set.empty

    def go(e: TExpr[?]): Set[QualifiedRef] = walk(reader)(e)

    // #743: `at`, the cell the formula sits in, as the operands of `e` are read (operandReader)
    def walk(at: Option[ARef])(e: TExpr[?]): Set[QualifiedRef] =
      def go(child: TExpr[?]): Set[QualifiedRef] = walk(at)(child)
      e match
        // Same-sheet references - qualify with current sheet
        case TExpr.Ref(at, _, _) => Set(QualifiedRef(currentSheet, at))
        case TExpr.PolyRef(at, _) => Set(QualifiedRef(currentSheet, at))
        case TExpr.RangeRef(range, _) => cellsFor(currentSheet, range)

        // Cross-sheet references - use the target sheet as the workbook spells it
        case TExpr.SheetRef(sheet, at, _, _) => Set(QualifiedRef(canonicalSheet(sheet), at))
        case TExpr.SheetPolyRef(sheet, at, _) => Set(QualifiedRef(canonicalSheet(sheet), at))
        case TExpr.SheetRange(sheet, range, _) => cellsFor(canonicalSheet(sheet), range)
        // GH-353: external-workbook refs target cells OUTSIDE the workbook — no edges ever
        case TExpr.ExternalRef(_, _, _, _) => Set.empty
        case TExpr.ExternalRange(_, _, _, _) => Set.empty
        // GH-680: a chain's left spine in one loop, not one recursion per operator
        case chain @ (_: TExpr.Add | _: TExpr.Sub | _: TExpr.Mul | _: TExpr.Div | _: TExpr.Pow |
            _: TExpr.Concat | _: TExpr.Eq[?] | _: TExpr.Neq[?] | _: TExpr.Lt[?] | _: TExpr.Lte[?] |
            _: TExpr.Gt[?] | _: TExpr.Gte[?]) =>
          BinarySpine
            .operandList(chain)
            .foldLeft(Set.empty[QualifiedRef])((acc, e) => union(acc, go(e)))
        // Unary operators
        case TExpr.ToInt(x) => go(x)
        case TExpr.UnaryPlus(x) => go(x)
        case TExpr.Percent(x) => go(x)

        // Reference functions
        case TExpr.Aggregate(_, location) => locCells(location)

        case TableColumnReferences.Ref(column) =>
          source
            .flatMap(sheet => TableColumnReferences.resolve(column, sheet, workbook).toOption)
            .fold(Set.empty[QualifiedRef])(ref => cellsFor(ref.sheet.name, ref.range))

        case call: TExpr.Call[?] =>
          // GH-631: SUMIF/AVERAGEIF read their third argument resized to their first — a static
          // shape, or a defined name's resolved through the workbook when one is at hand
          // GH-669: an intersection of static locations depends only on the cells they share
          // #733: a defined name bound to a range resizes like the range itself — its target's
          // resized cells are static edges
          val raw = call.spec.argSpec.toValues(call.args)
          // GH-713: a range of names resolved through the workbook gets its bounding edges
          val shapeOf: ReferenceOperators.ShapeOf = location =>
            ReferenceOperators
              .staticShape(location)
              .orElse(resolved(location).map((sheet, range) => (Some(sheet.name), range)))
          // #743: INDEX at a position folded for the reader reads its selected cells alone
          val values = ReferenceOperators.callDependencyValues(
            call,
            CriteriaRangeResize.resizedArgs(
              call.spec.name,
              raw,
              shapeOf = location => location.staticRange.orElse(resolved(location).map(_._2)),
              targetOf = {
                case named: TExpr.RangeLocation.Name =>
                  resolved(named).fold[TExpr.RangeLocation](named) { (target, range) =>
                    TExpr.RangeLocation.CrossSheet(target.name, range)
                  }
                case computed @ TExpr.RangeLocation.Computed(TableColumnReferences.Ref(_)) =>
                  resolved(computed).fold[TExpr.RangeLocation](computed) { (target, range) =>
                    TExpr.RangeLocation.CrossSheet(target.name, range)
                  }
                case other => other
              }
            ),
            shapeOf,
            at
          )
          val inner = ReferenceOperators.operandReader(call, shapeOf, at)
          // A name the resize replaced resolved statically to a range or a cell, through any chain
          // of names: its resized CrossSheet is everything the call reads, so the name's full
          // range adds no edge — it would make `C5 = SUMIF(A1:A3,">0",amt)` with amt = C1:C10 a
          // false cycle, where the written `C1:C10` is not. A name that did not resolve is still in
          // `values`, and the fold visits it as before.
          val selected = lookupCells(call.spec.name, values)
          values.zipWithIndex.foldLeft(Set.empty[QualifiedRef]) { case (acc, (value, index)) =>
            value match
              case ArgValue.Expr(expr) => union(acc, walk(inner)(expr))
              case ArgValue.Range(_) if index == 1 && selected.isDefined =>
                union(acc, selected.getOrElse(Set.empty))
              case ArgValue.Range(range) => union(acc, locCells(range))
          }

        // GH-193: LET — union of binding-value and body dependencies; BindingRef has none
        case TExpr.Let(bindings, body) =>
          bindings.foldLeft(go(body)) { case (acc, (_, value)) => union(acc, go(value)) }
        case TExpr.BindingRef(_) => Set.empty
        case TExpr.CoercedBindingRef(_, _) => Set.empty

        // GH-384: resolve the defined name and recurse into its refersTo, qualifying its
        // same-sheet refs to the DEFINING sheet — =IF(case=2, ...) contributes an edge to the
        // cells 'case' targets, so Kahn orders a computed toggle before its name-gated
        // dependents and Tarjan sees cycles routed through names. Unresolvable/unparseable names
        // contribute no edges (evaluation reports the per-cell error).
        case TExpr.NameRef(name) => expandName(None, name)

        // GH-394: a sheet-qualified name resolves like NameRef, but the lookup runs as seen
        // from its QUALIFIER (sheet-scoped names on that sheet shadow workbook-scoped ones)
        case TExpr.SheetNameRef(qualifier, name) =>
          expandName(Some(canonicalSheet(qualifier)), name)

        // GH-306: runtime coercion wrapper — transparent for analysis
        case TExpr.Coerced(inner, _) => go(inner)

        // Literals and nullary functions (no dependencies)
        case TExpr.Lit(_) => Set.empty

        case TExpr.ErrorLit(_, _) => Set.empty
        case TExpr.Missing => Set.empty
        case TExpr.DateToSerial(dateExpr) => go(dateExpr)
        case TExpr.DateTimeToSerial(dtExpr) => go(dtExpr)

    go(expr)
