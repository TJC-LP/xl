package com.tjclp.xl.formula.graph

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.formula.ast.{BinarySpine, TExpr}
import com.tjclp.xl.formula.eval.{Evaluator, NameWalk}
import com.tjclp.xl.formula.functions.{ArgValue, CriteriaRangeResize, ReferenceOperators}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.workbooks.{DefinedName, Workbook}

/**
 * Value reads for the positioned-evaluation cycle check. Reference inspection (ROW, COLUMN, ROWS,
 * COLUMNS and AREAS) needs the address or shape, not the referenced cells' values. A computed
 * reference still reads its selectors: ROW(INDEX(A:A, B1)) depends on B1, not A:A. The ordinary
 * dependency graph remains conservative for structural edits and impact analysis.
 */
private[xl] object ValueDependencies:
  private final case class Bound(expr: TExpr[?], env: Map[String, Bound])

  /**
   * Functions that read only the address or shape of every argument, never a value. The specs live
   * in FunctionSpecsReference; ValueDependenciesSpec checks each name is registered and that its
   * result does not change with the referenced cell's value.
   */
  private[graph] val inspectors = Set("ROW", "COLUMN", "ROWS", "COLUMNS", "AREAS")

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
    // return_array alone: a reference given as if_not_found (index 3) still counts as a value
    // read, an over-report kept on purpose, since a false cycle is safer than a stale value
    "XLOOKUP" -> ((index, _) => index == 2),
    ReferenceOperators.UnionName -> ((_, _) => true),
    ReferenceOperators.IntersectionName -> ((_, _) => true),
    // GH-713: `:` between computed references reads its operands as references
    ReferenceOperators.RangeName -> ((_, _) => true)
  )

  /**
   * The cells whose values a formula reads on its own sheet ([[reads]]), with ranges bounded by
   * `bounds`. `home` is that sheet's name. A reference qualified with it (`Data!D5` in a formula on
   * Data) is a read like `D5`. Sheet names compare case-insensitively, as Excel's do and as
   * [[ReferenceOperators.dependencyValues]] compares them. Any other qualified reference belongs to
   * another sheet and reads nothing here.
   *
   * With `localReads = false` the formula sits on another sheet: only its references qualified with
   * `home` read `home`, and its unqualified ones read its own sheet, so nothing here (`--with`'s
   * search for another sheet's formulas that read back into the overridden one).
   *
   * With `names`, a defined name read as a value or in a range slot (`SUMIF(rng, …)`, GH-761) reads
   * what its definition reads ([[Names]]), as resolved from `current`, the sheet the formula sits
   * on (`home` when it is not given). A name bound to a range also sizes SUMIF's and AVERAGEIF's
   * third argument as that range does. Without, a name reads nothing: the placement guard catches a
   * read through one at run time.
   *
   * One reader serves many formulas on one sheet: each distinct range is expanded once for all of
   * them, as `DependencyGraph.fromSheet` does (GH-522). It is a single-use cache for one analysis:
   * not thread-safe, so never share one across fibers or threads (`recalc --parallel`). Its cache
   * keys on a range's corners alone, so the expansion is only right for this reader's `bounds`:
   * never reuse a reader where other bounds apply; build another, as the cycle check does for the
   * one-cell bound around the placement.
   */
  private[xl] final class Reader(
    bounds: Option[CellRange],
    home: Option[SheetName],
    localReads: Boolean = true,
    names: Option[Names] = None,
    current: Option[SheetName] = None
  ):
    private val expanded = scala.collection.mutable.HashMap.empty[(ARef, ARef), Set[ARef]]

    /**
     * The cells `expr` reads. `at`, #743: the cell the formula sits in, when known — INDEX at a
     * position written with ROW() or COLUMN() then reads the cells it selects, not its whole array.
     */
    def reads(expr: TExpr[?], at: Option[ARef] = None): Set[ARef] =
      readRanges(expr, at).foldLeft(Set.empty[ARef])((acc, r) => acc ++ expand(r))

    /**
     * [[reads]] as the ranges read, never expanded: each range bounded by `bounds`, a single cell
     * as its one-cell range (unbounded, as [[reads]] has it). For an index that must not pay for
     * every cell of every range (`--with`'s stale-reader search over running totals).
     */
    def readRanges(expr: TExpr[?], at: Option[ARef] = None): Set[CellRange] =
      loop(expr, false, Map.empty, at)

    private def expand(range: CellRange): Set[ARef] =
      if range.start == range.end then Set(range.start)
      else expanded.getOrElseUpdate((range.start, range.end), range.cells.toSet)

    private def cells(range: CellRange): Set[CellRange] =
      bounds.fold(Option(range))(range.intersect).fold(Set.empty[CellRange])(Set(_))

    private def local(sheet: SheetName): Boolean =
      home.exists(_.value.equalsIgnoreCase(sheet.value))

    private def loop(
      e: TExpr[?],
      referenceOnly: Boolean,
      env: Map[String, Bound],
      at: Option[ARef]
    ): Set[CellRange] =
      def value(child: TExpr[?]): Set[CellRange] = loop(child, false, env, at)
      def binding(name: String): Set[CellRange] =
        env.get(name).fold(Set.empty[CellRange])(b => loop(b.expr, referenceOnly, b.env, at))
      def cell(at: ARef): Set[CellRange] =
        if referenceOnly then Set.empty else Set(CellRange(at, at))
      def range(r: CellRange): Set[CellRange] = if referenceOnly then Set.empty else cells(r)
      // a defined name's reads: what its definition reads on `home`, resolved from `current`
      def resolved(qualifier: Option[SheetName], name: String): Set[CellRange] =
        (names, home, current.orElse(home)) match
          case (Some(resolver), Some(target), Some(from)) =>
            resolver.reads(qualifier, name, from, target, bounds)
          case _ => Set.empty
      // a defined name read as a value
      def named(qualifier: Option[SheetName], name: String): Set[CellRange] =
        if referenceOnly then Set.empty else resolved(qualifier, name)
      // the range a defined name in a range slot names, resolved from `current`, when it names one
      def namedRange(location: TExpr.RangeLocation): Option[(SheetName, CellRange)] =
        (location, names, current.orElse(home)) match
          case (TExpr.RangeLocation.Name(name, scope), Some(resolver), Some(from)) =>
            resolver.range(scope, name, from)
          case _ => None
      // an unqualified reference reads the formula's own sheet: `home` unless it sits elsewhere
      def localCell(at: ARef): Set[CellRange] = if localReads then cell(at) else Set.empty
      def localRange(r: CellRange): Set[CellRange] = if localReads then range(r) else Set.empty
      def arg(a: ArgValue, addressOnly: Boolean, reader: Option[ARef]): Set[CellRange] = a match
        case ArgValue.Expr(child) => loop(child, addressOnly, env, reader)
        case ArgValue.Range(location) =>
          // Non-static lexical locations print as their declared identifier. This also covers
          // range-slot bindings without confusing qualified names, cell addresses or errors with
          // legal LET names. Follow the binding in its defining scope and the consumer's mode.
          val bound = if location.staticRange.isEmpty then env.get(location.toA1) else None
          bound match
            case Some(b) => loop(b.expr, addressOnly, b.env, reader)
            case None if addressOnly => Set.empty
            case None =>
              location match
                case TExpr.RangeLocation.Local(r, _) => if localReads then cells(r) else Set.empty
                case TExpr.RangeLocation.CrossSheet(sheet, r, _) if local(sheet) => cells(r)
                // GH-761: a defined name in a range slot (SUMIF(rng, …)) reads what it names, as
                // a name read as a value does
                case TExpr.RangeLocation.Name(name, scope) => resolved(scope, name)
                case _ => Set.empty
        case ArgValue.Cells(r) => if addressOnly || !localReads then Set.empty else cells(r)

      // Exhaustive, no wildcard: a new TExpr case is a compiler warning here, never a silent
      // structural fallback that would lose referenceOnly and own-sheet qualified reads.
      e match
        // GH-680: a chain's left spine in one loop; unary minus is `0 - x`, so it lands here too
        case chain @ (_: TExpr.Add | _: TExpr.Sub | _: TExpr.Mul | _: TExpr.Div | _: TExpr.Pow |
            _: TExpr.Concat | _: TExpr.Eq[?] | _: TExpr.Neq[?] | _: TExpr.Lt[?] | _: TExpr.Lte[?] |
            _: TExpr.Gt[?] | _: TExpr.Gte[?]) =>
          BinarySpine.operandList(chain).foldLeft(Set.empty[CellRange])(_ ++ value(_))
        case TExpr.Ref(at, _, _) => localCell(at)
        case TExpr.PolyRef(at, _) => localCell(at)
        case TExpr.SheetRef(sheet, at, _, _) => if local(sheet) then cell(at) else Set.empty
        case TExpr.SheetPolyRef(sheet, at, _) => if local(sheet) then cell(at) else Set.empty
        case TExpr.RangeRef(r, _) => localRange(r)
        case TExpr.SheetRange(sheet, r, _) => if local(sheet) then range(r) else Set.empty
        case TExpr.Coerced(inner, _) => loop(inner, referenceOnly, env, at)
        case TExpr.ToInt(inner) => value(inner)
        case TExpr.UnaryPlus(inner) => value(inner)
        case TExpr.Percent(inner) => value(inner)
        case TExpr.DateToSerial(inner) => value(inner)
        case TExpr.DateTimeToSerial(inner) => value(inner)
        case TExpr.BindingRef(name) => binding(name)
        case TExpr.CoercedBindingRef(name, _) => binding(name)
        case TExpr.Let(bindings, body) =>
          val (scoped, bindingReads) = bindings.foldLeft((env, Set.empty[CellRange])) {
            case ((scope, reads), (name, definition)) =>
              (
                scope.updated(name, Bound(definition, scope)),
                reads ++ loop(definition, referenceDefinition(definition), scope, at)
              )
          }
          // Plain-cell LET stores a reference without reading its values. Eager scalar bindings
          // and reference selectors still evaluate, even if the body does not use the binding.
          bindingReads ++ loop(body, referenceOnly, scoped, at)
        case TExpr.Aggregate(_, location) => arg(ArgValue.Range(location), false, at)
        case call: TExpr.Call[?] =>
          val name = call.spec.name
          // GH-761: SUMIF and AVERAGEIF size their third argument to their first, a defined name
          // in either slot as the range it names, as evaluation and DependencyGraph size it (#733)
          // #743: INDEX at a position folded for the reader reads its selected cells alone
          val args = ReferenceOperators.callDependencyValues(
            call,
            CriteriaRangeResize.resizedArgs(
              name,
              call.spec.argSpec.toValues(call.args),
              shapeOf = location => location.staticRange.orElse(namedRange(location).map(_._2)),
              targetOf = location =>
                namedRange(location).fold(location)((sheet, range) =>
                  TExpr.RangeLocation.CrossSheet(sheet, range)
                )
            ),
            at = at
          )
          val inner = ReferenceOperators.operandReader(call, ReferenceOperators.staticShape, at)
          val passes = if referenceOnly then referencePassing.get(name) else None
          args.zipWithIndex.foldLeft(Set.empty[CellRange]) { case (acc, (a, index)) =>
            // CELL's reference is address-only because every info type `cellFn` supports
            // (FunctionSpecsReference: filename, address, row, col) reads the address, never the
            // value; info_type itself is a value read. An arm that reads the value ("contents",
            // "type") must make this a read, at least when info_type names that arm.
            // ValueDependenciesSpec pins the supported arms, so adding one fails until this is
            // revisited.
            val addressOnly = inspectors.contains(name) || (name == "CELL" && index == 1) ||
              passes.exists(_(index, args.size))
            acc ++ arg(a, addressOnly, inner)
          }
        // defined names: their targets live in workbook metadata ([[Names]] when given)
        case TExpr.NameRef(name) => named(None, name)
        case TExpr.SheetNameRef(sheet, name) => named(Some(sheet), name)
        // literals and external references: no cell of this sheet
        case TExpr.Lit(_) | TExpr.ErrorLit(_, _) | TExpr.Missing | TExpr.ExternalRef(_, _, _, _) |
            TExpr.ExternalRange(_, _, _, _) =>
          Set.empty

  /**
   * GH-715: what the defined names of `wb` read, for `--with`'s stale-reader search. A reference
   * resolves as the evaluator resolves it: from the qualifier's sheet, else the referencing
   * formula's, a sheet-scoped name shadowing a workbook-scoped one, its body read from its scope
   * sheet, else from the referencing formula's. The body's value reads count as the referencing
   * formula's own, through further names. A body that does not parse, a name that does not resolve,
   * and a chain past [[NameWalk.MaxDepth]] names read nothing, as the evaluator resolves none of
   * them; a name in a cycle reads what its chain found before the cycle closed, and its readers
   * evaluate to an error whatever the overrides.
   *
   * Single-use, like [[Reader]]: one analysis of one workbook, never shared across threads. Its
   * memo keys on the target sheet, so every reader it serves for one target must share `bounds`.
   */
  private[xl] final class Names(wb: Workbook):
    private val positions: Map[String, Int] =
      wb.sheets.zipWithIndex.reverseIterator.map((sheet, i) => fold(sheet.name) -> i).toMap
    private val bodies = scala.collection.mutable.HashMap.empty[DefinedName, Option[TExpr[?]]]
    // (definition, sheet its body resolves from, target) → what it reads there
    private val memo =
      scala.collection.mutable.HashMap.empty[(DefinedName, String, String), Set[CellRange]]
    private val open = scala.collection.mutable.Set.empty[(DefinedName, String, String)]
    // (name, qualifier, sheet it is read from) → the range it names, None when it names none
    private val ranges = scala.collection.mutable.HashMap
      .empty[(String, Option[String], String), Option[(SheetName, CellRange)]]

    /**
     * GH-761: the (sheet, range) a defined name in a range slot names, read from `current`, as
     * evaluation resolves it ([[Evaluator.resolveRangeLocation]]): a name bound to a range or a
     * cell, through any chain of names. None for a name that computes its reference (OFFSET,
     * INDIRECT) or does not resolve. SUMIF and AVERAGEIF size their third argument by it (#733).
     */
    def range(
      qualifier: Option[SheetName],
      name: String,
      current: SheetName
    ): Option[(SheetName, CellRange)] =
      ranges.getOrElseUpdate(
        (DependencyGraph.resolutionKey(name), qualifier.map(fold), fold(current)),
        wb(current).toOption.flatMap { sheet =>
          Evaluator
            .resolveRangeLocation(TExpr.RangeLocation.Name(name, qualifier), sheet, Some(wb))
            .toOption
            .map((target, r) => (target.name, r))
        }
      )

    /** Whether `wb` defines any name: without one, no formula reads through a name. */
    def nonEmpty: Boolean = wb.metadata.definedNames.nonEmpty

    /** The lower-cased names `wb` defines, for a text prefilter. */
    lazy val keys: Set[String] =
      wb.metadata.definedNames.iterator.map(_.name.toLowerCase(java.util.Locale.ROOT)).toSet

    def reads(
      qualifier: Option[SheetName],
      name: String,
      current: SheetName,
      target: SheetName,
      bounds: Option[CellRange]
    ): Set[CellRange] =
      Evaluator
        .lookupDefinedNameAt(wb, positions.get(fold(qualifier.getOrElse(current))), name)
        .fold(Set.empty[CellRange]) { definition =>
          val from = Evaluator.definedNameScope(wb, definition).fold(current)(_.name)
          val key = (definition, fold(from), fold(target))
          memo.get(key) match
            case Some(found) => found
            case None if open.contains(key) || open.sizeIs >= NameWalk.MaxDepth => Set.empty
            case None =>
              open += key
              val body = bodies.getOrElseUpdate(
                definition,
                FormulaParser.parse(definition.formula).toOption
              )
              val read = body.fold(Set.empty[CellRange]) { expr =>
                Reader(bounds, Some(target), fold(from) == fold(target), Some(this), Some(from))
                  .readRanges(expr)
              }
              open -= key
              memo.update(key, read)
              read
        }

    private def fold(sheet: SheetName): String = sheet.value.toLowerCase(java.util.Locale.ROOT)

  private def referenceDefinition(expr: TExpr[?]): Boolean = expr match
    // Literal range bindings are substituted by the parser and skipped by LET evaluation.
    case _: TExpr.RangeRef | _: TExpr.SheetRange => true
    case _ => Evaluator.denotesReference(expr)
