package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.graph.{DependencyGraph, ValueDependencies}
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * A closure evaluation failure, with the existing library diagnostic and, for a detected cycle, its
 * full path. A run-time guard knows only the placement, so its path has that cell alone.
 * Diagnostics may abbreviate long paths; `cycle` never does.
 */
final case class FormulaEvaluationError(error: XLError, cycle: Option[List[ARef]] = None)

/** Single-use evaluation state stays behind the public SheetEvaluator closure APIs. */
private[formula] object FormulaEvaluation:
  private type Result[A] = Either[FormulaEvaluationError, A]

  def scalar(
    wb: Workbook,
    sheet: Sheet,
    formula: String,
    overrides: Map[ARef, CellValue],
    at: Option[ARef]
  ): Result[CellValue] =
    val normalized = normalize(formula)
    precedentsEvaluated(wb, sheet, normalized, overrides, at).flatMap { e =>
      e.checked(
        SheetEvaluator
          .evaluateFormulaUsing(e.sheet, normalized, e.evaluator, Some(e.book), currentCell = at),
        normalized
      )
    }

  def array(
    wb: Workbook,
    sheet: Sheet,
    formula: String,
    origin: ARef,
    overrides: Map[ARef, CellValue]
  ): Result[(Sheet, CellRange)] =
    val normalized = normalize(formula)
    precedentsEvaluated(wb, sheet, normalized, overrides, None).flatMap { e =>
      e.checked(
        SheetEvaluator
          .evaluateArrayFormulaUsing(e.sheet, normalized, origin, e.evaluator, Some(e.book)),
        normalized
      )
    }

  private def normalize(formula: String): String =
    if formula.startsWith("=") then formula else s"=$formula"

  /**
   * A cycle found while ordering the closure: the evaluator's own error, with its code. A path
   * longer than [[MaxCyclePath]] keeps its first [[CycleHead]] cells and its last [[CycleTail]] and
   * counts the cells between. A 20,000-cell chain would otherwise make a 169 KB message.
   */
  private def cyclic(error: EvalError, formula: String): FormulaEvaluationError =
    error match
      case EvalError.CircularRef(cycle) if cycle.size > MaxCyclePath =>
        withReason(error, formula, s"Circular reference detected: ${elided(cycle)}")
      case _ => FormulaEvaluationError(EvalError.toXLError(error, Some(formula)), cycleOf(error))

  /**
   * `error` as the evaluator reports it, with only the formula error's reason replaced, so any
   * structured field `toXLError` carries survives.
   */
  private def withReason(
    error: EvalError,
    formula: String,
    reason: String
  ): FormulaEvaluationError =
    FormulaEvaluationError(
      (EvalError.toXLError(error, Some(formula)) match
        case reported @ XLError.FormulaError(_, _) => reported.copy(reason = reason)
        case other => other
      ),
      cycleOf(error)
    )

  /** A cycle path longer than [[MaxCyclePath]], its middle counted. */
  private def elided(cycle: List[ARef]): String =
    val between = cycle.size - CycleHead - CycleTail
    val middle = "… (%,d more) …".formatLocal(java.util.Locale.ROOT, between)
    (cycle.take(CycleHead).map(_.toA1) ++ (middle :: cycle.takeRight(CycleTail).map(_.toA1)))
      .mkString(" → ")

  /** The longest cycle path [[cyclic]] prints in full. */
  private val MaxCyclePath = 10

  /** The cells an elided cycle path keeps from its start. */
  private val CycleHead = 3

  /**
   * The cells an elided cycle path keeps from its end: the cell that reads the start, the start.
   */
  private val CycleTail = 2

  private def cycleOf(error: EvalError): Option[List[ARef]] = error match
    case EvalError.CircularRef(path) => Some(path)
    case _ => None

  /**
   * The closure [[precedentsEvaluated]] evaluated: `sheet` with it threaded, the `evaluator` and
   * `book` the formula then evaluates with, and under `--at` (GH-715) the `placement` whose guard
   * saw any read of the `--at` cell.
   */
  private final case class Evaluated(
    sheet: Sheet,
    evaluator: Evaluator,
    book: Workbook,
    placement: Option[(ARef, SheetEvaluator.Placement)]
  ):
    /**
     * Check the guard before reporting a result or an ordinary evaluation failure. If the
     * evaluation read the placement, by any path and even when IFERROR swallowed the guard's
     * failure, the plain cell would be circular: a read the static check could not see, computed at
     * run time by INDIRECT, OFFSET or a defined name, or made by another sheet's formula. The error
     * is the evaluator's for the one-cell cycle `CircularRef(List(cell))`, with only its reason
     * text replaced: the guard sees the read, not the path that made it, so any structured cycle
     * field names the placement alone.
     */
    def checked[A](result: XLResult[A], formula: String): Result[A] =
      placement.filter(_._2.wasRead) match
        case Some((cell, _)) =>
          Left(
            withReason(
              EvalError.CircularRef(List(cell)),
              formula,
              s"Circular reference detected: the formula reads its own placement ${cell.toA1} " +
                "at run time (e.g. through INDIRECT, OFFSET, a defined name or another sheet's " +
                "formula)"
            )
          )
        case None => result.left.map(FormulaEvaluationError(_))

  /**
   * Apply typed overrides and invalidate their cached readers, then visit only the formula's
   * value-read closure. One evaluator carries spill results and the placement guard through the
   * entire fold and final evaluation. Thread each result into both the sheet and workbook so
   * own-sheet qualifiers and cross-sheet read-backs see the same temporary values.
   */
  private def precedentsEvaluated(
    wb: Workbook,
    sheet: Sheet,
    formula: String,
    overrides: Map[ARef, CellValue],
    at: Option[ARef]
  ): Result[Evaluated] =
    for
      _ <-
        if at.exists(overrides.contains) then
          Left(
            FormulaEvaluationError(
              XLError.FormulaError(
                formula,
                s"An override sets the placement cell ${at.fold("")(_.toA1)}, which the formula replaces"
              )
            )
          )
        else Right(())
      applied = overrides.foldLeft(sheet) { case (s, (ref, value)) => s.put(ref, value) }
      tempBook =
        if overrides.isEmpty then wb.put(applied)
        else uncacheReaders(wb.put(applied), applied.name, overrides.keySet)
      tempSheet = tempBook.sheets.find(_.name == applied.name).getOrElse(applied)
      expr <- FormulaParser
        .parse(formula)
        .left
        .map(e => FormulaEvaluationError(ParseError.toXLError(e, formula)))
      bounds = tempSheet.usedRange
      reads = ValueDependencies.Reader(bounds, Some(tempSheet.name))
      targetDeps = bounds.fold(Set.empty[ARef])(_ => reads.reads(expr, at))
      closure = valueReadClosure(tempSheet, reads, targetDeps)
      _ <- at.flatMap(cycleThrough(tempSheet.name, closure, expr, targetDeps, _)) match
        case Some(cycle) => Left(cyclic(EvalError.CircularRef(cycle), formula))
        case None => Right(())
      graph = closure.graph
      formulaDeps = closure.reached.filter(ref =>
        tempSheet(ref).value match
          case _: CellValue.Formula => true
          case _ => false
      )
      evalOrder <-
        if formulaDeps.isEmpty then Right(List.empty[ARef])
        else
          DependencyGraph
            .topologicalSort(graph)
            .map(_.filter(formulaDeps.contains))
            .left
            .map(cyclic(_, formula))
      placedSheet = at.fold(tempSheet)(SheetEvaluator.placed(tempSheet, _, formula))
      book = tempBook.put(placedSheet)
      placement = at.map(cell =>
        (cell, SheetEvaluator.placement(formula, book, tempSheet.name, cell))
      )
      // Spill detection reads the evaluation copy: the supplied sheet and any formula overrides
      // may carry `x#` readers the caller's workbook does not.
      evaluator = placement.fold(SheetEvaluator.spillTrackingEvaluator(formula, book))(
        _._2.evaluator
      )
      start = Evaluated(placedSheet, evaluator, book, placement)
      evaluated <- evalOrder.foldLeft[Result[Evaluated]](Right(start)) { (acc, ref) =>
        acc.flatMap { e =>
          e.checked(
            SheetEvaluator.evaluateCellUsing(e.sheet, ref, evaluator, Some(e.book)),
            formula
          ).map { value =>
            val threaded = SheetEvaluator.threadComputed(e.sheet, ref, value)
            e.copy(sheet = threaded, book = e.book.put(threaded))
          }
        }
      }
    yield evaluated

  /**
   * The closure [[valueReadClosure]] walks: `graph` has an edge from each formula it reached to the
   * cells whose values that formula reads (`dependents` only among those formulas, which is all
   * Kahn's order consults); `reached` is every cell reached, the formula's own reads included;
   * `formulas` holds each parse, so the cycle search does not parse again.
   */
  private final case class ValueReads(
    graph: DependencyGraph,
    reached: Set[ARef],
    formulas: Map[ARef, TExpr[?]]
  )

  /**
   * The dependency graph `eval` (positioned or not) and `evala` order by, built from `roots` (the
   * formula's own reads) out. `DependencyGraph.fromSheet` will not do, for three reasons:
   *   - its edges are structural. ROW(D5), ROWS(A:A), or a LET name that ROW inspects, make an edge
   *     to cells whose values are never read. A formula that inspects its own address, or a
   *     precedent that inspects the placement, would look circular, and the closure would be
   *     ordered by reads that do not happen. [[ValueDependencies]] keeps only the value reads.
   *   - it drops every sheet-qualified reference, but `Data!D5` in a cell on Data reads D5 as
   *     surely as `D5` does: the plain cell at D5 reading itself.
   *   - it analyses every formula on the sheet, where the plain cell needs only its closure. This
   *     walk parses and extracts each formula in the closure once, sharing one range cache, and
   *     touches nothing else.
   *
   * Like `fromSheet`, a data table's cells are value sources, not nodes (GH-430), and a formula
   * that does not parse is a node with no edges.
   */
  private def valueReadClosure(
    sheet: Sheet,
    reads: ValueDependencies.Reader,
    roots: Set[ARef]
  ): ValueReads =
    @scala.annotation.tailrec
    def walk(
      frontier: Set[ARef],
      reached: Set[ARef],
      edges: Map[ARef, Set[ARef]],
      formulas: Map[ARef, TExpr[?]]
    ): ValueReads =
      val visit = frontier -- reached
      if visit.isEmpty then
        val formulaEdges =
          edges.toList.flatMap((ref, deps) => deps.filter(edges.contains).map(_ -> ref))
        val dependents = formulaEdges.groupMap(_._1)(_._2).view.mapValues(_.toSet).toMap
        ValueReads(DependencyGraph(edges, dependents), reached, formulas)
      else
        val (nextEdges, nextFormulas, next) =
          visit.foldLeft((edges, formulas, Set.empty[ARef])) { case ((e, f, n), ref) =>
            sheet(ref).value match
              case CellValue.Formula(_, _, _: FormulaKind.DataTable) => (e, f, n)
              case CellValue.Formula(text, _, _) =>
                FormulaParser.parse(text) match
                  case Right(expr) =>
                    val deps = reads.reads(expr, Some(ref))
                    (e.updated(ref, deps), f.updated(ref, expr), n ++ deps)
                  case Left(_) => (e.updated(ref, Set.empty), f, n)
              case _ => (e, f, n)
          }
        walk(next, reached ++ visit, nextEdges, nextFormulas)
    walk(roots, Set.empty, Map.empty, Map.empty)

  /**
   * The cycle `expr` would close as the cell `at` on the sheet named `home`, `at` first and last:
   * `expr` reads `at`, or a formula in its closure does. Breadth first, so the shortest. Whether a
   * formula reads `at` is asked of its references themselves (bounded to `at` alone): `at` may lie
   * outside the used range that bounds the graph's range edges, as a total placed below its column
   * does. A reference qualified with `home` is a read of this sheet; another sheet's formulas are
   * not followed (the placed evaluator's run-time guard catches those reads).
   *
   * For `=B1+1 --at D5` with B1 = `=C1*2` and C1 = `=D5`: the formula does not read D5 itself, so
   * the search starts from its reads, B1, each seeded with `at` as its parent. B1 does not read D5;
   * its read C1 does, and walking the parents back gives `D5 → B1 → C1 → D5`.
   */
  private def cycleThrough(
    home: SheetName,
    closure: ValueReads,
    expr: TExpr[?],
    targetDeps: Set[ARef],
    at: ARef
  ): Option[List[ARef]] =
    val only = ValueDependencies.Reader(Some(CellRange(at, at)), Some(home))
    def reads(e: TExpr[?], from: ARef): Boolean = only.reads(e, Some(from)).contains(at)
    def readsBack(ref: ARef): Boolean = closure.formulas.get(ref).exists(reads(_, ref))
    // Invariant: every root's parent is `at`, and `at` itself has none, so walking parents back
    // from any reached cell ends at `at`
    @scala.annotation.tailrec
    def trace(ref: ARef, parent: Map[ARef, ARef], path: List[ARef]): List[ARef] =
      parent.get(ref) match
        case Some(from) if ref != at => trace(from, parent, ref :: path)
        case _ => at :: path
    @scala.annotation.tailrec
    def search(frontier: Vector[ARef], parent: Map[ARef, ARef]): Option[List[ARef]] =
      frontier.find(readsBack) match
        case Some(last) => Some(trace(last, parent, List(at)))
        case None =>
          val (next, reached) = frontier.foldLeft((Vector.empty[ARef], parent)) {
            case ((acc, seen), ref) =>
              closure.graph.dependencies.getOrElse(ref, Set.empty).foldLeft((acc, seen)) {
                case ((a, s), dep) =>
                  if dep == at || s.contains(dep) then (a, s) else (a :+ dep, s.updated(dep, ref))
              }
          }
          if next.isEmpty then None else search(next, reached)
    if reads(expr, at) then Some(List(at, at))
    else
      val roots = (targetDeps - at).toVector
      search(roots, roots.map(_ -> at).toMap)

  /**
   * `book` with every cached formula, on any sheet, that reads `changed` (cells of the sheet named
   * `home`), transitively across sheets, made uncached.
   *
   * Formulas are indexed by the ranges they read, never by those ranges' cells: a running total
   * (`B{n} = SUM($A$1:A{n})`) has a distinct range per row, and expanding them would cost the
   * square of the row count. One index per sheet a stale cell lands on, built when the first one
   * does: it holds that sheet's formulas with their unqualified and own-name reads, and another
   * sheet's formulas with their reads qualified with its name. A defined name read as a value reads
   * what its definition reads, resolved as the evaluator resolves it ([[ValueDependencies.Names]]),
   * so `rate*2` with `rate` naming `Data!$B$1` reads Data!B1. A formula on another sheet whose text
   * neither qualifies a reference with the sheet's name (`Data!`, `'Data'!`) nor holds a defined
   * name's token cannot read it and is not parsed for it. Each range registers in the grid blocks
   * it overlaps (one column by [[StaleBlockRows]] rows); a range spanning more than
   * [[StaleBlockLimit]] blocks is checked against every stale cell of its sheet instead. The walk
   * asks each stale cell's block for candidate formulas and keeps those with a range containing the
   * cell. Each formula is parsed at most once, whatever indexes it joins.
   *
   * A formula leaves every candidate set it is found in once it is stale, so later cells never
   * check it again. Without that, a running total of running totals (`C{n} = SUM($B$1:B{n})` over
   * stale B cells) re-checked each C reader for every stale B in its blocks, about n²/2 checks. Now
   * a formula is checked only while still cached, at most once per cell of a block it registered
   * in: [[StaleBlockRows]] checks per registration at worst, bounded by the indexes' own size, not
   * by the square of the sheet. A wide reader is checked against each stale cell, as before.
   *
   * Every formula [[DependencyGraph.dynamicCells]] lists (#763) is stale from the start, whatever
   * `changed` holds: its reads are known only to evaluation, so `SUM(OFFSET(Data!B1,0,0,2,1))`
   * reads Data!B2 though no range it names contains it. Its readers are then found like any stale
   * cell's. A sheet holding one then gets an index even when no override touches it, so the
   * formulas that could read another sheet are gathered once, in one pass over the book, by the
   * sheet names their text qualifies ([[qualifiedSheets]]): an index visits its own sheet's
   * formulas and those candidates. Indexing every sheet then costs one pass over the book's cells,
   * plus, per sheet, the formulas that mention a defined name (a name may hide any sheet), not a
   * pass over the book per sheet.
   *
   * The maps, sets and bit set below are local to one call and never escape it: the result is a
   * new, immutable workbook.
   */
  private[xl] def uncacheReaders(book: Workbook, home: SheetName, changed: Set[ARef]): Workbook =
    import scala.collection.mutable
    val sheets = book.sheets
    // every formula the walk meets, numbered on first sight: (sheet index, cell)
    val nodes = mutable.ArrayBuffer.empty[(Int, ARef)]
    val ids = mutable.HashMap.empty[(Int, ARef), Int]
    def idOf(node: (Int, ARef)): Int = ids.getOrElseUpdate(node, { nodes += node; nodes.size - 1 })
    val parsed = mutable.HashMap.empty[(Int, ARef), Option[TExpr[?]]]
    def parse(node: (Int, ARef), text: String): Option[TExpr[?]] =
      parsed.getOrElseUpdate(node, FormulaParser.parse(text).toOption)

    // the formulas reading one sheet: by (column, row block) the ones still cached with a range
    // overlapping it, each set shrinking as its formulas go stale
    final class Readers:
      val ranges = mutable.HashMap.empty[Int, Vector[CellRange]]
      val blocks = mutable.HashMap.empty[(Int, Int), mutable.Set[Int]]
      val wide = mutable.Set.empty[Int]
      def register(id: Int, read: Vector[CellRange]): Unit =
        ranges.update(id, read)
        read.foreach { r =>
          val firstBlock = r.rowStart.index0 / StaleBlockRows
          val lastBlock = r.rowEnd.index0 / StaleBlockRows
          if r.width.toLong * (lastBlock - firstBlock + 1) > StaleBlockLimit then wide += id
          else
            for
              col <- r.colStart.index0 to r.colEnd.index0
              block <- firstBlock to lastBlock
            do blocks.getOrElseUpdate((col, block), mutable.Set.empty[Int]) += id
        }

    // defined names, resolved as the evaluator resolves them: `rate*2` reads what `rate` names
    val names = ValueDependencies.Names(book)
    // whether a formula's text holds a token that could be a defined name, asked once per formula
    val namesAName = mutable.HashMap.empty[(Int, ARef), Boolean]
    def mentionsName(node: (Int, ARef), text: String): Boolean =
      names.nonEmpty && namesAName.getOrElseUpdate(
        node,
        NameToken
          .findAllIn(text)
          .exists(token => names.keys.contains(token.toLowerCase(java.util.Locale.ROOT)))
      )

    // every formula that could read a sheet other than its own, gathered in one pass over the
    // book: by each sheet name its text qualifies a reference with (folded, [[qualifiedSheets]]),
    // and every formula mentioning a defined name, which may hide one. An index for one target
    // then visits that sheet's formulas and these candidates, never every cell of the book, so
    // indexing every sheet a dynamic cell sits on costs O(cells), not O(sheets x cells)
    lazy val crossReaders: (Map[String, Vector[(Int, ARef, String)]], Vector[(Int, ARef, String)]) =
      val qualifying = mutable.HashMap.empty[String, mutable.ArrayBuffer[(Int, ARef, String)]]
      val naming = mutable.ArrayBuffer.empty[(Int, ARef, String)]
      sheets.indices.foreach { at =>
        sheets(at).cells.foreach { (ref, cell) =>
          cell.value match
            // GH-430: a data table's cells are pinned value sources
            case CellValue.Formula(_, _, _: FormulaKind.DataTable) => ()
            case CellValue.Formula(text, _, _) =>
              qualifiedSheets(text).foreach { sheet =>
                qualifying.getOrElseUpdate(sheet, mutable.ArrayBuffer.empty) += ((at, ref, text))
              }
              if mentionsName((at, ref), text) then naming += ((at, ref, text))
            case _ => ()
        }
      }
      (qualifying.view.mapValues(_.toVector).toMap, naming.toVector)

    def readersOf(target: Int): Readers =
      val readers = Readers()
      val name = sheets(target).name
      val bounds = sheets(target).usedRange
      val reader = mutable.HashMap.empty[Int, ValueDependencies.Reader]
      def readerOn(at: Int): ValueDependencies.Reader =
        reader.getOrElseUpdate(
          at,
          ValueDependencies.Reader(
            bounds,
            Some(name),
            localReads = at == target,
            names = Some(names),
            current = Some(sheets(at).name)
          )
        )
      def index(at: Int, ref: ARef, text: String): Unit =
        parse((at, ref), text).foreach { expr =>
          val read = readerOn(at).readRanges(expr, Some(ref)).toVector
          if read.nonEmpty then readers.register(idOf((at, ref)), read)
        }
      sheets(target).cells.foreach { (ref, cell) =>
        cell.value match
          // GH-430: a data table's cells are pinned value sources
          case CellValue.Formula(_, _, _: FormulaKind.DataTable) => ()
          case CellValue.Formula(text, _, _) => index(target, ref, text)
          case _ => ()
      }
      // another sheet's formula reads the target only through a reference qualified with its
      // name, perhaps behind a defined name
      val (qualifying, naming) = crossReaders
      (qualifying.getOrElse(foldSheet(name.value), Vector.empty) ++ naming).distinct.foreach {
        (at, ref, text) => if at != target then index(at, ref, text)
      }
      readers

    val indexes = mutable.HashMap.empty[Int, Readers]
    val stale = mutable.BitSet.empty
    // the candidates in `among` whose range contains `cell`; a formula already stale leaves it
    def hitsIn(readers: Readers, among: mutable.Set[Int], cell: ARef): List[Int] =
      among.filterInPlace(i => !stale.contains(i))
      among.iterator.filter(i => readers.ranges.get(i).exists(_.exists(_.contains(cell)))).toList
    @scala.annotation.tailrec
    def walk(frontier: List[(Int, ARef)]): Unit = frontier match
      case Nil => ()
      case (target, cell) :: rest =>
        val readers = indexes.getOrElseUpdate(target, readersOf(target))
        val block = readers.blocks.get((cell.col.index0, cell.row.index0 / StaleBlockRows))
        val hits = (block.fold(List.empty[Int])(hitsIn(readers, _, cell)) ++
          hitsIn(readers, readers.wide, cell)).distinct
        stale ++= hits
        walk(hits.map(nodes(_)) ::: rest)
    sheets.indexWhere(_.name == home) match
      case -1 => book
      case homeIndex =>
        // #763: a reader whose references are computed at run time (OFFSET, INDIRECT, a name whose
        // chain reaches one) reads what no static range shows, so it is stale under any override,
        // as recalculation treats it, and its own readers are walked from it
        val dynamic = DependencyGraph.dynamicCells(book).toList.flatMap { q =>
          sheets.indexWhere(_.name == q.sheet) match
            case -1 => Nil
            case at => List(at -> q.ref)
        }
        stale ++= dynamic.map(idOf)
        walk(changed.toList.map(homeIndex -> _) ::: dynamic)
        stale.toList.map(nodes(_)).groupMap(_._1)(_._2).foldLeft(book) { case (b, (at, refs)) =>
          b.put(refs.foldLeft(sheets(at)) { (s, ref) =>
            s(ref).value match
              case f @ CellValue.Formula(_, Some(_), _) => s.put(ref, f.copy(cachedValue = None))
              case _ => s
          })
        }

  /**
   * The sheet names `text` qualifies a reference with, each folded by [[foldSheet]]: `Data` for
   * `Data!B1`, `my sheet's` for `'My Sheet''s'!B1`. A prefilter for [[uncacheReaders]]: it may list
   * a name that qualifies nothing (text inside a string literal, `[1]Data!` of an external book),
   * never miss one the parser reads. An unquoted name runs back from its `!` to a character no
   * unquoted sheet name holds; a quoted one is matched whole, `''` an escaped quote.
   */
  private[xl] def qualifiedSheets(text: String): Set[String] =
    if text.indexOf('!') < 0 then Set.empty
    else
      // two scans, so a quoted match that spans a string literal (`"'"&Data!A1&"'!"`) never hides
      // an unquoted qualifier inside it
      (QuotedSheet.findAllMatchIn(text).map(_.group(1).replace("''", "'")) ++
        UnquotedSheet.findAllMatchIn(text).map(_.group(1))).map(foldSheet).toSet

  /**
   * A sheet name folded as `String.equalsIgnoreCase` compares (each character upper-cased, then
   * lower-cased), the comparison [[ValueDependencies.Reader]] makes.
   */
  private def foldSheet(name: String): String =
    name.map(c => Character.toLowerCase(Character.toUpperCase(c)))

  /** A quoted qualifier up to its `!`, group 1 its name still escaped. */
  private val QuotedSheet = """'((?:[^']|'')+)'!""".r

  /** An unquoted qualifier up to its `!`: no character an unquoted sheet name cannot hold. */
  private val UnquotedSheet = """([^\s'!(),+\-*/^&=<>%:;{}"\[\]]+)!""".r

  /** A token of formula text that could be a defined name, for [[uncacheReaders]]'s prefilter. */
  private val NameToken = """[\p{L}_\\][\p{L}\p{N}_.\\?]*""".r

  /** The rows in one block of [[uncacheReaders]]'s index. */
  private val StaleBlockRows = 1024

  /** The most blocks one range registers in; a wider range is checked against every cell. */
  private val StaleBlockLimit = 4096L
