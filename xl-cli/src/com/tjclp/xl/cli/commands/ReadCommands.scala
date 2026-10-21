package com.tjclp.xl.cli.commands

import cats.effect.IO
import cats.implicits.*
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, RefType}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.cli.contract.{CliError, CliException}
import com.tjclp.xl.cli.helpers.{Resolve, SheetResolver, ValueParser}
import com.tjclp.xl.cli.output.{Escape, Format, JsonRenderer}
import com.tjclp.xl.error.{XLError, XLException}
import com.tjclp.xl.formula.{DependencyGraph, FormulaParser, SheetEvaluator, TExpr}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.graph.ValueDependencies
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Read command handlers that are not [[com.tjclp.xl.cli.read.Reads]] queries: `bounds` over a
 * loaded workbook, `eval` and `evala`. The record-based verbs (`view`, `cell`, `search`, `stats`,
 * `filter`) live in `com.tjclp.xl.cli.read` (W2.4).
 */
object ReadCommands:

  // --- Typed failures (ADR-017 §2.3): every refusal names its code ------------------------------

  /** A wrong flag or argument combination: `USAGE` (exit 2). */
  private def usage(message: String): CliException = CliException(CliError.usage(message, None))

  /** A ref of the wrong shape (a range where one cell is needed): `INVALID_REFERENCE`. */
  private def invalidRef(reason: String): CliException =
    CliException(Resolve.invalidReference(reason))

  /** A formula that does not parse: `FORMULA_ERROR` with the parser's own message. */
  private def unparseable(error: ParseError, formula: String): XLException =
    XLException(ParseError.toXLError(error, formula))

  /**
   * A cycle found while ordering the closure: the evaluator's own error, with its code. A path
   * longer than [[MaxCyclePath]] keeps its first [[CycleHead]] cells and its last [[CycleTail]] and
   * counts the cells between. A 20,000-cell chain would otherwise make a 169 KB message.
   */
  private def cyclic(error: EvalError, formula: String): XLException =
    error match
      case EvalError.CircularRef(cycle) if cycle.size > MaxCyclePath =>
        withReason(error, formula, s"Circular reference detected: ${elided(cycle)}")
      case _ => XLException(EvalError.toXLError(error, Some(formula)))

  /**
   * `error` as the evaluator reports it, with only the formula error's reason replaced, so any
   * structured field `toXLError` carries survives.
   */
  private def withReason(error: EvalError, formula: String, reason: String): XLException =
    XLException(EvalError.toXLError(error, Some(formula)) match
      case reported @ XLError.FormulaError(_, _) => reported.copy(reason = reason)
      case other => other)

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

  /**
   * Show used range of current sheet.
   */
  def bounds(wb: Workbook, sheetOpt: Option[Sheet]): IO[String] =
    SheetResolver.requireSheet(wb, sheetOpt, "bounds").map { sheet =>
      val name = sheet.name.value
      val usedRange = sheet.usedRange
      val cellCount = sheet.cells.size
      usedRange match
        case Some(range) =>
          val rowCount = range.end.row.index0 - range.start.row.index0 + 1
          val colCount = range.end.col.index0 - range.start.col.index0 + 1
          s"""Sheet: $name
             |Used range: ${range.toA1}
             |Rows: ${range.start.row.index1}-${range.end.row.index1} ($rowCount total)
             |Columns: ${range.start.col.toLetter}-${range.end.col.toLetter} ($colCount total)
             |Non-empty: $cellCount cells""".stripMargin
        case None =>
          s"""Sheet: $name
             |Used range: (empty)
             |Non-empty: 0 cells""".stripMargin
    }

  /**
   * Evaluate formula without modifying sheet.
   *
   * For constant formulas (e.g., =1+1, =PI()*2) that don't reference any cells, both --file and
   * --sheet are optional. For formulas that reference cells, a file and sheet must be provided.
   *
   * Without `at` the formula has no cell of its own and evaluates as typed into a new Excel 365
   * cell (array mode, its top-left value). With `at` (GH-715) it evaluates as the plain cell at
   * that ref would — the legacy formula `putf` writes, references implicitly intersected with its
   * row or column — on the sheet THE sheet rule gives the ref; nothing is written.
   */
  def eval(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    formulaStr: String,
    at: Option[String],
    overrides: List[String]
  ): IO[String] =
    evalResult(wb, sheetOpt, formulaStr, at, overrides).map { (formula, position, result) =>
      Format.evalSuccess(formula, position, result, overrides)
    }

  /**
   * `eval` as data (`eval --json`): `{formula, at, result: {type, value, formatted}, overrides}` —
   * the normalized formula (leading `=`), the `--at` cell as evaluated (`Data!D5`, or `null` when
   * positionless), the value in the `view --format json` cell shape, and the `--with` overrides as
   * given. JSON TEXT, not a ujson tree: the number lexeme is exactly what [[JsonRenderer]] prints
   * (`12345678901234567` stays so), for the runner to splice as `Payload.Raw`.
   */
  def evalData(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    formulaStr: String,
    at: Option[String],
    overrides: List[String]
  ): IO[String] =
    evalResult(wb, sheetOpt, formulaStr, at, overrides).map { (formula, position, result) =>
      s"{\"formula\": ${Escape.json(formula)}, " +
        s"\"at\": ${position.fold("null")(Escape.json)}, " +
        s"\"result\": ${JsonRenderer.valueJson(result, NumFmt.General)}, " +
        s"\"overrides\": ${overridesJson(overrides)}}"
    }

  /** The `--with` list as a JSON array of strings. */
  private def overridesJson(overrides: List[String]): String =
    ujson.write(ujson.Arr.from(overrides.map(ujson.Str.apply)))

  /** No file was given, but the formula needs one. */
  private def fileRequired: CliException =
    usage("Formula references cells but no file provided. Use --file to specify an Excel file.")

  /**
   * The evaluation both renderings share: the normalized formula, the `--at` cell as evaluated
   * (qualified by its sheet when there is a workbook), and the value.
   */
  private def evalResult(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    formulaStr: String,
    at: Option[String],
    overrides: List[String]
  ): IO[(String, Option[String], CellValue)] =
    val formula = if formulaStr.startsWith("=") then formulaStr else s"=$formulaStr"
    at match
      case None =>
        positionless(wb, sheetOpt, formula, overrides).map((formula, None, _))
      case Some(atStr) =>
        positioned(wb, sheetOpt, formula, atStr, overrides).map { (position, result) =>
          (formula, Some(position), result)
        }

  /**
   * Whether the formula reads cells: `(unqualified, any)`. `--with` counts as an unqualified read.
   * GH-197: `containsCellReferences`, not `extractDependencies`, so a full-column range is not
   * enumerated. GH-210: a formula whose refs are all qualified (`=SUM(Revenue!B2:B5)`) needs a
   * workbook but no ambient sheet. A parse failure reads as neither, so the constant branch raises
   * the parser's diagnostic.
   */
  private def cellReads(formula: String, overrides: List[String]): (Boolean, Boolean) =
    FormulaParser.parse(formula) match
      case Right(expr) =>
        (
          DependencyGraph.containsUnqualifiedCellReferences(expr) || overrides.nonEmpty,
          DependencyGraph.containsCellReferences(expr) || overrides.nonEmpty
        )
      case Left(_) => (false, false)

  /** `eval` without `--at`: array mode, the top-left value. */
  private def positionless(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    formula: String,
    overrides: List[String]
  ): IO[CellValue] =
    val (needsSheet, hasAnyCellRefs) = cellReads(formula, overrides)
    if (needsSheet || hasAnyCellRefs) && wb.sheets.isEmpty then IO.raiseError(fileRequired)
    else if needsSheet then
      SheetResolver
        .requireSheet(wb, sheetOpt, "eval")
        .flatMap(closureEval(wb, _, formula, overrides, None))
    else if hasAnyCellRefs then
      // GH-210: every ref is qualified — any sheet is the ambient context, since the evaluator
      // resolves cross-sheet refs by name. The workbook is non-empty (checked above).
      sheetOpt
        .orElse(wb.sheets.headOption)
        .fold(IO.raiseError(fileRequired))(closureEval(wb, _, formula, overrides, None))
    else constantEval(wb, sheetOpt.getOrElse(Sheet("_eval")), formula, None)

  /**
   * `eval --at` (GH-715): the formula as the plain cell at `atStr`. With a workbook the ref follows
   * THE sheet rule — a qualifier names the sheet, else the default, else the only sheet, else
   * `SHEET_REQUIRED` — and the formula evaluates on that sheet even when it reads no cell (`ROW()`
   * is the position's row either way). Without one, only a cell-free formula at an unqualified cell
   * can be answered. A range or an unparseable ref is `INVALID_REFERENCE`.
   */
  private def positioned(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    formula: String,
    atStr: String,
    overrides: List[String]
  ): IO[(String, CellValue)] =
    if wb.sheets.isEmpty then
      IO.fromEither(Resolve.ref(atStr).left.map(CliException(_))).flatMap {
        case (Some(name), _) =>
          IO.raiseError(
            usage(
              s"--at $atStr names sheet ${name.value} but no file provided. " +
                "Use --file to specify an Excel file."
            )
          )
        case (None, Resolve.Target.Range(_)) => IO.raiseError(notOneCell(atStr))
        case (None, Resolve.Target.Cell(cell)) =>
          val (_, hasAnyCellRefs) = cellReads(formula, overrides)
          if hasAnyCellRefs then IO.raiseError(fileRequired)
          else constantEval(wb, Sheet("_eval"), formula, Some(cell)).map((cell.toA1, _))
      }
    else
      SheetResolver.resolveRef(wb, sheetOpt, atStr, "eval").flatMap {
        case (sheet, Left(cell)) =>
          overrides.find(overridesCell(_, sheet.name, cell)) match
            case Some(clash) => IO.raiseError(placementOverridden(clash, atStr))
            case None =>
              closureEval(wb, sheet, formula, overrides, Some(cell))
                .map((RefType.QualifiedCell(sheet.name, cell).toA1, _))
        case (_, Right(_)) => IO.raiseError(notOneCell(atStr))
      }

  /**
   * Whether the `--with` override `override_` sets `cell` on the sheet named `sheet`, as
   * [[parseOverride]] reads it. A malformed override is not a match; [[applyOverrides]] reports it.
   */
  private def overridesCell(override_ : String, sheet: SheetName, cell: ARef): Boolean =
    parseOverride(override_, sheet).exists((ref, _) => ref == cell)

  /**
   * One `--with` override, `ref=value`, as the cell it sets on the sheet named `sheet` and its
   * value. The ref is a single cell, unqualified or qualified with that sheet's name in any case,
   * as Excel compares sheet names. A range or another sheet's cell is `INVALID_REFERENCE`; a
   * missing `=` or an empty value is `USAGE`.
   */
  private def parseOverride(
    override_ : String,
    sheet: SheetName
  ): Either[CliException, (ARef, CellValue)] =
    override_.split("=", 2) match
      case Array(refStr, valueStr) if valueStr.trim.nonEmpty =>
        // INVALID_REFERENCE: the parser's own text, a range, or another sheet's cell
        Resolve.ref(refStr.trim).left.map(CliException(_)).flatMap {
          case (None, Resolve.Target.Cell(ref)) =>
            Right((ref, ValueParser.parseValue(valueStr.trim)))
          case (Some(own), Resolve.Target.Cell(ref)) if own.value.equalsIgnoreCase(sheet.value) =>
            Right((ref, ValueParser.parseValue(valueStr.trim)))
          case (Some(other), Resolve.Target.Cell(_)) =>
            Left(
              invalidRef(
                s"Cross-sheet override not supported: ${refStr.trim}. " +
                  s"Eval operates on ${sheet.value}, not ${other.value}"
              )
            )
          case (_, Resolve.Target.Range(_)) =>
            Left(invalidRef(s"Override requires single cell, not range: ${refStr.trim}"))
        }
      case Array(refStr, _) =>
        Left(usage(s"Empty value for override: ${refStr.trim}. Use ref=value (e.g., B5=1000)"))
      case _ =>
        Left(usage(s"Invalid override format: $override_. Use ref=value (e.g., B5=1000)"))

  /**
   * `--with` on the `--at` cell: the formula IS that cell, so an override of it could never be
   * read. Refused up front as `USAGE`, before it could surface as a circular reference.
   */
  private def placementOverridden(override_ : String, atStr: String): CliException =
    CliException(
      CliError.usage(
        s"--with ${override_.trim} sets the --at cell $atStr, which the formula replaces",
        Some("drop that override, or override the cells the formula reads")
      )
    )

  /** `--at` names the one cell the formula would sit in: a range is `INVALID_REFERENCE`. */
  private def notOneCell(atStr: String): CliException =
    invalidRef(s"--at needs a single cell, not a range: $atStr")

  /** A formula that reads no cell: no closure to evaluate, no workbook needed. */
  private def constantEval(
    wb: Workbook,
    sheet: Sheet,
    formula: String,
    position: Option[ARef]
  ): IO[CellValue] =
    // For constant formulas, workbook context is not needed
    val wbOpt = if wb.sheets.nonEmpty then Some(wb) else None
    for
      // the parser's own diagnostic, as the referencing branches raise it
      _ <- IO.fromEither(FormulaParser.parse(formula).left.map(unparseable(_, formula)))
      result <- IO.fromEither(
        SheetEvaluator
          .evaluateFormula(sheet)(formula, workbook = wbOpt, currentCell = position)
          .left
          .map(XLException(_))
      )
    yield result

  /**
   * The formula on `sheet` after its precedent closure ([[precedentsEvaluated]]): positionless, or
   * as the plain cell at `position`.
   */
  private def closureEval(
    wb: Workbook,
    sheet: Sheet,
    formula: String,
    overrides: List[String],
    position: Option[ARef]
  ): IO[CellValue] =
    precedentsEvaluated(wb, sheet, formula, overrides, position).flatMap { evaluated =>
      IO.fromEither(
        evaluated.checked(
          SheetEvaluator.evaluateFormulaUsing(
            evaluated.sheet,
            formula,
            evaluated.evaluator,
            Some(evaluated.book),
            currentCell = position
          ),
          formula
        )
      )
    }

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
     * `result` as the CLI raises it. If the evaluation read the placement, by any path and even
     * when IFERROR swallowed the guard's failure, the plain cell would be circular: a read the
     * static check could not see, computed at run time by INDIRECT, OFFSET or a defined name, or
     * made by another sheet's formula. The error is the evaluator's for the one-cell cycle
     * `CircularRef(List(cell))`, with only its reason text replaced: the guard sees the read, not
     * the path that made it, so any structured cycle field names the placement alone.
     */
    def checked[A](result: XLResult[A], formula: String): Either[XLException, A] =
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
        case None => result.left.map(XLException(_))

  /**
   * `sheet` with the `--with` overrides applied and only the formula cells in the formula's
   * dependency closure evaluated, in dependency order, through one evaluator the caller then
   * evaluates the formula with (GH-695: `x#` reads the arrays the fold computed):
   *   - formula chains work (`=C1` where C1=B1+50, B1=A1*2);
   *   - overrides propagate through chains (TJC-698): A1=200 → B1=400 → C1=450;
   *   - only the O(k) closure is evaluated, not the O(n) sheet. GH-197: bounded extraction, so a
   *     full-column range does not enumerate 1M+ cells. A sheet with no cells has no used range to
   *     bound by, but it also has no precedent to evaluate, so nothing is extracted.
   *
   * With `at` (GH-715) the formula is the cell there: a closure that reads `at` back is the
   * circular reference the plain cell would be, raised as the recalc cycle is, not `at`'s old
   * content. [[cycleThrough]] names the path of a static read; the placed evaluator catches a read
   * made at run time ([[Evaluated.checked]]). Its closure is the one [[valueReadClosure]] walks.
   */
  private def precedentsEvaluated(
    wb: Workbook,
    sheet: Sheet,
    formula: String,
    overrides: List[String],
    at: Option[ARef]
  ): IO[Evaluated] =
    for
      withOverrides <- overridden(wb, sheet, overrides)
      (tempSheet, tempBook) = withOverrides
      expr <- IO.fromEither(FormulaParser.parse(formula).left.map(unparseable(_, formula)))
      bounds = tempSheet.usedRange
      // Positioned or not, only the formula's value-read closure is analysed: never the whole
      // sheet, which a running total (B{n} = SUM($A$1:A{n})) makes quadratic to expand
      reads = ValueDependencies.Reader(bounds, Some(tempSheet.name))
      targetDeps = bounds.fold(Set.empty[ARef])(_ => reads.reads(expr))
      closure = valueReadClosure(tempSheet, reads, targetDeps)
      _ <- at.flatMap(cycleThrough(tempSheet.name, closure, expr, targetDeps, _)) match
        case Some(cycle) => IO.raiseError(cyclic(EvalError.CircularRef(cycle), formula))
        case None => IO.unit
      graph = closure.graph
      allDeps = closure.reached
      formulaDeps = allDeps.filter(ref =>
        tempSheet(ref).value match
          case _: CellValue.Formula => true
          case _ => false
      )
      evalOrder <- IO.fromEither(
        if formulaDeps.isEmpty then scala.util.Right(List.empty[ARef])
        else
          DependencyGraph
            .topologicalSort(graph)
            .map(_.filter(formulaDeps.contains))
            .left
            .map(cyclic(_, formula))
      )
      // GH-715: the placed formula sits, uncached, at `at` in the sheet and in the workbook, and
      // the placed evaluator guards that cell, so a read the static check cannot see (INDIRECT,
      // OFFSET, a defined name, another sheet's formula) fails instead of reading its old content
      placedSheet = at.fold(tempSheet)(SheetEvaluator.placed(tempSheet, _, formula))
      // The book holds this sheet as evaluated, so a formula on another sheet that reads back
      // into it sees the `--with` overrides, with `--at`, without it and in evala alike
      book = tempBook.put(placedSheet)
      placement = at.map(cell =>
        (cell, SheetEvaluator.placement(formula, wb, tempSheet.name, cell))
      )
      evaluator = placement.fold(SheetEvaluator.spillTrackingEvaluator(formula, wb))(
        _._2.evaluator
      )
      start = Evaluated(placedSheet, evaluator, book, placement)
      evaluated <- evalOrder.foldLeft(IO.pure(start)) { (evaluatedIO, ref) =>
        evaluatedIO.flatMap { e =>
          IO.fromEither(
            e.checked(
              SheetEvaluator.evaluateCellUsing(e.sheet, ref, evaluator, Some(e.book)),
              formula
            ).map { value =>
              // the book keeps step with the sheet, so a read qualified with this sheet's name, or
              // one back from another sheet, sees what the fold computed
              val threaded = SheetEvaluator.threadComputed(e.sheet, ref, value)
              e.copy(sheet = threaded, book = e.book.put(threaded))
            }
          )
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
                    val deps = reads.reads(expr)
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
    def reads(e: TExpr[?]): Boolean = only.reads(e).contains(at)
    def readsBack(ref: ARef): Boolean = closure.formulas.get(ref).exists(reads)
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
    if reads(expr) then Some(List(at, at))
    else
      val roots = (targetDeps - at).toVector
      search(roots, roots.map(_ -> at).toMap)

  /**
   * Evaluate array formula and display result as table.
   *
   * Array formulas like TRANSPOSE return multiple values that are displayed as a grid. Unlike
   * regular eval which returns a single value, evala shows the full spilled array.
   */
  def evalArray(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    formulaStr: String,
    targetRefOpt: Option[String],
    overrides: List[String]
  ): IO[String] =
    evalArrayResult(wb, sheetOpt, formulaStr, targetRefOpt, overrides).map {
      (formula, updatedSheet, spillRange) =>
        Format.evalArraySuccess(formula, updatedSheet, spillRange, overrides)
    }

  /**
   * `evala` as data (`evala --json`): `{formula, spillRange, result, overrides}` where `result` is
   * the spilled grid in the exact `view --format json` shape (`{sheet, range, rows}`) — as the text
   * [[JsonRenderer]] prints it, so no number is re-parsed (`Payload.Raw`).
   */
  def evalArrayData(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    formulaStr: String,
    targetRefOpt: Option[String],
    overrides: List[String]
  ): IO[String] =
    evalArrayResult(wb, sheetOpt, formulaStr, targetRefOpt, overrides).map {
      (formula, updatedSheet, spillRange) =>
        s"{\"formula\": ${Escape.json(formula)}, " +
          s"\"spillRange\": \"${spillRange.toA1}\", " +
          s"\"result\": ${JsonRenderer.renderRange(updatedSheet, spillRange)}, " +
          s"\"overrides\": ${overridesJson(overrides)}}"
    }

  /** The evaluation both renderings share: the normalized formula, the spilled sheet, the range. */
  private def evalArrayResult(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    formulaStr: String,
    targetRefOpt: Option[String],
    overrides: List[String]
  ): IO[(String, Sheet, CellRange)] =
    val formula = if formulaStr.startsWith("=") then formulaStr else s"=$formulaStr"

    // Array formulas always need a sheet context
    if wb.sheets.isEmpty then
      IO.raiseError(
        usage("Array formula evaluation requires a file. Use --file to specify an Excel file.")
      )
    else
      for
        anchored <- arrayAnchor(wb, sheetOpt, targetRefOpt)
        (sheet, originRef) = anchored
        // positionless: evala anchors its display, it does not place the formula in a cell
        precedents <- precedentsEvaluated(wb, sheet, formula, overrides, None)
        evalSheet = precedents.sheet
        evaluator = precedents.evaluator
        result <- IO.fromEither(
          SheetEvaluator
            .evaluateArrayFormulaUsing(
              evalSheet,
              formula,
              originRef,
              evaluator,
              Some(precedents.book)
            )
            .left
            .map(XLException(_))
        )
        (updatedSheet, spillRange) = result
      yield (formula, updatedSheet, spillRange)

  /**
   * The sheet `evala` evaluates on and the cell its display is anchored at. `--at` follows THE
   * sheet rule as `eval --at` does (#727): a qualifier names the sheet, else the default, else the
   * only sheet, else `SHEET_REQUIRED`; a range or an unparseable ref is `INVALID_REFERENCE`.
   * Without `--at` the anchor is a virtual cell far from data, Z1000.
   */
  private def arrayAnchor(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    targetRefOpt: Option[String]
  ): IO[(Sheet, ARef)] =
    targetRefOpt match
      case Some(atStr) =>
        SheetResolver.resolveRef(wb, sheetOpt, atStr, "evala").flatMap {
          case (sheet, Left(cell)) => IO.pure((sheet, cell))
          case (_, Right(_)) => IO.raiseError(notOneCell(atStr))
        }
      case None =>
        SheetResolver.requireSheet(wb, sheetOpt, "evala").map((_, ARef.from0(25, 999)))

  // ==========================================================================
  // Private helpers
  // ==========================================================================

  /**
   * `sheet` with the `--with` overrides applied, and `wb` holding it, with every cached formula in
   * the book that reads an overridden cell, directly or through other formulas on any sheet, made
   * uncached: the cache predates the override. However the formula is then read (from the closure,
   * at run time through `INDIRECT("B1")`, or back from another sheet as `Other!A1` holding
   * `=Data!B1*2`), it evaluates with the overrides.
   *
   * This is the one step of `eval` (and of `--at`) that is not O(k) in the formula's closure: with
   * overrides, every formula on the sheet is parsed once to find its value reads
   * ([[ValueDependencies]], so `ROW(A1)` does not go stale when A1 changes), and every formula on
   * another sheet that names a sheet a stale cell sits on. The closure alone would not do: the
   * stale formulas it cannot see are the ones read only at run time or back from another sheet.
   * EvalCommandSpec pins both shapes, cached. Without overrides nothing is parsed.
   */
  private def overridden(
    wb: Workbook,
    sheet: Sheet,
    overrides: List[String]
  ): IO[(Sheet, Workbook)] =
    if overrides.isEmpty then IO.pure((sheet, wb))
    else
      applyOverrides(sheet, overrides).map { (applied, changed) =>
        val book = uncacheReaders(wb.put(applied), applied.name, changed)
        (book.sheets.find(_.name == applied.name).getOrElse(applied), book)
      }

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
   * cell's.
   *
   * The maps, sets and bit set below are local to one call and never escape it: the result is a
   * new, immutable workbook.
   */
  private[cli] def uncacheReaders(book: Workbook, home: SheetName, changed: Set[ARef]): Workbook =
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

    def readersOf(target: Int): Readers =
      val readers = Readers()
      val name = sheets(target).name
      val bounds = sheets(target).usedRange
      // a reference qualified with the target spells its name, perhaps quoted, then `!`, in any
      // case; a name may hide one, so a formula mentioning a name is read too
      val qualified = java.util.regex.Pattern.compile(
        "'?" + java.util.regex.Pattern.quote(name.value.replace("'", "''")) + "'?!",
        java.util.regex.Pattern.CASE_INSENSITIVE | java.util.regex.Pattern.UNICODE_CASE
      )
      sheets.indices.foreach { at =>
        val reader = ValueDependencies.Reader(
          bounds,
          Some(name),
          localReads = at == target,
          names = Some(names),
          current = Some(sheets(at).name)
        )
        sheets(at).cells.foreach { (ref, cell) =>
          cell.value match
            // GH-430: a data table's cells are pinned value sources
            case CellValue.Formula(_, _, _: FormulaKind.DataTable) => ()
            case CellValue.Formula(text, _, _)
                if at == target || qualified.matcher(text).find() ||
                  mentionsName((at, ref), text) =>
              parse((at, ref), text).foreach { expr =>
                val read = reader.readRanges(expr).toVector
                if read.nonEmpty then readers.register(idOf((at, ref)), read)
              }
            case _ => ()
        }
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

  /** A token of formula text that could be a defined name, for [[uncacheReaders]]'s prefilter. */
  private val NameToken = """[\p{L}_\\][\p{L}\p{N}_.\\?]*""".r

  /** The rows in one block of [[uncacheReaders]]'s index. */
  private val StaleBlockRows = 1024

  /** The most blocks one range registers in; a wider range is checked against every cell. */
  private val StaleBlockLimit = 4096L

  /** `sheet` with the `--with` overrides applied, and the cells they set. */
  private def applyOverrides(sheet: Sheet, overrides: List[String]): IO[(Sheet, Set[ARef])] =
    IO.fromEither(
      overrides.foldLeft[Either[CliException, (Sheet, Set[ARef])]](Right((sheet, Set.empty))) {
        case (acc, override_) =>
          acc.flatMap { (s, changed) =>
            parseOverride(override_, sheet.name).map((ref, value) =>
              (s.put(ref, value), changed + ref)
            )
          }
      }
    )
