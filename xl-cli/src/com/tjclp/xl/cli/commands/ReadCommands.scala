package com.tjclp.xl.cli.commands

import cats.effect.IO
import cats.implicits.*
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, RefType}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.{CliError, CliException}
import com.tjclp.xl.cli.helpers.{Resolve, SheetResolver, ValueParser}
import com.tjclp.xl.cli.output.{Escape, Format, JsonRenderer}
import com.tjclp.xl.error.XLException
import com.tjclp.xl.formula.{DependencyGraph, FormulaParser, SheetEvaluator, TExpr}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
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

  /** A cycle found while ordering the closure: the evaluator's own error, with its code. */
  private def cyclic(error: EvalError, formula: String): XLException =
    XLException(EvalError.toXLError(error, Some(formula)))

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
      // resolves cross-sheet refs by name. Safe: the workbook is non-empty (checked above).
      val sheet = (sheetOpt.orElse(wb.sheets.headOption): @unchecked) match
        case Some(s) => s
      closureEval(wb, sheet, formula, overrides, None)
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
          closureEval(wb, sheet, formula, overrides, Some(cell))
            .map((RefType.QualifiedCell(sheet.name, cell).toA1, _))
        case (_, Right(_)) => IO.raiseError(notOneCell(atStr))
      }

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
    precedentsEvaluated(wb, sheet, formula, overrides, position).flatMap { (evalSheet, evaluator) =>
      IO.fromEither(
        SheetEvaluator
          .evaluateFormulaUsing(evalSheet, formula, evaluator, Some(wb), currentCell = position)
          .left
          .map(XLException(_))
      )
    }

  /**
   * `sheet` with the `--with` overrides applied and only the formula cells in the formula's
   * dependency closure evaluated, in dependency order, through one evaluator the caller then
   * evaluates the formula with (GH-695: `x#` reads the arrays the fold computed):
   *   - formula chains work (`=C1` where C1=B1+50, B1=A1*2);
   *   - overrides propagate through chains (TJC-698): A1=200 → B1=400 → C1=450;
   *   - only the O(k) closure is evaluated, not the O(n) sheet. GH-197: bounded extraction, so a
   *     full-column range does not enumerate 1M+ cells.
   *
   * With `at` (GH-715) the formula is the cell there: a closure that reads `at` back is the
   * circular reference the plain cell would be, raised as the recalc cycle is, never `at`'s old
   * content.
   */
  private def precedentsEvaluated(
    wb: Workbook,
    sheet: Sheet,
    formula: String,
    overrides: List[String],
    at: Option[ARef] = None
  ): IO[(Sheet, Evaluator)] =
    for
      tempSheet <- applyOverrides(sheet, overrides)
      expr <- IO.fromEither(FormulaParser.parse(formula).left.map(unparseable(_, formula)))
      targetDeps = DependencyGraph.extractDependenciesBounded(expr, tempSheet.usedRange)
      graph = DependencyGraph.fromSheet(tempSheet)
      _ <- at.flatMap(cycleThrough(tempSheet, graph, expr, targetDeps, _)) match
        case Some(cycle) => IO.raiseError(cyclic(EvalError.CircularRef(cycle), formula))
        case None => IO.unit
      allDeps = DependencyGraph.transitiveDependencies(graph, targetDeps)
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
      evaluator = SheetEvaluator.spillTrackingEvaluator(formula, wb)
      evalSheet <- evalOrder.foldLeft(IO.pure(tempSheet)) { (sheetIO, ref) =>
        sheetIO.flatMap { s =>
          IO.fromEither(
            SheetEvaluator
              .evaluateCellUsing(s, ref, evaluator, Some(wb))
              .map(value => SheetEvaluator.threadComputed(s, ref, value))
              .left
              .map(XLException(_))
          )
        }
      }
    yield (evalSheet, evaluator)

  /**
   * The cycle `expr` would close as the cell `at` on `sheet`, `at` first and last: `expr` reads
   * `at`, or a formula in its closure does. Breadth first, so the shortest. Whether a formula reads
   * `at` is asked of its references themselves (bounded to `at` alone): `at` may lie outside the
   * used range that bounds the graph's range edges, as a total placed below its column does.
   */
  private def cycleThrough(
    sheet: Sheet,
    graph: DependencyGraph,
    expr: TExpr[?],
    targetDeps: Set[ARef],
    at: ARef
  ): Option[List[ARef]] =
    val only = Some(CellRange(at, at))
    def reads(e: TExpr[?]): Boolean =
      DependencyGraph.extractDependenciesBounded(e, only).contains(at)
    def readsBack(ref: ARef): Boolean =
      sheet(ref).value match
        case CellValue.Formula(text, _, _) => FormulaParser.parse(text).exists(reads)
        case _ => false
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
              graph.dependencies.getOrElse(ref, Set.empty).foldLeft((acc, seen)) {
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
        sheet <- SheetResolver.requireSheet(wb, sheetOpt, "evala")
        precedents <- precedentsEvaluated(wb, sheet, formula, overrides)
        (evalSheet, evaluator) = precedents
        // Parse target ref or use a virtual cell far from data
        originRef = targetRefOpt
          .flatMap(ARef.parse(_).toOption)
          .getOrElse(ARef.from0(25, 999)) // Z1000
        result <- IO.fromEither(
          SheetEvaluator
            .evaluateArrayFormulaUsing(evalSheet, formula, originRef, evaluator, Some(wb))
            .left
            .map(XLException(_))
        )
        (updatedSheet, spillRange) = result
      yield (formula, updatedSheet, spillRange)

  // ==========================================================================
  // Private helpers
  // ==========================================================================

  private def applyOverrides(sheet: Sheet, overrides: List[String]): IO[Sheet] =
    overrides.foldLeft(IO.pure(sheet)) { (sheetIO, override_) =>
      sheetIO.flatMap { s =>
        override_.split("=", 2) match
          case Array(refStr, valueStr) if valueStr.trim.nonEmpty =>
            // INVALID_REFERENCE: the parser's own text, a range, or another sheet's cell
            IO.fromEither(Resolve.ref(refStr.trim).left.map(CliException(_))).flatMap {
              case (None, Resolve.Target.Cell(ref)) =>
                val value = ValueParser.parseValue(valueStr.trim)
                IO.pure(s.put(ref, value))
              case (Some(sheetName), Resolve.Target.Cell(ref)) =>
                if sheetName == sheet.name then
                  val value = ValueParser.parseValue(valueStr.trim)
                  IO.pure(s.put(ref, value))
                else
                  IO.raiseError(
                    invalidRef(
                      s"Cross-sheet override not supported: ${refStr.trim}. " +
                        s"Eval operates on ${sheet.name.value}, not ${sheetName.value}"
                    )
                  )
              case (_, Resolve.Target.Range(_)) =>
                IO.raiseError(
                  invalidRef(s"Override requires single cell, not range: ${refStr.trim}")
                )
            }
          case Array(refStr, _) =>
            IO.raiseError(
              usage(s"Empty value for override: ${refStr.trim}. Use ref=value (e.g., B5=1000)")
            )
          case _ =>
            IO.raiseError(
              usage(s"Invalid override format: $override_. Use ref=value (e.g., B5=1000)")
            )
      }
    }
