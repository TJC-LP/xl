package com.tjclp.xl.cli.commands

import cats.effect.IO
import cats.implicits.*
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.{CliError, CliException}
import com.tjclp.xl.cli.helpers.{Resolve, SheetResolver, ValueParser}
import com.tjclp.xl.cli.output.{Escape, Format, JsonRenderer}
import com.tjclp.xl.error.XLException
import com.tjclp.xl.formula.{DependencyGraph, FormulaParser, SheetEvaluator}
import com.tjclp.xl.formula.eval.EvalError
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
   */
  def eval(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    formulaStr: String,
    overrides: List[String]
  ): IO[String] =
    evalResult(wb, sheetOpt, formulaStr, overrides).map { (formula, result) =>
      Format.evalSuccess(formula, result, overrides)
    }

  /**
   * `eval` as data (`eval --json`): `{formula, result: {type, value, formatted}, overrides}` — the
   * normalized formula (leading `=`), the value in the `view --format json` cell shape, and the
   * `--with` overrides as given. JSON TEXT, not a ujson tree: the number lexeme is exactly what
   * [[JsonRenderer]] prints (`12345678901234567` stays so), for the runner to splice as
   * `Payload.Raw`.
   */
  def evalData(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    formulaStr: String,
    overrides: List[String]
  ): IO[String] =
    evalResult(wb, sheetOpt, formulaStr, overrides).map { (formula, result) =>
      s"{\"formula\": ${Escape.json(formula)}, " +
        s"\"result\": ${JsonRenderer.valueJson(result, NumFmt.General)}, " +
        s"\"overrides\": ${overridesJson(overrides)}}"
    }

  /** The `--with` list as a JSON array of strings. */
  private def overridesJson(overrides: List[String]): String =
    ujson.write(ujson.Arr.from(overrides.map(ujson.Str.apply)))

  /** The evaluation both renderings share: the normalized formula and its value. */
  private def evalResult(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    formulaStr: String,
    overrides: List[String]
  ): IO[(String, CellValue)] =
    val formula = if formulaStr.startsWith("=") then formulaStr else s"=$formulaStr"

    // Check if formula needs a sheet by parsing and checking for cell references
    // GH-197: Use containsCellReferences instead of extractDependencies to avoid
    // enumerating 1M+ cells for full-column ranges like A:A
    // GH-210: Distinguish between formulas that need an ambient sheet (unqualified refs)
    // and those where all refs are qualified (e.g., =SUM(Revenue!B2:B5)).
    // Qualified-only formulas still need a workbook but can use any sheet as context.
    val (needsSheet, hasAnyCellRefs) = FormulaParser.parse(formula) match
      case Right(expr) =>
        val unqualified =
          DependencyGraph.containsUnqualifiedCellReferences(expr) || overrides.nonEmpty
        val anyRefs = DependencyGraph.containsCellReferences(expr) || overrides.nonEmpty
        (unqualified, anyRefs)
      case Left(_) =>
        // Parse error - let SheetEvaluator handle it (it will give a better error message)
        (false, false)

    if needsSheet then
      // Formula references unqualified cells - require sheet selection
      // Check if workbook is empty (no file provided)
      if wb.sheets.isEmpty then
        IO.raiseError(
          usage(
            "Formula references cells but no file provided. Use --file to specify an Excel file."
          )
        )
      else
        for
          sheet <- SheetResolver.requireSheet(wb, sheetOpt, "eval")
          tempSheet <- applyOverrides(sheet, overrides)
          // Optimization: Only pre-evaluate formulas in the dependency closure of the target.
          // This ensures:
          // 1. Formula chains work correctly (e.g., =C1 where C1=B1+50, B1=A1*2)
          // 2. Overrides propagate through chains (TJC-698): A1=200 → B1=400 → C1=450
          // 3. Performance: Only evaluate O(k) formulas in closure, not O(n) total formulas

          // 1. Extract target formula's direct dependencies
          // GH-197: Use bounded extraction to avoid enumerating 1M+ cells for full-column ranges
          targetDeps <- IO.fromEither(
            FormulaParser
              .parse(formula)
              .map(expr => DependencyGraph.extractDependenciesBounded(expr, tempSheet.usedRange))
              .left
              .map(unparseable(_, formula))
          )

          // 2. Build dependency graph and compute transitive closure
          graph = DependencyGraph.fromSheet(tempSheet)
          allDeps = DependencyGraph.transitiveDependencies(graph, targetDeps)

          // 3. Filter to only formula cells in the closure
          formulaDeps = allDeps.filter(ref =>
            tempSheet(ref).value match
              case _: CellValue.Formula => true
              case _ => false
          )

          // 4. Get evaluation order (topological sort filtered to closure)
          evalOrder <- IO.fromEither(
            if formulaDeps.isEmpty then scala.util.Right(List.empty[ARef])
            else
              DependencyGraph
                .topologicalSort(graph)
                .map(_.filter(formulaDeps.contains))
                .left
                .map(cyclic(_, formula))
          )

          // 5. Evaluate only formulas in the closure
          evalSheet <- evalOrder.foldLeft(IO.pure(tempSheet)) { (sheetIO, ref) =>
            sheetIO.flatMap { s =>
              IO.fromEither(
                SheetEvaluator
                  .evaluateCell(s)(ref, workbook = Some(wb))
                  .map(value => s.put(ref, value))
                  .left
                  .map(XLException(_))
              )
            }
          }

          result <- IO.fromEither(
            SheetEvaluator
              .evaluateFormula(evalSheet)(formula, workbook = Some(wb))
              .left
              .map(XLException(_))
          )
        yield (formula, result)
    else if hasAnyCellRefs then
      // GH-210: All cell refs are qualified (e.g., =SUM(Revenue!B2:B5)) - need workbook but
      // any sheet works as the ambient context since the evaluator resolves cross-sheet refs by name
      if wb.sheets.isEmpty then
        IO.raiseError(
          usage(
            "Formula references cells but no file provided. Use --file to specify an Excel file."
          )
        )
      else
        // Safe: wb.sheets.isEmpty is checked above, so headOption is always Some
        val sheet = (sheetOpt.orElse(wb.sheets.headOption): @unchecked) match
          case Some(s) => s
        for
          tempSheet <- applyOverrides(sheet, overrides)

          targetDeps <- IO.fromEither(
            FormulaParser
              .parse(formula)
              .map(expr => DependencyGraph.extractDependenciesBounded(expr, tempSheet.usedRange))
              .left
              .map(unparseable(_, formula))
          )

          graph = DependencyGraph.fromSheet(tempSheet)
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

          evalSheet <- evalOrder.foldLeft(IO.pure(tempSheet)) { (sheetIO, ref) =>
            sheetIO.flatMap { s =>
              IO.fromEither(
                SheetEvaluator
                  .evaluateCell(s)(ref, workbook = Some(wb))
                  .map(value => s.put(ref, value))
                  .left
                  .map(XLException(_))
              )
            }
          }

          result <- IO.fromEither(
            SheetEvaluator
              .evaluateFormula(evalSheet)(formula, workbook = Some(wb))
              .left
              .map(XLException(_))
          )
        yield (formula, result)
    else
      // Constant formula - use empty sheet or provided sheet
      val sheet = sheetOpt.getOrElse(Sheet("_eval"))
      // For constant formulas, workbook context is not needed
      val wbOpt = if wb.sheets.nonEmpty then Some(wb) else None
      for result <- IO.fromEither(
          SheetEvaluator
            .evaluateFormula(sheet)(formula, workbook = wbOpt)
            .left
            .map(XLException(_))
        )
      yield (formula, result)

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
        tempSheet <- applyOverrides(sheet, overrides)
        // Pre-evaluate formulas in the dependency closure (same as eval)
        // GH-197: Use bounded extraction to avoid enumerating 1M+ cells for full-column ranges
        targetDeps <- IO.fromEither(
          FormulaParser
            .parse(formula)
            .map(expr => DependencyGraph.extractDependenciesBounded(expr, tempSheet.usedRange))
            .left
            .map(unparseable(_, formula))
        )
        graph = DependencyGraph.fromSheet(tempSheet)
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
        evalSheet <- evalOrder.foldLeft(IO.pure(tempSheet)) { (sheetIO, ref) =>
          sheetIO.flatMap { s =>
            IO.fromEither(
              SheetEvaluator
                .evaluateCell(s)(ref, workbook = Some(wb))
                .map(value => s.put(ref, value))
                .left
                .map(XLException(_))
            )
          }
        }
        // Parse target ref or use a virtual cell far from data
        originRef = targetRefOpt
          .flatMap(ARef.parse(_).toOption)
          .getOrElse(ARef.from0(25, 999)) // Z1000
        result <- IO.fromEither(
          SheetEvaluator
            .evaluateArrayFormula(evalSheet)(formula, originRef, workbook = Some(wb))
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
