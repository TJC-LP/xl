package com.tjclp.xl.formula.eval

import scala.annotation.tailrec

import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.sheets.{DataValidation, DvKind, Sheet}
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.{Cell, CellError, CellValue, FormulaKind}
import com.tjclp.xl.cf.{CfRule, Cfvo, ConditionalFormat}
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.formula.printer.{FormulaPrinter, FormulaShifter}

/**
 * GH-128 / GH-129: workbook-level structural editing (insert/delete rows & columns) WITH formula
 * rewriting.
 *
 * The pure cell/merge/property shift lives in `xl-core` (`Sheet.insertRows`, ...). This layer adds
 * what xl-core cannot do (it has no formula parser): after shifting cells on the edited sheet, it
 * rewrites the formula strings of EVERY sheet so references track the edit. A formula that
 * references a fully-deleted cell/range becomes `#REF!` (`CellValue.Error(Ref)`); partially
 * overlapped ranges shrink. Cross-sheet references to the edited sheet are rewritten too.
 *
 * Determinism: the edit is a pure `Workbook => Workbook`; re-running on identical input yields
 * identical output.
 */
object StructuralEditor:

  /** Insert `count` rows at 0-based row index `at` on `sheet`. */
  def insertRows(wb: Workbook, sheet: SheetName, at: Int, count: Int): Workbook =
    edit(wb, sheet, isRow = true, at = at, delta = count)

  /** Delete `count` rows starting at 0-based row index `at` on `sheet`. */
  def deleteRows(wb: Workbook, sheet: SheetName, at: Int, count: Int): Workbook =
    edit(wb, sheet, isRow = true, at = at, delta = -count)

  /** Insert `count` columns at 0-based column index `at` on `sheet`. */
  def insertColumns(wb: Workbook, sheet: SheetName, at: Int, count: Int): Workbook =
    edit(wb, sheet, isRow = false, at = at, delta = count)

  /** Delete `count` columns starting at 0-based column index `at` on `sheet`. */
  def deleteColumns(wb: Workbook, sheet: SheetName, at: Int, count: Int): Workbook =
    edit(wb, sheet, isRow = false, at = at, delta = -count)

  private def edit(
    wb: Workbook,
    target: SheetName,
    isRow: Boolean,
    at: Int,
    delta: Int
  ): Workbook =
    val editedName = target.value
    // GH-455 follow-up: the non-participation fast path below keeps formula TEXT byte-identical,
    // but a cached VALUE can be stale even when the text never names the edited sheet —
    // Sheet2!Y ==X*2 over Sheet2!X =='Sheet1'!A1 reads Sheet1 two hops away. Invalidate the cache
    // of every formula transitively dependent (cross-sheet, through any chain) on any cell of the
    // edited sheet. One PRE-edit graph build per edit, forced lazily and only when a
    // non-participating formula actually rides the fast path (the CLI path recalculates after).
    lazy val staleCaches: Set[QualifiedRef] =
      val (_, dependents) = DependencyGraph.fromWorkbookBounded(wb)
      val seeds = dependents.keySet.filter(_.sheet == target)
      // A dynamic reference can reach the edited sheet without contributing a static graph edge.
      // Conservatively invalidate every dynamic cell and its static dependent closure: the edit
      // may have changed what its unchanged reference text resolves to.
      val dynamic = wb.sheets.iterator.flatMap { s =>
        DependencyGraph.dynamicCells(s).iterator.map(r => QualifiedRef(s.name, r))
      }.toSet
      dynamic ++ DependencyGraph.qualifiedTransitiveDependents(dependents, seeds ++ dynamic)
    val updatedSheets = wb.sheets.map { s =>
      // 1. Pure cell/merge/property shift — only on the edited sheet. Its own typed charts
      //    (anchors + same-sheet data refs) are handled INSIDE the shift (GH-222).
      val shifted =
        if s.name == target then
          (isRow, delta >= 0) match
            case (true, true) => s.insertRows(at, delta)
            case (true, false) => s.deleteRows(at, -delta)
            case (false, true) => s.insertColumns(at, delta)
            case (false, false) => s.deleteColumns(at, -delta)
        else
          // GH-222: typed charts on OTHER sheets track the edited sheet's geometry (no
          // double-shift: the edited sheet's charts were already shifted above).
          s.shiftChartRefs(editedName, isRow, at, delta)
      // 2. Rewrite formula references on every sheet (local refs only on the edited sheet;
      //    sheet-qualified refs to the edited sheet on all sheets).
      rewriteFormulas(
        shifted,
        shiftLocal = s.name == target,
        editedName,
        isRow,
        at,
        delta,
        stale = r => staleCaches.contains(QualifiedRef(s.name, r))
      )
    }
    // A structural edit can touch any sheet's formulas (cross-sheet refs), so mark every sheet
    // modified — otherwise the writer's clean/verbatim fast-path would copy stale bytes from the
    // source file and silently drop the edit.
    val updatedContext =
      wb.sourceContext.map(ctx => wb.sheets.indices.foldLeft(ctx)((c, i) => c.markSheetModified(i)))
    wb.copy(sheets = updatedSheets, sourceContext = updatedContext)

  private def rewriteFormulas(
    sheet: Sheet,
    shiftLocal: Boolean,
    editedSheet: String,
    isRow: Boolean,
    at: Int,
    delta: Int,
    stale: ARef => Boolean
  ): Sheet =
    // GH-455 follow-up: a non-participating formula keeps its TEXT byte-identical, but a cached
    // value that transitively reads the edited sheet is stale and must drop (the writer would
    // otherwise emit it into <v>). Independent formulas keep text AND cache untouched.
    def keepNonParticipant(ref: ARef, cell: Cell, f: CellValue.Formula): Cell =
      if f.cachedValue.isDefined && stale(ref) then cell.copy(value = f.copy(cachedValue = None))
      else cell
    val updatedCells = sheet.cells.map { case (ref, cell) =>
      cell.value match
        // GH-430: a data-table record's payload (ref/r1/r2) is LOCAL sheet geometry — it moves
        // only when this sheet is the edited one; its TABLE(...) text is derived, never parsed.
        case CellValue.Formula(_, cachedOpt, dt: FormulaKind.DataTable) =>
          if !shiftLocal then (ref, cell)
          else
            val newRange = Option
              .when(!editIntersects(dt.ref, isRow, at, delta))(())
              .flatMap(_ => shiftedRange(dt.ref, isRow, at, delta))
            val newR1 = dt.r1.map(r => shiftedRef(r, isRow, at, delta))
            val newR2 = dt.r2.map(r => shiftedRef(r, isRow, at, delta))
            newRange match
              case Some(range) =>
                // GH-435: an edit that removes an input cell is Excel's own del1/del2 case — the
                // record survives with that input omitted and its flag set, interior caches
                // intact. Flags are sticky: a record that arrived del-flagged stays so.
                val newKind = dt.copy(
                  ref = range,
                  r1 = newR1.flatten,
                  r2 = newR2.flatten,
                  del1 = dt.del1 || newR1.exists(_.isEmpty),
                  del2 = dt.del2 || newR2.exists(_.isEmpty)
                )
                val newValue =
                  CellValue.Formula(FormulaKind.displayExpression(newKind), cachedOpt, newKind)
                (ref, cell.copy(value = newValue))
              case None =>
                // The band tears the table interior itself. Excel refuses to "change part of a
                // data table"; the total-function analog degrades the cell to its cached constant
                // (visible, deterministic).
                (ref, cell.copy(value = cachedOpt.getOrElse(CellValue.Empty)))
        case f @ CellValue.Formula(formulaStr, _, kind) =>
          // GH-455: a formula on a NON-edited sheet participates only through sheet-qualified
          // references to the edited sheet. Everything else rides byte-identical in TEXT
          // (reprinting canonicalizes text it has no business touching), and keeps its cached
          // value too unless it transitively depends on the edited sheet (`stale`). The text
          // gate skips the parse entirely (any real reference spells the sheet name in the
          // formula text); the AST gate settles the false positives exactly.
          if !shiftLocal && !mentionsSheet(formulaStr, editedSheet) then
            (ref, keepNonParticipant(ref, cell, f))
          else
            FormulaParser.parse(formulaStr) match
              case Right(expr)
                  if !shiftLocal && !FormulaShifter.referencesSheet(expr, editedSheet) =>
                (ref, keepNonParticipant(ref, cell, f))
              case Right(expr) =>
                FormulaShifter.shiftStructural(
                  expr,
                  shiftLocal,
                  editedSheet,
                  isRow,
                  at,
                  delta
                ) match
                  case Some(shiftedExpr) =>
                    // GH-427: the model's canonical formula form is equals-free (the reader strips
                    // the '='; the writer serializes the string VERBATIM into <f>, where a leading
                    // '=' is a spec deviation openpyxl reads back as '==...').
                    //
                    // A structural edit invalidates formula caches even when every parsed reference
                    // survives: a shortened range can produce a different aggregate, and moving a
                    // cell changes position-sensitive formulas such as ROW(). Leave the expression
                    // evaluatable and uncached so the next recalculation cannot expose stale data.
                    val newStr = FormulaPrinter.print(shiftedExpr, includeEquals = false)
                    val newKind = shiftedArrayKind(kind, shiftLocal, isRow, at, delta)
                    (ref, cell.copy(value = CellValue.Formula(newStr, None, newKind)))
                  case None =>
                    (ref, cell.copy(value = CellValue.Error(CellError.Ref)))
              // Preserve unparseable text rather than guessing at a rewrite, but invalidate its
              // cache too: the edit may have moved the cell or changed a dynamically-read value.
              case Left(_) => (ref, cell.copy(value = f.copy(cachedValue = None)))
        case _ => (ref, cell)
    }
    sheet.copy(
      cells = updatedCells,
      conditionalFormats =
        rewriteCfFormulas(sheet.conditionalFormats, shiftLocal, editedSheet, isRow, at, delta),
      dataValidations =
        rewriteDvFormulas(sheet.dataValidations, shiftLocal, editedSheet, isRow, at, delta)
    )

  /**
   * GH-455: conservative text pre-filter for the non-edited-sheet fast path — false means the
   * formula text cannot contain a reference to `editedSheet`, so parsing is pointless. Quoted sheet
   * names double their apostrophes in formula text (`'It''s'!A1`), so a doubled apostrophe in the
   * formula matches a single apostrophe in the name — collapsed inline during the case-insensitive
   * scan (matching the shifter's `equalsIgnoreCase` convention) instead of allocating an undoubled
   * copy per formula per sheet per edit. False positives (the name inside a string literal, say)
   * are settled exactly by the AST gate.
   */
  private def mentionsSheet(formula: String, editedSheet: String): Boolean =
    val n = editedSheet.length
    val len = formula.length
    // String.regionMatches(ignoreCase = true) equality: exact, or upper, or lower match.
    def charsMatch(a: Char, b: Char): Boolean =
      a == b || Character.toUpperCase(a) == Character.toUpperCase(b) ||
        Character.toLowerCase(a) == Character.toLowerCase(b)
    @tailrec
    def matchesAt(fi: Int, pi: Int): Boolean =
      if pi >= n then true
      else if fi >= len then false
      else
        val fc = formula.charAt(fi)
        if !charsMatch(fc, editedSheet.charAt(pi)) then false
        else
          // A matched apostrophe consumes its doubled partner ('' in text is ' in the name).
          val step = if fc == '\'' && fi + 1 < len && formula.charAt(fi + 1) == '\'' then 2 else 1
          matchesAt(fi + step, pi + 1)
    @tailrec
    def scan(i: Int): Boolean =
      if i >= len then false
      else if matchesAt(i, 0) then true
      else scan(i + 1)
    n == 0 || scan(0)

  // ========== GH-430: record-payload geometry (FormulaKind ref/r1/r2 shifting) ==========

  /** Shift a 0-based index for an axis edit at `at` by `delta`; None = deleted/out of bounds. */
  private def shiftedIndex0(index0: Int, at: Int, delta: Int, max: Int): Option[Int] =
    val shifted =
      if delta >= 0 then Some(if index0 >= at then index0 + delta else index0)
      else
        val cut = at - delta // first index past the deleted band
        if index0 >= cut then Some(index0 + delta)
        else if index0 >= at then None // inside the deleted band
        else Some(index0)
    shifted.filter(i => i >= 0 && i <= max)

  /** Shift a single cell ref along the edited axis; None = the cell was deleted/pushed out. */
  private def shiftedRef(ref: ARef, isRow: Boolean, at: Int, delta: Int): Option[ARef] =
    if isRow then
      shiftedIndex0(ref.row.index0, at, delta, Row.MaxIndex0)
        .map(r => ARef.from0(ref.col.index0, r))
    else
      shiftedIndex0(ref.col.index0, at, delta, Column.MaxIndex0)
        .map(c => ARef.from0(c, ref.row.index0))

  private def shiftedRange(
    range: CellRange,
    isRow: Boolean,
    at: Int,
    delta: Int
  ): Option[CellRange] =
    for
      start <- shiftedRef(range.start, isRow, at, delta)
      end <- shiftedRef(range.end, isRow, at, delta)
    yield CellRange(start, end)

  /**
   * True when the edit band TEARS the range on the edited axis: a delete band overlapping any of
   * it, or an insert strictly inside it (insert at the range start moves the whole range intact).
   */
  private def editIntersects(range: CellRange, isRow: Boolean, at: Int, delta: Int): Boolean =
    val (s, e) =
      if isRow then (range.start.row.index0, range.end.row.index0)
      else (range.start.col.index0, range.end.col.index0)
    if delta < 0 then at <= e && (at - delta - 1) >= s
    else at > s && at <= e

  /**
   * GH-430: shift an ArrayFormula record's anchor range with the edit; a band tearing the range
   * degrades the kind to Normal (the shifted TEXT is kept — Excel's visible-degradation analog).
   * The degraded cell keeps `ca` (GH-435: volatile marking is legal on a plain formula) and drops
   * the array-only `aca`. DataTable kinds never reach here (dedicated branch above); Normal is
   * identity.
   */
  private def shiftedArrayKind(
    kind: FormulaKind,
    shiftLocal: Boolean,
    isRow: Boolean,
    at: Int,
    delta: Int
  ): FormulaKind =
    kind match
      case arr: FormulaKind.ArrayFormula if shiftLocal =>
        val degraded = FormulaKind.Normal(ca = arr.ca)
        if editIntersects(arr.ref, isRow, at, delta) then degraded
        else
          shiftedRange(arr.ref, isRow, at, delta)
            .map(r => arr.copy(ref = r))
            .getOrElse(degraded)
      case other => other

  /**
   * Rewrite bare formula TEXT (no leading '=') through the structural shift: unparseable text —
   * including inline list literals like `"yes,no"`, which parse as string literals and print back
   * verbatim — rides unchanged; a fully-deleted reference degrades the text to `"#REF!"` (the
   * Excel-observable surface). Shared by the CF and DV formula rewrites.
   */
  private def shiftFormulaText(
    formula: String,
    shiftLocal: Boolean,
    editedSheet: String,
    isRow: Boolean,
    at: Int,
    delta: Int
  ): String =
    // GH-455: same non-participation gates as the cell-formula path — a CF/DV formula on a
    // non-edited sheet that cannot reference the edited sheet rides byte-identical.
    if !shiftLocal && !mentionsSheet(formula, editedSheet) then formula
    else
      FormulaParser.parse(s"=$formula") match
        case Right(expr) if !shiftLocal && !FormulaShifter.referencesSheet(expr, editedSheet) =>
          formula
        case Right(expr) =>
          FormulaShifter.shiftStructural(expr, shiftLocal, editedSheet, isRow, at, delta) match
            case Some(shifted) => FormulaPrinter.print(shifted, includeEquals = false)
            case None => "#REF!"
        case Left(_) => formula

  /**
   * GH-136: rewrite TYPED conditional-format formula text (CellIs.formula1/formula2,
   * Expression.formula, Cfvo.Formula inside ColorScale points and DataBar bounds) through the same
   * shift as cell formulas. A fully-deleted reference degrades the formula TEXT to "#REF!" with the
   * rule kept — the Excel-observable surface, consistent with the whole-cell `CellValue.Error(Ref)`
   * behavior above. Unparseable text rides verbatim (same precedent). Text-family rules store no
   * formula (derived at emission) and Preserved payloads are never touched; their typed envelopes
   * were already shifted by the pure core shift.
   */
  private def rewriteCfFormulas(
    cfs: Vector[ConditionalFormat],
    shiftLocal: Boolean,
    editedSheet: String,
    isRow: Boolean,
    at: Int,
    delta: Int
  ): Vector[ConditionalFormat] =
    def shiftText(formula: String): String =
      shiftFormulaText(formula, shiftLocal, editedSheet, isRow, at, delta)
    def shiftCfvo(cfvo: Cfvo): Cfvo = cfvo match
      case Cfvo.Formula(f) => Cfvo.Formula(shiftText(f))
      case other => other
    cfs.map {
      case ConditionalFormat.Rules(ranges, rules, pivot) =>
        val shifted = rules.map {
          case r: CfRule.CellIs =>
            r.copy(formula1 = shiftText(r.formula1), formula2 = r.formula2.map(shiftText))
          case r: CfRule.Expression => r.copy(formula = shiftText(r.formula))
          case r: CfRule.ColorScale =>
            r.copy(
              min = r.min.copy(cfvo = shiftCfvo(r.min.cfvo)),
              mid = r.mid.map(p => p.copy(cfvo = shiftCfvo(p.cfvo))),
              max = r.max.copy(cfvo = shiftCfvo(r.max.cfvo))
            )
          case r: CfRule.DataBar => r.copy(min = shiftCfvo(r.min), max = shiftCfvo(r.max))
          case other @ (_: CfRule.Top10 | _: CfRule.Text | _: CfRule.Preserved) => other
        }
        ConditionalFormat.Rules(ranges, shifted, pivot)
      case preserved: ConditionalFormat.Preserved => preserved
    }

  /**
   * GH-429: rewrite TYPED data-validation formula text through the same shift as cell and cf
   * formulas — without this, absolute list sources (`$Z$1:$Z$3`, `Lists!$A$1:$A$5`) detach from
   * their data on every structural edit. Inline literals (`"yes,no"`) parse as string literals and
   * ride verbatim; a fully-deleted source degrades to `"#REF!"` text; Preserved payload formulas
   * are never rewritten (their sqref envelope was already shifted by the pure core shift).
   */
  private def rewriteDvFormulas(
    dvs: Vector[DataValidation],
    shiftLocal: Boolean,
    editedSheet: String,
    isRow: Boolean,
    at: Int,
    delta: Int
  ): Vector[DataValidation] =
    def shiftText(formula: String): String =
      shiftFormulaText(formula, shiftLocal, editedSheet, isRow, at, delta)
    dvs.map {
      case rules: DataValidation.Rules =>
        val newKind = rules.kind match
          case DvKind.List(f) => DvKind.List(shiftText(f))
          case DvKind.Custom(f) => DvKind.Custom(shiftText(f))
          case DvKind.AnyValue => DvKind.AnyValue
          case DvKind.Bounded(t, op, f1, f2) =>
            DvKind.Bounded(t, op, shiftText(f1), f2.map(shiftText))
        rules.copy(kind = newKind)
      case preserved: DataValidation.Preserved => preserved
    }

  /** Ergonomic workbook extensions. */
  extension (wb: Workbook)
    def insertRowsShifted(sheet: SheetName, at: Int, count: Int): Workbook =
      StructuralEditor.insertRows(wb, sheet, at, count)
    def deleteRowsShifted(sheet: SheetName, at: Int, count: Int): Workbook =
      StructuralEditor.deleteRows(wb, sheet, at, count)
    def insertColumnsShifted(sheet: SheetName, at: Int, count: Int): Workbook =
      StructuralEditor.insertColumns(wb, sheet, at, count)
    def deleteColumnsShifted(sheet: SheetName, at: Int, count: Int): Workbook =
      StructuralEditor.deleteColumns(wb, sheet, at, count)
