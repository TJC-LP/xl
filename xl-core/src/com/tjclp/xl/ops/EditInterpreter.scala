package com.tjclp.xl.ops

import scala.annotation.tailrec

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.{Sheet, SheetEdits}
import com.tjclp.xl.sheets.styleSyntax.{getCellStyle, withCellStyle}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.workbooks.{DefinedName, Workbook}

/**
 * The one interpreter of [[Edit]] (ADR-017 §2.12): `validate` is the edit-local check that needs no
 * workbook, `applyAll` the fail-fast, all-or-nothing left fold that gives every surface — batch
 * JSON, the CLI verbs, `wb.edit(...)` — the same meaning, `plan` the same fold keeping only its row
 * list (a semantic dry-run: it applies and discards), `lower` the bridge to the law-tested `Patch`
 * kernel. Reached through the `Edit` companion.
 *
 * Invariants kept here: every sheet goes back through `Workbook.put` (invariant 1, so
 * `modifiedSheets` is right); THE sheet rule resolves each target (a qualifier, then the scope's
 * default, then the only sheet of a single-sheet book, else `SheetRequired`); a hint on a value is
 * `Inferred | Explicit` and only `Explicit` overrides a non-General numFmt (invariant 4); a failure
 * at position `i` is `EditFailed(i, name, cause)` and the input workbook is never partially
 * written.
 */
private[xl] object EditInterpreter:

  /** Excel's ceiling on a column width, in character units. */
  private val MaxColumnWidth: Double = 255.0

  /** Excel's ceiling on a row height, in points. */
  private val MaxRowHeight: Double = 409.0

  private def refuse(message: String): XLResult[Unit] = Left(XLError.Other(message))

  // ===== validate =====

  /** The checks that need no workbook: counts, ranges, guards the model enforces with `require`. */
  def validate(edit: Edit)(using fs: FormulaSupport): XLResult[Unit] =
    val name = EditSchema.nameOf(edit)
    edit match
      case Edit.Put(_, _, _) => Right(())
      case Edit.PutValues(at, values, _) => countMatch(name, at, values.size)
      case Edit.PutFormula(_, formula, _) => fs.validate(formula)
      case Edit.PutFormulas(at, formulas, _) =>
        countMatch(name, at, formulas.size).flatMap(_ => each(formulas)(fs.validate))
      case Edit.DragFormula(_, formula, _, _) => fs.validate(formula)
      case Edit.Fill(source, target, direction) =>
        SheetEdits.validateFill(source.range, target, direction)
      case Edit.Copy(source, target, _) => copyTarget(source, target).map(_ => ())
      case Edit.Sort(at, keys, _) => SheetEdits.validateSortKeys(at.range, keys)
      case Edit.Clear(_, what) =>
        if what.isEmpty then refuse(s"$name: nothing to clear (contents, styles or comments)")
        else Right(())
      case Edit.Style(_, overlay, mode) =>
        if overlay.isEmpty && mode == StyleMode.Merge then
          refuse(s"$name: at least one property is required (an empty merge changes nothing)")
        else overlay.validate
      case Edit.Merge(_) | Edit.Unmerge(_) => Right(())
      case Edit.ColWidth(_, _, width) =>
        if width >= 0 && width <= MaxColumnWidth then Right(())
        else refuse(s"$name: width must be 0-${MaxColumnWidth.toInt} character units, got $width")
      case Edit.RowHeight(_, _, height) =>
        if height >= 0 && height <= MaxRowHeight then Right(())
        else refuse(s"$name: height must be 0-${MaxRowHeight.toInt} points, got $height")
      case Edit.HideCols(_, _) | Edit.ShowCols(_, _) | Edit.HideRows(_, _) | Edit.ShowRows(_, _) =>
        Right(())
      case Edit.GroupRows(_, _, level, _) => SheetEdits.validateLevel(level)
      case Edit.GroupCols(_, _, level, _) => SheetEdits.validateLevel(level)
      case Edit.UngroupRows(_, _) | Edit.UngroupCols(_, _) | Edit.AutoFit(_, _) => Right(())
      case Edit.SetComment(_, _) | Edit.RemoveComment(_) => Right(())
      case Edit.Hyperlink(_, target) =>
        if target.exists(_.trim.isEmpty) then
          refuse(s"$name: target cannot be empty (omit it to clear the link)")
        else Right(())
      case Edit.AddConditionalFormat(_, ranges, rules) =>
        if ranges.isEmpty || rules.isEmpty then
          refuse(s"$name: at least one range and one rule are required")
        else Right(())
      case Edit.AddChart(_, _, _) | Edit.AddImage(_, _, _) => Right(())
      case Edit.Freeze(_) | Edit.Unfreeze(_) => Right(())
      case Edit.SetSheetView(_, _, zoom, _) => SheetEdits.validateZoom(zoom)
      case Edit.SetTabColor(_, _) | Edit.SetAutoFilter(_, _) => Right(())
      case Edit.SetPageSetup(_, orientation, scale, fitToWidth, fitToHeight, _) =>
        SheetEdits.validatePageSetup(orientation, scale, fitToWidth, fitToHeight)
      case Edit.SetHeaderFooter(_, _, _, _, _, _, _, _, _) => Right(())
      case Edit.InsertRows(_, _, count) => positiveCount(name, count)
      case Edit.DeleteRows(_, _, count) => positiveCount(name, count)
      case Edit.InsertCols(_, _, count) => positiveCount(name, count)
      case Edit.DeleteCols(_, _, count) => positiveCount(name, count)
      case Edit.AddSheet(_, after, before) =>
        if after.isDefined && before.isDefined then
          refuse(s"$name: after and before are mutually exclusive")
        else Right(())
      case Edit.RemoveSheet(_) | Edit.RenameSheet(_, _) | Edit.CopySheet(_, _) => Right(())
      case Edit.MoveSheet(_, toIndex, after, before) =>
        if Vector(toIndex, after, before).count(_.isDefined) == 1 then Right(())
        else refuse(s"$name: exactly one of to, after or before is required")
      case Edit.HideSheet(_, _) | Edit.ShowSheet(_) => Right(())
      case Edit.DefineName(defined, refersTo, _) =>
        if defined.trim.isEmpty then refuse(s"$name: name cannot be empty")
        else if refersTo.trim.isEmpty then refuse(s"$name: refersTo cannot be empty")
        else Right(())
      case Edit.RemoveName(defined, _) =>
        if defined.trim.isEmpty then refuse(s"$name: name cannot be empty") else Right(())

  private def countMatch(name: String, at: Area, actual: Int): XLResult[Unit] =
    val expected = at.range.cellCount
    if expected == actual.toLong then Right(())
    else Left(XLError.ValueCountMismatch(expected, actual, s"$name ${at.range.toA1}"))

  private def positiveCount(name: String, count: Int): XLResult[Unit] =
    if count >= 1 then Right(()) else refuse(s"$name: count must be at least 1, got $count")

  private def each[A](as: Vector[A])(f: A => XLResult[Unit]): XLResult[Unit] =
    as.foldLeft[XLResult[Unit]](Right(()))((acc, a) => acc.flatMap(_ => f(a)))

  /** `foldLeft` that stops at the first `Left`. */
  private def foldE[A, B](as: Iterator[A])(zero: B)(f: (B, A) => XLResult[B]): XLResult[B] =
    as.foldLeft[XLResult[B]](Right(zero)) {
      case (Right(acc), a) => f(acc, a)
      case (left, _) => left
    }

  /** The range a copy lands on: `target` expanded to the source's shape, refused past the grid. */
  private def copyTarget(source: Area, target: Loc): XLResult[CellRange] =
    target.ref
      .tryShift(source.range.width - 1, source.range.height - 1)
      .map(end => CellRange(target.ref, end))
      .toRight(
        XLError.OutOfBounds(
          target.toA1,
          s"copying ${source.range.toA1} there would run past the sheet's last cell " +
            ARef.from0(Column.MaxIndex0, Row.MaxIndex0).toA1
        )
      )

  // ===== applyAll / plan =====

  def applyAll(wb: Workbook, edits: Vector[Edit], scope: Scope)(using
    FormulaSupport
  ): XLResult[Applied] =
    @tailrec
    def loop(
      current: Workbook,
      currentScope: Scope,
      remaining: Vector[Edit],
      index: Int,
      rows: Vector[Planned]
    ): XLResult[Applied] =
      remaining match
        case edit +: rest =>
          applyOne(current, edit, currentScope, index) match
            case Right((next, row)) =>
              loop(next, currentScope.after(edit), rest, index + 1, rows :+ row)
            case Left(cause) => Left(XLError.EditFailed(index, EditSchema.nameOf(edit), cause))
        case _ => Right(Applied(current, rows, currentScope))
    loop(wb, scope, edits, 1, Vector.empty)

  def plan(wb: Workbook, edits: Vector[Edit], scope: Scope)(using
    FormulaSupport
  ): XLResult[Vector[Planned]] =
    applyAll(wb, edits, scope).map(_.planned)

  // ===== THE sheet rule =====

  private def findSheet(wb: Workbook, name: SheetName): XLResult[Sheet] =
    wb.sheets.find(_.name == name).toRight(XLError.SheetNotFound(name.value))

  private def indexOf(wb: Workbook, name: SheetName): XLResult[Int] =
    wb.sheets.indexWhere(_.name == name) match
      case -1 => Left(XLError.SheetNotFound(name.value))
      case i => Right(i)

  /** A qualifier wins, then the scope's default, then the only sheet, else `SheetRequired`. */
  private def resolve(
    wb: Workbook,
    qualifier: Option[SheetName],
    scope: Scope,
    op: String
  ): XLResult[Sheet] =
    qualifier.orElse(scope.defaultSheet) match
      case Some(name) => findSheet(wb, name)
      case None =>
        wb.sheets match
          case Vector(only) => Right(only)
          case _ => Left(XLError.SheetRequired(op, wb.sheets.map(_.name.value)))

  // ===== one edit =====

  private def applyOne(wb: Workbook, edit: Edit, scope: Scope, index: Int)(using
    fs: FormulaSupport
  ): XLResult[(Workbook, Planned)] =
    val name = EditSchema.nameOf(edit)

    /** A sheet-local edit: resolve, apply `f`, write back through `put`. */
    def onSheet(qualifier: Option[SheetName], touched: Vector[CellRange])(
      f: Sheet => XLResult[Sheet]
    ): XLResult[(Workbook, Planned)] =
      resolve(wb, qualifier, scope, name).flatMap { sheet =>
        f(sheet).map(updated => (wb.put(updated), Planned(index, edit, Some(sheet.name), touched)))
      }

    def pure(qualifier: Option[SheetName], touched: Vector[CellRange])(
      f: Sheet => Sheet
    ): XLResult[(Workbook, Planned)] =
      onSheet(qualifier, touched)(s => Right(f(s)))

    /** A structural edit: the whole workbook is rewritten through the formula support. */
    def structural(qualifier: Option[SheetName])(
      f: SheetName => XLResult[Workbook]
    ): XLResult[(Workbook, Planned)] =
      resolve(wb, qualifier, scope, name).flatMap { sheet =>
        f(sheet.name).map(next => (next, Planned(index, edit, Some(sheet.name), Vector.empty)))
      }

    def workbookLevel(result: XLResult[Workbook]): XLResult[(Workbook, Planned)] =
      result.map(next => (next, Planned(index, edit, None, Vector.empty)))

    validate(edit).flatMap { _ =>
      edit match
        // ----- cell content -----
        case Edit.Put(at, value, format) =>
          pure(at.sheet, Vector(CellRange(at.ref, at.ref)))(putHinted(_, at.ref, value, format))

        case Edit.PutValues(at, values, format) =>
          pure(at.sheet, Vector(at.range)) { sheet =>
            at.range.cellsRowMajor.zip(values.iterator).foldLeft(sheet) { case (acc, (ref, v)) =>
              putHinted(acc, ref, v, format)
            }
          }

        case Edit.PutFormula(at, formula, format) =>
          pure(at.sheet, Vector(CellRange(at.ref, at.ref)))(
            putHinted(_, at.ref, CellValue.Formula(formulaText(formula), None), format)
          )

        case Edit.PutFormulas(at, formulas, format) =>
          pure(at.sheet, Vector(at.range)) { sheet =>
            at.range.cellsRowMajor.zip(formulas.iterator).foldLeft(sheet) { case (acc, (ref, f)) =>
              putHinted(acc, ref, CellValue.Formula(formulaText(f), None), format)
            }
          }

        case Edit.DragFormula(at, formula, anchor, format) =>
          onSheet(at.sheet, Vector(at.range)) { sheet =>
            val text = formula.trim
            foldE(at.range.cellsRowMajor)(sheet) { (acc, ref) =>
              val colDelta = ref.col.index0 - anchor.col.index0
              val rowDelta = ref.row.index0 - anchor.row.index0
              val shifted =
                if colDelta == 0 && rowDelta == 0 then Right(text)
                else fs.shift(text, colDelta, rowDelta)
              shifted.map(t => putHinted(acc, ref, CellValue.Formula(formulaText(t), None), format))
            }
          }

        case Edit.Fill(source, target, direction) =>
          onSheet(source.sheet, Vector(target))(_.fill(source.range, target, direction))

        case Edit.Copy(source, target, valuesOnly) =>
          for
            targetRange <- copyTarget(source, target)
            sourceSheet <- resolve(wb, source.sheet, scope, name)
            targetSheet <- resolve(wb, target.sheet, scope, name)
            updated <-
              if sourceSheet.name == targetSheet.name then
                sourceSheet.copyRange(source.range, targetRange, valuesOnly)
              else targetSheet.copyRangeFrom(sourceSheet, source.range, targetRange, valuesOnly)
          yield (wb.put(updated), Planned(index, edit, Some(targetSheet.name), Vector(targetRange)))

        case Edit.Sort(at, keys, hasHeader) =>
          onSheet(at.sheet, Vector(at.range))(_.sort(at.range, keys, hasHeader))

        case Edit.Clear(at, what) =>
          val touched = if what.contents then Vector(at.range) else Vector.empty
          pure(at.sheet, touched)(_.clearRange(at.range, what))

        // ----- style & layout -----
        case Edit.Style(at, overlay, mode) =>
          pure(at.sheet, Vector.empty)(applyStyle(_, at.range, overlay, mode))

        case Edit.Merge(at) => pure(at.sheet, Vector.empty)(_.merge(at.range))
        case Edit.Unmerge(at) => pure(at.sheet, Vector.empty)(_.unmerge(at.range))

        case Edit.ColWidth(sheet, cols, width) =>
          pure(sheet, Vector.empty)(columns(_, cols)(_.copy(width = Some(width))))
        case Edit.RowHeight(sheet, rows, height) =>
          pure(sheet, Vector.empty)(rowsOf(_, rows)(_.copy(height = Some(height))))
        case Edit.HideCols(sheet, cols) =>
          pure(sheet, Vector.empty)(columns(_, cols)(_.copy(hidden = true)))
        case Edit.ShowCols(sheet, cols) =>
          pure(sheet, Vector.empty)(columns(_, cols)(_.copy(hidden = false)))
        case Edit.HideRows(sheet, rows) =>
          pure(sheet, Vector.empty)(rowsOf(_, rows)(_.copy(hidden = true)))
        case Edit.ShowRows(sheet, rows) =>
          pure(sheet, Vector.empty)(rowsOf(_, rows)(_.copy(hidden = false)))

        case Edit.GroupRows(sheet, rows, level, collapsed) =>
          onSheet(sheet, Vector.empty)(_.groupRows(rows, level, collapsed))
        case Edit.GroupCols(sheet, cols, level, collapsed) =>
          onSheet(sheet, Vector.empty)(_.groupCols(cols, level, collapsed))
        case Edit.UngroupRows(sheet, rows) => pure(sheet, Vector.empty)(_.ungroupRows(rows))
        case Edit.UngroupCols(sheet, cols) => pure(sheet, Vector.empty)(_.ungroupCols(cols))

        case Edit.AutoFit(sheet, cols) =>
          pure(sheet, Vector.empty)(s => cols.fold(s.autoFitAll)(span => s.autoFit(span.columns)))

        // ----- annotations & objects -----
        case Edit.SetComment(at, comment) =>
          pure(at.sheet, Vector.empty)(_.comment(at.ref, comment))
        case Edit.RemoveComment(at) => pure(at.sheet, Vector.empty)(_.removeComment(at.ref))
        case Edit.Hyperlink(at, target) =>
          pure(at.sheet, Vector.empty) { s =>
            s.put(target.fold(s(at.ref).clearHyperlink)(s(at.ref).withHyperlink))
          }
        case Edit.AddConditionalFormat(sheet, ranges, rules) =>
          pure(sheet, Vector.empty)(_.conditionalFormat(ranges, rules))
        case Edit.AddChart(sheet, chart, anchor) =>
          pure(sheet, Vector.empty)(_.addChart(chart, anchor))
        case Edit.AddImage(sheet, image, anchor) =>
          pure(sheet, Vector.empty)(_.addImage(image, anchor))

        // ----- sheet view & print -----
        case Edit.Freeze(at) => pure(at.sheet, Vector.empty)(_.freezeAt(at.ref))
        case Edit.Unfreeze(sheet) => pure(sheet, Vector.empty)(_.unfreeze)
        case Edit.SetSheetView(sheet, gridlines, zoom, tabSelected) =>
          onSheet(sheet, Vector.empty)(_.mergeSheetView(gridlines, zoom, tabSelected))
        case Edit.SetTabColor(sheet, color) =>
          pure(sheet, Vector.empty)(s => color.fold(s.withoutTabColor)(s.withTabColor))
        case Edit.SetAutoFilter(sheet, range) =>
          pure(sheet, Vector.empty)(s => range.fold(s.removeAutoFilter)(s.withAutoFilter))
        case Edit.SetPageSetup(sheet, orientation, scale, fitToWidth, fitToHeight, fitToPage) =>
          onSheet(sheet, Vector.empty)(
            _.mergePageSetup(orientation, scale, fitToWidth, fitToHeight, fitToPage)
          )
        case Edit.SetHeaderFooter(
              sheet,
              oh,
              of,
              eh,
              ef,
              fh,
              ff,
              differentOddEven,
              differentFirst
            ) =>
          pure(sheet, Vector.empty)(
            _.mergeHeaderFooter(oh, of, eh, ef, fh, ff, differentOddEven, differentFirst)
          )

        // ----- structure -----
        case Edit.InsertRows(sheet, at, count) =>
          structural(sheet)(fs.insertRows(wb, _, at, count))
        case Edit.DeleteRows(sheet, at, count) =>
          structural(sheet)(fs.deleteRows(wb, _, at, count))
        case Edit.InsertCols(sheet, at, count) =>
          structural(sheet)(fs.insertCols(wb, _, at, count))
        case Edit.DeleteCols(sheet, at, count) =>
          structural(sheet)(fs.deleteCols(wb, _, at, count))

        // ----- workbook -----
        case Edit.AddSheet(sheetName, after, before) =>
          val position = (after, before) match
            case (Some(anchor), _) => indexOf(wb, anchor).map(_ + 1)
            case (_, Some(anchor)) => indexOf(wb, anchor)
            case _ => Right(wb.sheets.size)
          workbookLevel(position.flatMap(wb.insertAt(_, Sheet(sheetName))))

        case Edit.RemoveSheet(sheetName) => workbookLevel(wb.remove(sheetName))

        case Edit.RenameSheet(from, to) => workbookLevel(fs.renameSheet(wb, from, to))

        case Edit.MoveSheet(sheetName, toIndex, after, before) =>
          val rest = wb.sheets.map(_.name).filterNot(_ == sheetName)
          def relativeTo(anchor: SheetName, offset: Int): XLResult[Int] =
            if anchor == sheetName then
              Left(XLError.Other(s"$name: cannot move a sheet relative to itself"))
            else
              rest.indexOf(anchor) match
                case -1 => Left(XLError.SheetNotFound(anchor.value))
                case i => Right(i + offset)
          val position: XLResult[Int] = (toIndex, after, before) match
            case (Some(i), _, _) =>
              if i >= 0 && i <= rest.size then Right(i)
              else Left(XLError.OutOfBounds(s"$name[$i]", s"Valid range: 0 to ${rest.size}"))
            case (_, Some(anchor), _) => relativeTo(anchor, 1)
            case (_, _, Some(anchor)) => relativeTo(anchor, 0)
            case _ => Left(XLError.Other(s"$name: exactly one of to, after or before is required"))
          workbookLevel(
            for
              _ <- indexOf(wb, sheetName)
              at <- position
              moved <- wb.reorder(rest.patch(at, Vector(sheetName), 0))
            yield moved
          )

        case Edit.CopySheet(source, target) =>
          workbookLevel(
            findSheet(wb, source).flatMap(s => wb.insertAt(wb.sheets.size, s.copy(name = target)))
          )

        case Edit.HideSheet(sheetName, veryHidden) =>
          workbookLevel(
            wb.setSheetState(sheetName, Some(if veryHidden then "veryHidden" else "hidden"))
          )
        case Edit.ShowSheet(sheetName) => workbookLevel(wb.setSheetState(sheetName, None))

        case Edit.DefineName(defined, refersTo, nameScope) =>
          workbookLevel(localSheetId(wb, nameScope).map { localId =>
            val others =
              wb.metadata.definedNames.filterNot(d =>
                d.name == defined && d.localSheetId == localId
              )
            withDefinedNames(wb, others :+ DefinedName(defined, refersTo, localSheetId = localId))
          })

        case Edit.RemoveName(defined, nameScope) =>
          workbookLevel(localSheetId(wb, nameScope).flatMap { localId =>
            val (matching, others) =
              wb.metadata.definedNames.partition(d =>
                d.name == defined && d.localSheetId == localId
              )
            if matching.isEmpty then Left(XLError.Other(s"Named range '$defined' not found"))
            else Right(withDefinedNames(wb, others))
          })
    }

  // ===== shared steps =====

  /** The formula text as stored: no leading `=`, no surrounding whitespace. */
  private def formulaText(formula: String): String = formula.trim.stripPrefix("=").trim

  /**
   * `Sheet.put` plus the hint's style step, mirroring `Sheet.putSingle`: the hint resolves against
   * the cell's current style (or the default) and the result is registered and set — so a `General`
   * hint stamps the default style exactly as a typed `put(ref, 42)` does.
   */
  private def putHinted(
    sheet: Sheet,
    ref: ARef,
    value: CellValue,
    hint: Option[FormatHint]
  ): Sheet =
    val withValue = sheet.put(ref, value)
    hint.fold(withValue) { h =>
      val existing = withValue.getCellStyle(ref).getOrElse(CellStyle.default)
      withValue.withCellStyle(ref, h.resolve(existing))
    }

  /** The style a cell ends with under an overlay: onto its own style, or onto the default. */
  private def styledCell(
    sheet: Sheet,
    ref: ARef,
    overlay: StyleOverlay,
    mode: StyleMode
  ): CellStyle =
    val base = mode match
      case StyleMode.Merge => sheet.getCellStyle(ref).getOrElse(CellStyle.default)
      case StyleMode.Replace => CellStyle.default
    overlay.applyTo(base)

  private def applyStyle(
    sheet: Sheet,
    range: CellRange,
    overlay: StyleOverlay,
    mode: StyleMode
  ): Sheet =
    range.cells.foldLeft(sheet)((acc, ref) =>
      acc.withCellStyle(ref, styledCell(acc, ref, overlay, mode))
    )

  private def columns(sheet: Sheet, cols: ColSpan)(
    f: com.tjclp.xl.sheets.ColumnProperties => com.tjclp.xl.sheets.ColumnProperties
  ): Sheet =
    cols.columns.foldLeft(sheet)((s, c) => s.setColumnProperties(c, f(s.getColumnProperties(c))))

  private def rowsOf(sheet: Sheet, rows: RowSpan)(
    f: com.tjclp.xl.sheets.RowProperties => com.tjclp.xl.sheets.RowProperties
  ): Sheet =
    rows.rows.foldLeft(sheet)((s, r) => s.setRowProperties(r, f(s.getRowProperties(r))))

  private def localSheetId(wb: Workbook, scope: Option[SheetName]): XLResult[Option[Int]] =
    scope.fold[XLResult[Option[Int]]](Right(None))(s => indexOf(wb, s).map(Some(_)))

  private def withDefinedNames(wb: Workbook, names: Vector[DefinedName]): Workbook =
    wb.copy(
      metadata = wb.metadata.copy(definedNames = names),
      sourceContext = wb.sourceContext.map(_.markMetadataModified)
    )

  // ===== lower =====

  /**
   * The Sheet-local subset as a `Patch` over `existing` (the sheet before the edit), or `None` when
   * the edit needs the formula support, another sheet, the workbook, or a state the kernel has no
   * case for (comments removed, panes, views, print setup). "Another sheet" is decided by the
   * edit's own qualifier: an edit that names a sheet other than `existing.name` lowers to `None`,
   * so a caller can never apply `Other!A1` onto `Data` through the kernel. Total: an edit whose
   * guards would fail lowers to `None` rather than to a patch the kernel would throw on. Coherence
   * law (EditLawsSpec): `lower(e, s) == Some(p)` and `e` applying to `s` imply the result is
   * `Patch.applyPatch(s, p)`.
   */
  def lower(edit: Edit, existing: Sheet): Option[Patch] =
    if qualifierOf(edit).exists(_ != existing.name) then None
    else lowerOnto(edit, existing)

  /**
   * The sheet an edit names for itself, or `None` for "the scope's default" and for the
   * workbook-level cases. The field differs per family: `Loc`/`Area` targets carry `sheet`, `Fill`
   * names it on its source, `Copy` on its target (where the write lands), the span and
   * sheet-setting families on `sheet`.
   */
  private def qualifierOf(edit: Edit): Option[SheetName] = edit match
    case Edit.Put(at, _, _) => at.sheet
    case Edit.PutValues(at, _, _) => at.sheet
    case Edit.PutFormula(at, _, _) => at.sheet
    case Edit.PutFormulas(at, _, _) => at.sheet
    case Edit.DragFormula(at, _, _, _) => at.sheet
    case Edit.Fill(source, _, _) => source.sheet
    case Edit.Copy(_, target, _) => target.sheet
    case Edit.Sort(at, _, _) => at.sheet
    case Edit.Clear(at, _) => at.sheet
    case Edit.Style(at, _, _) => at.sheet
    case Edit.Merge(at) => at.sheet
    case Edit.Unmerge(at) => at.sheet
    case Edit.ColWidth(sheet, _, _) => sheet
    case Edit.RowHeight(sheet, _, _) => sheet
    case Edit.HideCols(sheet, _) => sheet
    case Edit.ShowCols(sheet, _) => sheet
    case Edit.HideRows(sheet, _) => sheet
    case Edit.ShowRows(sheet, _) => sheet
    case Edit.GroupRows(sheet, _, _, _) => sheet
    case Edit.GroupCols(sheet, _, _, _) => sheet
    case Edit.UngroupRows(sheet, _) => sheet
    case Edit.UngroupCols(sheet, _) => sheet
    case Edit.AutoFit(sheet, _) => sheet
    case Edit.SetComment(at, _) => at.sheet
    case Edit.RemoveComment(at) => at.sheet
    case Edit.Hyperlink(at, _) => at.sheet
    case Edit.AddConditionalFormat(sheet, _, _) => sheet
    case Edit.AddChart(sheet, _, _) => sheet
    case Edit.AddImage(sheet, _, _) => sheet
    case Edit.Freeze(at) => at.sheet
    case Edit.Unfreeze(sheet) => sheet
    case Edit.SetSheetView(sheet, _, _, _) => sheet
    case Edit.SetTabColor(sheet, _) => sheet
    case Edit.SetAutoFilter(sheet, _) => sheet
    case Edit.SetPageSetup(sheet, _, _, _, _, _) => sheet
    case Edit.SetHeaderFooter(sheet, _, _, _, _, _, _, _, _) => sheet
    case Edit.InsertRows(sheet, _, _) => sheet
    case Edit.DeleteRows(sheet, _, _) => sheet
    case Edit.InsertCols(sheet, _, _) => sheet
    case Edit.DeleteCols(sheet, _, _) => sheet
    case Edit.AddSheet(_, _, _) | Edit.RemoveSheet(_) | Edit.RenameSheet(_, _) |
        Edit.MoveSheet(_, _, _, _) | Edit.CopySheet(_, _) | Edit.HideSheet(_, _) |
        Edit.ShowSheet(_) | Edit.DefineName(_, _, _) | Edit.RemoveName(_, _) =>
      None

  /** [[lower]] once the qualifier is known to be `existing` (or absent). */
  private def lowerOnto(edit: Edit, existing: Sheet): Option[Patch] = edit match
    case Edit.Put(at, value, format) => Some(putPatch(existing, at.ref, value, format))
    case Edit.PutValues(at, values, format) if at.range.cellCount == values.size.toLong =>
      Some(
        Patch.Batch(
          at.range.cellsRowMajor
            .zip(values.iterator)
            .map((ref, v) => putPatch(existing, ref, v, format))
            .toVector
        )
      )
    case Edit.PutFormula(at, formula, format) =>
      Some(putPatch(existing, at.ref, CellValue.Formula(formulaText(formula), None), format))
    case Edit.PutFormulas(at, formulas, format) if at.range.cellCount == formulas.size.toLong =>
      Some(
        Patch.Batch(
          at.range.cellsRowMajor
            .zip(formulas.iterator)
            .map((ref, f) =>
              putPatch(existing, ref, CellValue.Formula(formulaText(f), None), format)
            )
            .toVector
        )
      )
    case Edit.Clear(at, what) if !what.comments && !what.isEmpty =>
      val range = at.range
      val contents =
        if what.contents then
          Patch.RemoveRange(range) +: existing.mergedRanges
            .filter(_.intersects(range))
            .toVector
            .sortBy(_.toA1)
            .map(Patch.Unmerge.apply)
        else Vector.empty
      // Styles clear only what contents left behind: after a removal nothing in range remains.
      val styles =
        if what.styles && !what.contents then
          existing.cells.keys
            .filter(range.contains)
            .toVector
            .sortBy(_.toA1)
            .map(Patch.ClearStyle.apply)
        else Vector.empty
      Some(Patch.Batch(contents ++ styles))
    case Edit.Style(at, overlay, mode)
        if overlay.validate.isRight && !(overlay.isEmpty && mode == StyleMode.Merge) =>
      Some(
        Patch.Batch(
          at.range.cells
            .map(ref => Patch.SetCellStyle(ref, styledCell(existing, ref, overlay, mode)))
            .toVector
        )
      )
    case Edit.Merge(at) => Some(Patch.Merge(at.range))
    case Edit.Unmerge(at) => Some(Patch.Unmerge(at.range))
    case Edit.ColWidth(_, cols, width) if width >= 0 && width <= MaxColumnWidth =>
      Some(columnPatch(existing, cols)(_.copy(width = Some(width))))
    case Edit.RowHeight(_, rows, height) if height >= 0 && height <= MaxRowHeight =>
      Some(rowPatch(existing, rows)(_.copy(height = Some(height))))
    case Edit.HideCols(_, cols) => Some(columnPatch(existing, cols)(_.copy(hidden = true)))
    case Edit.ShowCols(_, cols) => Some(columnPatch(existing, cols)(_.copy(hidden = false)))
    case Edit.HideRows(_, rows) => Some(rowPatch(existing, rows)(_.copy(hidden = true)))
    case Edit.ShowRows(_, rows) => Some(rowPatch(existing, rows)(_.copy(hidden = false)))
    case Edit.GroupRows(_, rows, level, collapsed) if SheetEdits.validateLevel(level).isRight =>
      val members = rowPatch(existing, rows)(p =>
        p.copy(outlineLevel = Some(level), hidden = collapsed || p.hidden)
      )
      val summary =
        if collapsed && rows.end.index0 < Row.MaxIndex0 then
          Vector(rowPatch(existing, RowSpan.single(rows.end + 1))(_.copy(collapsed = true)))
        else Vector.empty
      Some(Patch.Batch(members +: summary))
    case Edit.GroupCols(_, cols, level, collapsed) if SheetEdits.validateLevel(level).isRight =>
      val members = columnPatch(existing, cols)(p =>
        p.copy(outlineLevel = Some(level), hidden = collapsed || p.hidden)
      )
      val summary =
        if collapsed && cols.end.index0 < Column.MaxIndex0 then
          Vector(columnPatch(existing, ColSpan.single(cols.end + 1))(_.copy(collapsed = true)))
        else Vector.empty
      Some(Patch.Batch(members +: summary))
    case Edit.UngroupRows(_, rows) =>
      val members = rowPatch(existing, rows)(_.copy(outlineLevel = None, collapsed = false))
      val summary =
        if rows.end.index0 < Row.MaxIndex0 then
          Vector(rowPatch(existing, RowSpan.single(rows.end + 1))(_.copy(collapsed = false)))
        else Vector.empty
      Some(Patch.Batch(members +: summary))
    case Edit.UngroupCols(_, cols) =>
      val members = columnPatch(existing, cols)(_.copy(outlineLevel = None, collapsed = false))
      val summary =
        if cols.end.index0 < Column.MaxIndex0 then
          Vector(columnPatch(existing, ColSpan.single(cols.end + 1))(_.copy(collapsed = false)))
        else Vector.empty
      Some(Patch.Batch(members +: summary))
    case Edit.AutoFit(_, cols) =>
      val targets = cols.fold(
        existing.usedRange.fold(Vector.empty[Column])(r =>
          (r.colStart.index0 to r.colEnd.index0).map(Column.from0).toVector
        )
      )(_.columns)
      Some(
        Patch.Batch(
          targets.map(c =>
            Patch.SetColumnProperties(
              c,
              existing.getColumnProperties(c).copy(width = Some(existing.autoFitWidth(c)))
            )
          )
        )
      )
    case Edit.SetComment(at, comment) => Some(Patch.SetComment(at.ref, comment))
    case Edit.AddConditionalFormat(_, ranges, rules) if ranges.nonEmpty && rules.nonEmpty =>
      Some(Patch.SetConditionalFormat(ranges, rules))
    case _ => None

  private def putPatch(
    existing: Sheet,
    ref: ARef,
    value: CellValue,
    hint: Option[FormatHint]
  ): Patch =
    hint match
      case None => Patch.Put(ref, value)
      case Some(h) =>
        val style = existing.getCellStyle(ref).getOrElse(CellStyle.default)
        Patch.Batch(Vector(Patch.Put(ref, value), Patch.SetCellStyle(ref, h.resolve(style))))

  private def columnPatch(existing: Sheet, cols: ColSpan)(
    f: com.tjclp.xl.sheets.ColumnProperties => com.tjclp.xl.sheets.ColumnProperties
  ): Patch =
    Patch.Batch(
      cols.columns.map(c => Patch.SetColumnProperties(c, f(existing.getColumnProperties(c))))
    )

  private def rowPatch(existing: Sheet, rows: RowSpan)(
    f: com.tjclp.xl.sheets.RowProperties => com.tjclp.xl.sheets.RowProperties
  ): Patch =
    Patch.Batch(rows.rows.map(r => Patch.SetRowProperties(r, f(existing.getRowProperties(r)))))
