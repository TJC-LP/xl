package com.tjclp.xl.ops

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.{CellValue, Comment}
import com.tjclp.xl.cf.CfRule
import com.tjclp.xl.charts.Chart
import com.tjclp.xl.codec.CellWriter
import com.tjclp.xl.drawings.{DrawingAnchor, ImageData}
import com.tjclp.xl.error.XLResult
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.workbooks.Workbook

/**
 * The edit algebra (ADR-017 §2.12): one case per batch op of the Wave 1 `OpRegistry` (the two `put`
 * forms and three `putf` forms are their own cases) plus the verbs that had no batch twin
 * (`insert-rows`, `fill`, `sort`, `remove-sheet`, `add-image`, `name`, `sheets hide`, …). Every
 * mutating surface — batch JSON, the CLI verbs, `wb.edit(...)` in a script — lowers to
 * `Vector[Edit]`, and one interpreter ([[Edit.applyAll]]) gives them one meaning.
 *
 * Targets carry their own sheet ([[Loc]]/[[Area]] hold a qualifier; row/column edits hold `sheet`)
 * and `None` means "the scope's default" — THE sheet rule (ADR-017 §2.5): a qualifier wins, then
 * the scope's default, then the only sheet of a single-sheet book, else `SheetRequired`.
 *
 * The enums that would collide with xl-cli's `FillDirection`/`SortDirection`/`SortMode` are nested
 * in the companion (`Edit.FillDir`, `Edit.SortDir`, `Edit.SortMode`) and are not exported through
 * `api`.
 */
enum Edit derives CanEqual:
  // ----- cell content -----

  /**
   * Put one value. Build it through [[Edit.put]] to lift the codec's number-format hint as
   * [[FormatHint.Inferred]] — a bare `CellValue` never loses inference by accident.
   */
  case Put(at: Loc, value: CellValue, format: Option[FormatHint])

  /** Row-major values over `at`, one per cell; the hint (if any) applies to every cell. */
  case PutValues(at: Area, values: Vector[CellValue], format: Option[FormatHint])

  /** One formula (leading `=` optional), stored uncached. */
  case PutFormula(at: Loc, formula: String, format: Option[FormatHint])

  /** Row-major formulas over `at`, one per cell, stored as written. */
  case PutFormulas(at: Area, formulas: Vector[String], format: Option[FormatHint])

  /** One formula dragged across `at` from `anchor`, shifting relative references like fill-down. */
  case DragFormula(at: Area, formula: String, anchor: ARef, format: Option[FormatHint])

  /** Repeat `source` down or right through `target` (same sheet), shifting relative references. */
  case Fill(source: Area, target: CellRange, direction: Edit.FillDir)

  /**
   * Copy `source` so its top-left lands on `target` (which may name another sheet), shifting
   * relative references by the displacement; `valuesOnly` pastes cached values instead.
   */
  case Copy(source: Area, target: Loc, valuesOnly: Boolean)

  /** Sort the rows of `at` by `keys` (styles and comments move with their rows). */
  case Sort(at: Area, keys: Vector[Edit.SortKeySpec], hasHeader: Boolean)

  /** Clear contents (unmerging overlaps), styles and/or comments in `at`. */
  case Clear(at: Area, what: ClearWhat)

  // ----- style & layout -----

  /** Overlay a partial style onto every cell of `at`, merging or replacing. */
  case Style(at: Area, overlay: StyleOverlay, mode: StyleMode)
  case Merge(at: Area)
  case Unmerge(at: Area)
  case ColWidth(sheet: Option[SheetName], cols: ColSpan, width: Double)
  case RowHeight(sheet: Option[SheetName], rows: RowSpan, height: Double)
  case HideCols(sheet: Option[SheetName], cols: ColSpan)
  case ShowCols(sheet: Option[SheetName], cols: ColSpan)
  case HideRows(sheet: Option[SheetName], rows: RowSpan)
  case ShowRows(sheet: Option[SheetName], rows: RowSpan)

  /** Outline-group rows at `level` (1-7); `collapsed` hides them and marks the summary row. */
  case GroupRows(sheet: Option[SheetName], rows: RowSpan, level: Int, collapsed: Boolean)
  case GroupCols(sheet: Option[SheetName], cols: ColSpan, level: Int, collapsed: Boolean)
  case UngroupRows(sheet: Option[SheetName], rows: RowSpan)
  case UngroupCols(sheet: Option[SheetName], cols: ColSpan)

  /** Auto-fit `cols` (every used column when `None`) to their formatted content. */
  case AutoFit(sheet: Option[SheetName], cols: Option[ColSpan])

  // ----- annotations & objects -----

  case SetComment(at: Loc, comment: Comment)
  case RemoveComment(at: Loc)

  /** Set a hyperlink, or clear it with `None`. */
  case Hyperlink(at: Loc, target: Option[String])

  /** Append one conditional-formatting block over `ranges` (priorities auto-assigned). */
  case AddConditionalFormat(
    sheet: Option[SheetName],
    ranges: Vector[CellRange],
    rules: Vector[CfRule]
  )
  case AddChart(sheet: Option[SheetName], chart: Chart, anchor: DrawingAnchor)
  case AddImage(sheet: Option[SheetName], image: ImageData, anchor: DrawingAnchor)

  // ----- sheet view & print -----

  /** Freeze rows above and columns left of the cell. */
  case Freeze(at: Loc)
  case Unfreeze(sheet: Option[SheetName])
  case SetSheetView(
    sheet: Option[SheetName],
    gridlines: Option[Boolean],
    zoom: Option[Int],
    tabSelected: Option[Boolean]
  )

  /** Set the tab colour, or clear the modeled colour with `None`. */
  case SetTabColor(sheet: Option[SheetName], color: Option[Color])

  /** Set the sheet-level autoFilter, or actively strip it with `None`. */
  case SetAutoFilter(sheet: Option[SheetName], range: Option[CellRange])
  case SetPageSetup(
    sheet: Option[SheetName],
    orientation: Option[String],
    scale: Option[Int],
    fitToWidth: Option[Int],
    fitToHeight: Option[Int],
    fitToPage: Option[Boolean]
  )
  case SetHeaderFooter(
    sheet: Option[SheetName],
    oddHeader: Option[String],
    oddFooter: Option[String],
    evenHeader: Option[String],
    evenFooter: Option[String],
    firstHeader: Option[String],
    firstFooter: Option[String],
    differentOddEven: Boolean,
    differentFirst: Boolean
  )

  // ----- structure (formula references rewritten through FormulaSupport) -----

  case InsertRows(sheet: Option[SheetName], at: Row, count: Int)
  case DeleteRows(sheet: Option[SheetName], at: Row, count: Int)
  case InsertCols(sheet: Option[SheetName], at: Column, count: Int)
  case DeleteCols(sheet: Option[SheetName], at: Column, count: Int)

  // ----- workbook -----

  /** Add an empty sheet at the end, or after/before a named one (at most one of the two). */
  case AddSheet(name: SheetName, after: Option[SheetName], before: Option[SheetName])
  case RemoveSheet(name: SheetName)

  /** Rename a sheet and every reference to it; a rename of the scope's default retargets it. */
  case RenameSheet(from: SheetName, to: SheetName)

  /**
   * Move a sheet to `toIndex` (its final 0-based position among the sheets; the last position is
   * `sheetCount - 1`), or after/before a named one — exactly one of the three.
   */
  case MoveSheet(
    name: SheetName,
    toIndex: Option[Int],
    after: Option[SheetName],
    before: Option[SheetName]
  )
  case CopySheet(source: SheetName, target: SheetName)
  case HideSheet(name: SheetName, veryHidden: Boolean)
  case ShowSheet(name: SheetName)

  /** Add or replace a defined name, workbook-scoped or scoped to `scope`. */
  case DefineName(name: String, refersTo: String, scope: Option[SheetName])
  case RemoveName(name: String, scope: Option[SheetName])

object Edit:

  /** Direction of [[Edit.Fill]]. Nested to keep xl-cli's `FillDirection` unambiguous. */
  enum FillDir derives CanEqual:
    case Down, Right

  /** Direction of one [[SortKeySpec]]. */
  enum SortDir derives CanEqual:
    case Ascending, Descending

  /** Comparison mode of one [[SortKeySpec]]: case-insensitive text, or numbers before text. */
  enum SortMode derives CanEqual:
    case Alphanumeric, Numeric

  /** One sort key: the column (inside the sorted range), its direction and comparison mode. */
  final case class SortKeySpec(column: Column, direction: SortDir, mode: SortMode) derives CanEqual

  object SortKeySpec:
    def ascending(column: Column): SortKeySpec =
      SortKeySpec(column, SortDir.Ascending, SortMode.Alphanumeric)

    def descending(column: Column): SortKeySpec =
      SortKeySpec(column, SortDir.Descending, SortMode.Alphanumeric)

  /**
   * Put a typed value: the codec's format hint (a `LocalDate` → `Date`, a `BigDecimal` → `Decimal`,
   * a `Formatted` → its own format) becomes [[FormatHint.Inferred]], so applying the edit equals
   * `Sheet.put(ref, value)` — inference is never lost by lowering to a `CellValue`.
   */
  def put[A: CellWriter](at: Loc, value: A): Edit =
    val (cellValue, styleHint) = CellWriter[A].write(value)
    Put(at, cellValue, styleHint.map(style => FormatHint.Inferred(style.numFmt)))

  /** [[put]] at a bare cell of the scope's default sheet. */
  def put[A: CellWriter](ref: ARef, value: A): Edit = put(Loc(None, ref), value)

  // ----- the one interpreter (ADR-017 §2.12), reached through the companion -----

  /** The edit-local checks that need no workbook: counts, spans, the guards the model throws on. */
  def validate(edit: Edit)(using FormulaSupport): XLResult[Unit] = EditInterpreter.validate(edit)

  /** The [[Planned]] row of every edit — what [[applyAll]] would touch — without the workbook. */
  def plan(wb: Workbook, edits: Vector[Edit], scope: Scope)(using
    FormulaSupport
  ): XLResult[Vector[Planned]] =
    EditInterpreter.plan(wb, edits, scope)

  /**
   * Apply `edits` in order under `scope`: fail-fast and all-or-nothing, so a `Left` is
   * `EditFailed(index, name, cause)` at the 1-based position that failed and the input workbook is
   * never partially written. Every changed sheet goes back through `Workbook.put` (invariant 1).
   */
  def applyAll(wb: Workbook, edits: Vector[Edit], scope: Scope)(using
    FormulaSupport
  ): XLResult[Applied] =
    EditInterpreter.applyAll(wb, edits, scope)

  /**
   * The Sheet-local subset as a `Patch` over `existing`, or `None` when the edit needs the formula
   * support, another sheet, the workbook, or a state the kernel has no case for. Coherence law
   * (EditLawsSpec): `lower(e, s) == Some(p)` and `e` applying to `s` imply the result equals
   * `Patch.applyPatch(s, p)`.
   */
  def lower(edit: Edit, existing: Sheet): Option[Patch] = EditInterpreter.lower(edit, existing)
