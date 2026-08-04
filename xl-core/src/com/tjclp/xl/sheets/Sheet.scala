package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{ARef, CellRange, Column, RefType, Row, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue, Comment}
import com.tjclp.xl.cf.{CfRule, ConditionalFormat}
import com.tjclp.xl.charts.{Chart, DataRef, Series, SeriesName}
import com.tjclp.xl.codec.{CellCodec, CellWritable, CellWriter}
import com.tjclp.xl.drawings.{AnchorPoint, Drawing, DrawingAnchor, EditAs, Extent, ImageData}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.styles.{CellStyle, StyleRegistry}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.units.StyleId
import com.tjclp.xl.tables.{TableColumn, TableSpec}

import scala.collection.immutable.{Map, Set}
import scala.util.boundary, boundary.break

/**
 * A worksheet containing cells, merged ranges, and properties.
 *
 * Immutable design: all operations return new Sheet instances. Uses persistent data structures for
 * efficient updates.
 *
 * @param freezePane
 *   Freeze pane override with three-valued semantics:
 *   - `None`: preserve existing `<sheetViews>` XML (no change on write)
 *   - `Some(FreezePane.At(ref))`: inject/replace `<pane>` in sheetViews
 *   - `Some(FreezePane.Remove)`: strip any existing `<pane>` from sheetViews
 *
 * The distinction between `None` and `Some(Remove)` matters: `None` is the passive default;
 * `Some(Remove)` is the active intent to remove freeze panes even when the source XML had them.
 * Since GH-372 the reader populates this from frozen `<pane>` elements (anchor from xSplit/ySplit,
 * scroll target from the pane's topLeftCell attribute), so a read→modify→write regenerates the pane
 * from the model and keeps the freeze through the rewrite. Plain SPLIT panes stay unmodeled
 * (`None`) and ride the preserved XML.
 *
 * @param viewSettings
 *   Sheet view settings (gridline visibility, zoom, tab selection). `None` preserves any existing
 *   `<sheetView>` attributes on write; `Some(view)` sets them. The reader populates this whenever a
 *   modeled attribute is present in the source (GH-258/GH-372). Freeze panes and view settings
 *   share a single `<sheetView>` element in the serialized XML.
 *
 * @param drawings
 *   Drawing objects (pictures and preserved fragments, GH-221). Document order is z-order is
 *   emission order: appended drawings paint on top.
 *
 * @param conditionalFormats
 *   Conditional-formatting blocks (GH-136). Document order is emission order; use
 *   [[conditionalFormat]] to append (it assigns priorities) rather than constructing directly.
 *
 * @param tabColor
 *   Sheet tab color (GH-358), serialized as `<sheetPr><tabColor .../>`. `None` preserves any
 *   existing tabColor XML on write (passive default, the viewSettings precedent); `Some(color)`
 *   overlays it onto the preserved sheetPr. Both RGB and theme+tint colors are legal; the reader
 *   populates it whenever the source color is in the typed subset (rgb / theme+tint / indexed).
 *
 * @param dataValidations
 *   Data-validation entries (GH-375). Document order is emission order inside the single
 *   `<dataValidations>` container; use [[withDataValidation]] to append. Unmodeled entries ride
 *   through as [[DataValidation.Preserved]] (the conditionalFormats pattern).
 *
 * @param autoFilter
 *   Sheet-level autoFilter overlay (GH-429, lift-and-overlay tri-state). `None` is the passive
 *   default: any source `<autoFilter>` rides the preservedKnown passthrough verbatim.
 *   `Some(Ranged(r))` overlays the `ref` attribute on write (filterColumn/sortState children ride
 *   verbatim); the reader populates it from a parseable source `@ref` so structural edits keep the
 *   filter attached to its data. `Some(Remove)` actively strips the element (a collapsed range must
 *   not resurrect the stale source filter).
 */
final case class Sheet(
  name: SheetName,
  cells: Map[ARef, Cell] = Map.empty,
  mergedRanges: Set[CellRange] = Set.empty,
  columnProperties: Map[Column, ColumnProperties] = Map.empty,
  rowProperties: Map[Row, RowProperties] = Map.empty,
  defaultColumnWidth: Option[Double] = None,
  defaultRowHeight: Option[Double] = None,
  styleRegistry: StyleRegistry = StyleRegistry.default,
  comments: Map[ARef, Comment] = Map.empty,
  tables: Map[String, TableSpec] = Map.empty,
  pageSetup: Option[PageSetup] = None,
  freezePane: Option[FreezePane] = None,
  viewSettings: Option[SheetView] = None,
  drawings: Vector[Drawing] = Vector.empty,
  conditionalFormats: Vector[ConditionalFormat] = Vector.empty,
  tabColor: Option[Color] = None,
  dataValidations: Vector[DataValidation] = Vector.empty,
  autoFilter: Option[AutoFilterState] = None
):

  /** Get cell at reference (returns empty cell if not present) */
  def apply(ref: ARef): Cell =
    cells.getOrElse(ref, Cell.empty(ref))

  /** Get cell at A1 notation */
  def apply(a1: String): XLResult[Cell] =
    ARef
      .parse(a1)
      .left
      .map(err => XLError.InvalidCellRef(a1, err))
      .map(apply)

  /**
   * Access cell(s) using unified reference type.
   *
   * Sheet-qualified refs (Sales!A1) ignore the sheet name and use only the cell/range part.
   *
   * Returns Cell for single refs, Iterable[Cell] for ranges.
   */
  @annotation.targetName("applyRefType")
  def apply(ref: RefType): Cell | Iterable[Cell] =
    ref match
      case RefType.Cell(cellRef) => apply(cellRef)
      case RefType.Range(range) => getRange(range)
      case RefType.QualifiedCell(_, cellRef) => apply(cellRef)
      case RefType.QualifiedRange(_, range) => getRange(range)

  /** Check if cell exists (not empty) */
  def contains(ref: ARef): Boolean =
    cells.contains(ref)

  /**
   * Put cell at reference (always succeeds - Cell is pre-validated).
   *
   * If the deprecated `Cell.comment` field is set, it is converted into the sheet-level comment
   * store ([[comments]], the store the OOXML writer serializes) as a plain-text comment without
   * author, and the field is cleared on the stored cell. Previously the field silently vanished on
   * write (GH-295); an existing sheet comment at the same ref is overwritten (last write wins). A
   * cell without the field set never touches [[comments]].
   */
  @annotation.nowarn("cat=deprecation") // write-through shim for the deprecated Cell.comment field
  def put(cell: Cell): Sheet =
    cell.comment match
      case Some(text) =>
        copy(
          cells = cells.updated(cell.ref, cell.copy(comment = None)),
          comments = comments.updated(cell.ref, Comment.plainText(text))
        )
      case None =>
        copy(cells = cells.updated(cell.ref, cell))

  // ===== Structural editing: insert/delete rows & columns (GH-128, GH-129) =====
  // PURE cell/merge/property/freeze shifting along one axis. Formula-reference rewriting is layered
  // on top in xl-evaluator (Sheet lives in xl-core, which has no access to the formula engine).

  /**
   * Insert `count` blank rows at 0-based row index `at`; cells in that row and below shift down.
   */
  def insertRows(at: Int, count: Int): Sheet =
    if count <= 0 then this else shiftAxis(rowAxis = true, at, count, deleting = false)

  /**
   * Delete `count` rows from 0-based row index `at`; deleted cells removed, rows below shift up.
   */
  def deleteRows(at: Int, count: Int): Sheet =
    if count <= 0 then this else shiftAxis(rowAxis = true, at, count, deleting = true)

  /** Insert `count` blank columns at 0-based column index `at`; columns at/right shift right. */
  def insertColumns(at: Int, count: Int): Sheet =
    if count <= 0 then this else shiftAxis(rowAxis = false, at, count, deleting = false)

  /**
   * Delete `count` columns from 0-based column index `at`; deleted cells removed, right shifts
   * left.
   */
  def deleteColumns(at: Int, count: Int): Sheet =
    if count <= 0 then this else shiftAxis(rowAxis = false, at, count, deleting = true)

  /**
   * Highest populated 0-based index on one axis — the farthest position a structural insert would
   * shift (GH-472). "Populated" covers every carrier of absolute positions that [[shiftAxis]]
   * rebuilds WITHOUT clamping: cells, comments, row/column properties, drawing cell anchors, and
   * the freeze-pane anchor/scroll target. Range-shaped structures (merges, cf/dv envelopes, print
   * areas, tables, autoFilter) clamp at the sheet edge instead (GH-428) and can never overflow.
   * `None` = nothing populated on the axis. Callers use this to REFUSE an insert that would push
   * any of these past the sheet edge — Excel refuses such files outright, so writing one is never
   * acceptable.
   */
  private[xl] def maxPopulatedIndex(rowAxis: Boolean): Option[Int] =
    def axisOf(ref: ARef): Int = if rowAxis then ref.row.index0 else ref.col.index0
    def anchorCells(anchor: DrawingAnchor): List[ARef] = anchor match
      case DrawingAnchor.OneCell(from, _) => List(from.cell)
      case DrawingAnchor.TwoCell(from, to, _) => List(from.cell, to.cell)
      case _: DrawingAnchor.Absolute => Nil
    val drawingRefs = drawings.iterator.flatMap {
      case Drawing.Picture(anchor, _, _, _) => anchorCells(anchor)
      case Drawing.ChartFrame(anchor, _, _) => anchorCells(anchor)
      case _: Drawing.Preserved => Nil
    }
    val freezeRefs = freezePane.iterator.flatMap {
      case FreezePane.At(topLeftCell, scrolledTo) => topLeftCell :: scrolledTo.toList
      case FreezePane.Remove => Nil
    }
    val propIndices =
      if rowAxis then rowProperties.keysIterator.map(_.index0)
      else columnProperties.keysIterator.map(_.index0)
    ((cells.keysIterator ++ comments.keysIterator ++ drawingRefs ++ freezeRefs).map(axisOf)
      ++ propIndices).maxOption

  /**
   * Shared structural-shift engine for one axis (row or column).
   *
   * Maps a 0-based index on the active axis: insert shifts indices `>= at` by `+count`; delete
   * drops indices in `[at, at+count)` and shifts those `>= at+count` by `-count`. Cells/comments
   * are remapped (dropped cells removed), row/column properties shifted, merged ranges and freeze
   * panes split/clamped/dropped, all on the active axis only.
   */
  private def shiftAxis(rowAxis: Boolean, at: Int, count: Int, deleting: Boolean): Sheet =
    // Index transform on the active axis. None = the index was deleted.
    def idx(i: Int): Option[Int] =
      if !deleting then Some(if i >= at then i + count else i)
      else if i < at then Some(i)
      else if i >= at + count then Some(i - count)
      else None
    // Active-axis index of an ARef, and a rebuilder that replaces it.
    def axisOf(ref: ARef): Int = if rowAxis then ref.row.index0 else ref.col.index0
    def rebuild(ref: ARef, newIdx: Int): ARef =
      if rowAxis then ARef.from0(ref.col.index0, newIdx) else ARef.from0(newIdx, ref.row.index0)

    val newCells = cells.flatMap { case (ref, cell) =>
      idx(axisOf(ref)).map { ni =>
        val nr = rebuild(ref, ni); nr -> cell.copy(ref = nr)
      }
    }
    val newComments = comments.flatMap { case (ref, c) =>
      idx(axisOf(ref)).map(ni => rebuild(ref, ni) -> c)
    }
    val newRowProps =
      if !rowAxis then rowProperties
      else rowProperties.flatMap { case (row, p) => idx(row.index0).map(ni => Row.from0(ni) -> p) }
    val newColProps =
      if rowAxis then columnProperties
      else
        columnProperties.flatMap { case (col, p) =>
          idx(col.index0).map(ni => Column.from0(ni) -> p)
        }

    // Shared range algebra on the active axis (Sheet.shiftSpan): clamp the span at the sheet
    // edge (GH-428) and drop it when it collapses into a deletion or is pushed fully past the
    // edge. Used by merged ranges, conditional-format envelopes, data validations, print areas,
    // tables, and the autoFilter range — every sqref-shaped structure must shift identically.
    val axisMax = if rowAxis then Row.MaxIndex0 else Column.MaxIndex0
    def shiftRangeOnAxis(range: CellRange): Option[CellRange] =
      Sheet
        .shiftSpan(axisOf(range.start), axisOf(range.end), at, count, deleting, axisMax)
        .map((ns, ne) => CellRange(rebuild(range.start, ns), rebuild(range.end, ne)))

    // Merged ranges: clamp the active-axis span; drop if it collapses entirely into the deletion.
    val newMerges = mergedRanges.flatMap(shiftRangeOnAxis)

    // Freeze pane: shift/clamp BOTH the anchor and the scroll target on the active axis
    // (clamp a deleted ref to `at`) — dropping scrolledTo here would silently unscroll the
    // pane on a structural edit (GH-382).
    val newFreeze = freezePane.map {
      case FreezePane.At(ref, scrolledTo) =>
        val ni = idx(axisOf(ref)).getOrElse(at)
        val newScrolled = scrolledTo.map { s =>
          rebuild(s, idx(axisOf(s)).getOrElse(at))
        }
        FreezePane.At(rebuild(ref, ni), newScrolled)
      case other => other
    }

    // Drawings: remap each cell anchor point; a deleted anchor index clamps to `at` — unlike
    // comments, Excel keeps pictures when their anchor row/column is deleted (a fully-deleted
    // TwoCell range degenerates to zero extent). Absolute and Preserved anchors are untouched;
    // editAs-aware size recomputation is deliberately not attempted (GH-221, deferred).
    // ChartFrames additionally shift their SAME-SHEET data references (GH-222) — cross-sheet
    // chart shifts are layered in xl-evaluator's StructuralEditor (which sees all sheets).
    def remapPoint(p: AnchorPoint): AnchorPoint =
      val ni = idx(axisOf(p.cell)).getOrElse(at)
      p.copy(cell = rebuild(p.cell, ni))
    def remapAnchor(anchor: DrawingAnchor): DrawingAnchor = anchor match
      case DrawingAnchor.OneCell(from, extent) =>
        DrawingAnchor.OneCell(remapPoint(from), extent)
      case DrawingAnchor.TwoCell(from, to, editAs) =>
        DrawingAnchor.TwoCell(remapPoint(from), remapPoint(to), editAs)
      case abs: DrawingAnchor.Absolute => abs
    val newDrawings = drawings.map {
      case Drawing.Picture(anchor, image, n, d) =>
        Drawing.Picture(remapAnchor(anchor), image, n, d)
      case Drawing.ChartFrame(anchor, chart, n) =>
        Drawing.ChartFrame(
          remapAnchor(anchor),
          Sheet.shiftChartReferences(chart, name.value, rowAxis, at, count, deleting),
          n
        )
      case preserved: Drawing.Preserved => preserved
    }

    // Conditional formats (GH-136): the EXACT mergedRanges clamp/split/drop algebra on each typed
    // block's envelope. A range collapsing entirely is dropped; a block whose ranges all drop is
    // removed (Excel deletes rules whose entire range is deleted). Rules themselves — including
    // CfRule.Preserved payloads — are untouched here; typed formula rewriting is layered in
    // xl-evaluator's StructuralEditor. ConditionalFormat.Preserved blocks shift their root
    // sqref textually (GH-429, SqrefShift) — an unmovable payload rides unchanged, a fully
    // collapsed one drops.
    val newCondFmts = conditionalFormats.flatMap {
      case cf: ConditionalFormat.Rules =>
        val shiftedRanges = cf.ranges.flatMap(shiftRangeOnAxis)
        if shiftedRanges.isEmpty then None
        else Some(cf.copy(ranges = shiftedRanges))
      case ConditionalFormat.Preserved(xml) =>
        SqrefShift
          .shiftPayload(xml, "conditionalFormatting", shiftRangeOnAxis)
          .map(ConditionalFormat.Preserved.apply)
    }

    // Data validations (GH-375): same envelope algebra — without it, dropdowns detach from their
    // cells on row/column insert or delete. An entry whose ranges all drop is removed (Excel
    // deletes validations whose entire range is deleted); Preserved entries shift their sqref
    // textually like Preserved cf blocks (GH-429).
    val newDataValidations = dataValidations.flatMap {
      case dv: DataValidation.Rules =>
        val shiftedRanges = dv.ranges.flatMap(shiftRangeOnAxis)
        if shiftedRanges.isEmpty then None
        else Some(dv.copy(ranges = shiftedRanges))
      case DataValidation.Preserved(xml) =>
        SqrefShift
          .shiftPayload(xml, "dataValidation", shiftRangeOnAxis)
          .map(DataValidation.Preserved.apply)
    }

    // Print setup (GH-429): the print area shifts on both axes; the repeat-row span only on the
    // row axis (via the same shiftSpan clamp, 1-based <-> 0-based at the boundary). A collapsed
    // area/span clears its field — a documented divergence from Excel's #REF!-name (the model
    // has no error state for print names).
    val newPageSetup = pageSetup.map { ps =>
      ps.copy(
        printArea = ps.printArea.flatMap(shiftRangeOnAxis),
        repeatRows =
          if !rowAxis then ps.repeatRows
          else
            ps.repeatRows.flatMap((s, e) =>
              Sheet
                .shiftSpan(s - 1, e - 1, at, count, deleting, Row.MaxIndex0)
                .map((ns, ne) => (ns + 1, ne + 1))
            )
      )
    }

    // Tables (GH-429): shift each spec's range; drop the table when the range collapses or
    // shrinks below header+totals+one-data-row (never emit a table Excel must repair). Row-axis
    // edits keep columns; column-axis edits retabulate them (drop deleted columns, splice
    // synthesized ones for interior inserts) and stamp synthesized header names into their
    // header-row cells — Excel itself writes "Column1" into the header cell, and a name/cell
    // mismatch risks table repair.
    val (newTables, headerStamps) =
      tables.foldLeft((Map.empty[String, TableSpec], Vector.empty[(ARef, Cell)])) {
        case ((accTables, accStamps), (key, spec)) =>
          val shifted = shiftRangeOnAxis(spec.range).flatMap { newRange =>
            val minHeight =
              (if spec.showHeaderRow then 1 else 0) + (if spec.showTotalsRow then 1 else 0) + 1
            if newRange.height < minHeight then None
            else if rowAxis then Some((spec.copy(range = newRange), Vector.empty))
            else Sheet.retabulateColumns(spec, newRange, at, count, deleting)
          }
          shifted match
            case Some((newSpec, freshCols)) =>
              val stamps =
                if newSpec.showHeaderRow then
                  freshCols.map { (colIdx, colName) =>
                    val ref = ARef.from0(colIdx, newSpec.range.start.row.index0)
                    ref -> Cell(ref, CellValue.Text(colName))
                  }
                else Vector.empty
              (accTables.updated(key, newSpec), accStamps ++ stamps)
            case None => (accTables, accStamps)
      }

    // Sheet-level autoFilter overlay (GH-429): Ranged shifts like every sqref; a collapse becomes
    // an active Remove (a plain None would resurrect the stale source filter on write).
    val newAutoFilter = autoFilter.map {
      case AutoFilterState.Ranged(r) =>
        shiftRangeOnAxis(r).fold(AutoFilterState.Remove)(AutoFilterState.Ranged.apply)
      case AutoFilterState.Remove => AutoFilterState.Remove
    }

    copy(
      cells = newCells ++ headerStamps,
      comments = newComments,
      rowProperties = newRowProps,
      columnProperties = newColProps,
      mergedRanges = newMerges,
      freezePane = newFreeze,
      drawings = newDrawings,
      conditionalFormats = newCondFmts,
      dataValidations = newDataValidations,
      pageSetup = newPageSetup,
      tables = newTables,
      autoFilter = newAutoFilter
    )

  /** Put CellValue at reference (always succeeds - CellValue is pre-validated) */
  def put(ref: ARef, value: CellValue): Sheet =
    val updatedCell = cells.get(ref) match
      case Some(existing) => existing.withValue(value)
      case None => Cell(ref, value)
    put(updatedCell)

  /**
   * Put a single value at reference.
   *
   * Accepts any supported type: String, Int, Long, Double, BigDecimal, Boolean, LocalDate,
   * LocalDateTime, RichText, Formatted, CellValue. Returns XLResult for consistent error handling.
   *
   * '''Type safety''': Uses compile-time checked `CellWriter` type class. Unsupported types fail at
   * compile time, not runtime.
   *
   * Examples:
   * {{{
   * sheet.put(ref"A1", "Hello")           // String
   * sheet.put(ref"A1", 42)                // Int
   * sheet.put(ref"A1", money"$$100")      // Formatted
   * sheet.put(ref"A1", fx"=SUM(B1:B10)")  // Formula
   * }}}
   */
  def put[A: CellWriter](ref: ARef, value: A): Sheet =
    putSingle(ref, value)

  /**
   * Put a single value at string reference with compile-time or runtime validation.
   *
   * Uses `transparent inline` to enable '''type narrowing''' based on the argument:
   *   - '''String literal''' ("A1"): Validated at compile time, returns `Sheet` directly
   *   - '''Runtime expression''' (variable): Validated at runtime, returns `XLResult[Sheet]`
   *
   * The union return type `Sheet | XLResult[Sheet]` allows the compiler to narrow to the
   * appropriate type at each call site, providing both type safety and ergonomics.
   *
   * @example
   *   {{{
   *   // Literal string → compile-time validation → Sheet
   *   val s1: Sheet = sheet.put("A1", "Hello")
   *
   *   // Runtime string → runtime validation → XLResult[Sheet]
   *   val name = getUserInput()
   *   val s2: XLResult[Sheet] = sheet.put(name, 42)
   *   s2 match
   *     case Right(s) => // valid
   *     case Left(err) => // invalid reference
   *   }}}
   */
  @annotation.targetName("putString")
  transparent inline def put[A](inline ref: String, value: A)(using
    inline cw: CellWriter[A]
  ): Sheet | XLResult[Sheet] =
    ${ com.tjclp.xl.macros.PutLiteral.putImpl('{ this }, 'ref, 'value, 'cw) }

  /**
   * Put a single value at reference with explicit style.
   *
   * Merges explicit style with codec-inferred style: explicit properties take precedence, but if
   * explicit has General NumFmt and codec provides non-General, codec's NumFmt is used.
   */
  def put[A: CellWriter](ref: ARef, value: A, style: CellStyle): Sheet =
    val s = putSingle(ref, value)
    // Get the codec-inferred style (if any) from the cell after putSingle
    val codecStyle = s.cells.get(ref).flatMap(_.styleId).flatMap(s.styleRegistry.get)
    // Merge: explicit style properties take precedence, but use codec NumFmt if explicit is General
    val mergedStyle = codecStyle match
      case Some(cs)
          if style.numFmt == com.tjclp.xl.styles.numfmt.NumFmt.General && cs.numFmt != com.tjclp.xl.styles.numfmt.NumFmt.General =>
        style.copy(numFmt = cs.numFmt)
      case _ => style
    import com.tjclp.xl.sheets.styleSyntax.withCellStyle
    s.withCellStyle(ref, mergedStyle)

  /**
   * Put a single value at string reference with explicit style.
   *
   * For string literals, validates at compile time and returns `Sheet` directly. For runtime
   * strings, validates at runtime and returns `XLResult[Sheet]`.
   */
  @annotation.targetName("putStringStyled")
  transparent inline def put[A](inline ref: String, value: A, style: CellStyle)(using
    inline cw: CellWriter[A]
  ): Sheet | XLResult[Sheet] =
    ${ com.tjclp.xl.macros.PutLiteral.putStyledImpl('{ this }, 'ref, 'value, 'style, 'cw) }

  // Merge existing style with codec-inferred style
  // Preserves existing properties; codec NumFmt overrides only when existing is General
  // Rationale: If user explicitly set Currency format, keep it. If just Bold (General), apply type-appropriate format.
  private def mergeStyles(existing: CellStyle, codec: CellStyle): CellStyle =
    import com.tjclp.xl.styles.numfmt.NumFmt
    if existing.numFmt == NumFmt.General && codec.numFmt != NumFmt.General then
      existing.copy(numFmt = codec.numFmt)
    else existing

  // Internal helper for single-cell put with CellWriter type class
  // Uses the CellWriter[CellWritable] instance which handles all supported types via pattern matching
  // Returns Sheet directly (infallible) since CellWriter.write cannot fail
  // GH-297: style registration and the cell update are fused into a single copy (register once,
  // set the styleId directly). The previous register + withCellStyle sequence recomputed the
  // canonical key and copied the sheet three times per styled put.
  private def putSingle[A: CellWriter](ref: ARef, value: A): Sheet =
    import com.tjclp.xl.codec.given
    val (cellValue, styleOpt) = CellWriter[A].write(value)
    val existingCell = cells.get(ref)
    val updatedCell = existingCell match
      case Some(existing) => existing.withValue(cellValue)
      case None => Cell(ref, cellValue)
    styleOpt match
      case Some(codecStyle) =>
        val mergedStyle = existingCell.flatMap(_.styleId).flatMap(styleRegistry.get) match
          case Some(existingStyle) => mergeStyles(existingStyle, codecStyle)
          case None => codecStyle
        val (newRegistry, styleId) = styleRegistry.register(mergedStyle)
        copy(
          styleRegistry = newRegistry,
          cells = cells.updated(ref, updatedCell.withStyle(styleId))
        )
      case None =>
        copy(cells = cells.updated(ref, updatedCell))

  /**
   * Batch put with mixed value types and automatic style inference.
   *
   * Accepts (ARef, A) pairs where A is any type with a CellWriter instance. Auto-infers styles
   * based on value types (dates get date format, decimals get number format, etc.). Formatted
   * literals (money"", date"", percent"") preserve their NumFmt.
   *
   * '''Type safety''': Due to contravariance of `CellWriter[-A]` and the master
   * `CellWriter[CellWritable]` instance, heterogeneous types like `String | Int | LocalDate` are
   * accepted and checked at compile time.
   *
   * This is the recommended API for large batch upserts due to its clean, token-efficient syntax:
   * {{{
   * sheet.put(
   *   ref"A1" -> "Revenue",
   *   ref"B1" -> LocalDate.of(2025, 11, 10),
   *   ref"C1" -> money"$$1,234.56"
   * ).unsafe
   * }}}
   *
   * Supported types: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime,
   * RichText, Formatted, CellValue. Unsupported types fail at compile time.
   *
   * For demos/REPLs, use .unsafe (requires explicit import):
   * {{{
   * import com.tjclp.xl.unsafe.*
   * sheet.put(ref"A1" -> "Hello").unsafe
   * }}}
   *
   * @param updates
   *   Varargs of (ARef, A) pairs
   * @return
   *   Updated sheet (always succeeds - type safety is enforced at compile time)
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def put[A: CellWriter](updates: (ARef, A)*): Sheet =
    import com.tjclp.xl.codec.given

    // NOTE: Local mutation for performance - buffers are private to this method
    // and never escape. The function remains pure (referentially transparent) because:
    // 1. All mutations are confined to this scope
    // 2. No shared mutable state is accessed
    // 3. Output depends only on inputs (deterministic)
    // This is a common FP optimization pattern for bulk operations (similar to Scala stdlib).

    // Single-pass: build cells and collect resolved style ids simultaneously (GH-297: the
    // style is registered once here; the previous per-cell withCellStyle fold re-registered
    // each style and copied the whole sheet once per styled cell)
    val builtCells = scala.collection.mutable.ArrayBuffer[Cell]()
    val cellsWithStyles = scala.collection.mutable.ArrayBuffer[(ARef, StyleId)]()
    var registry = styleRegistry
    val writer = CellWriter[A]

    updates.foreach { (ref, value) =>
      val (cellValue, styleOpt) = writer.write(value)
      val existingCell = this.cells.get(ref)
      val updatedCell = existingCell match
        case Some(existing) => existing.withValue(cellValue)
        case None => Cell(ref, cellValue)
      builtCells += updatedCell
      styleOpt.foreach { codecStyle =>
        val mergedStyle = existingCell.flatMap(_.styleId).flatMap(this.styleRegistry.get) match
          case Some(existingStyle) => mergeStyles(existingStyle, codecStyle)
          case None => codecStyle
        val (newRegistry, styleId) = registry.register(mergedStyle)
        registry = newRegistry
        cellsWithStyles += ((ref, styleId))
      }
    }

    applyBulkCells(builtCells, cellsWithStyles, registry)

  /**
   * Type-safe variant of [[put]] that requires a single `CellCodec` for all values.
   *
   * This avoids runtime type inspection and lets the compiler inline the codec implementation,
   * ensuring zero-overhead writes when the value type is known statically.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def putTyped[A](updates: (ARef, A)*)(using CellCodec[A]): Sheet =
    val builtCells = scala.collection.mutable.ArrayBuffer[Cell]()
    val cellsWithStyles = scala.collection.mutable.ArrayBuffer[(ARef, StyleId)]()
    var registry = styleRegistry
    val codec = summon[CellCodec[A]]

    updates.foreach { (ref, value) =>
      val (cellValue, styleOpt) = codec.write(value)
      val existingCell = this.cells.get(ref)
      val updatedCell = existingCell match
        case Some(existing) => existing.withValue(cellValue)
        case None => Cell(ref, cellValue)
      builtCells += updatedCell
      styleOpt.foreach { codecStyle =>
        val mergedStyle = existingCell.flatMap(_.styleId).flatMap(this.styleRegistry.get) match
          case Some(existingStyle) => mergeStyles(existingStyle, codecStyle)
          case None => codecStyle
        val (newRegistry, styleId) = registry.register(mergedStyle)
        registry = newRegistry
        cellsWithStyles += ((ref, styleId))
      }
    }

    applyBulkCells(builtCells, cellsWithStyles, registry)

  /**
   * Batch put multiple values using string references with compile-time validation.
   *
   * When all refs are string literals, validates them at compile time and returns `Sheet` directly.
   * When any ref is a runtime expression, falls back to runtime parsing and returns
   * `XLResult[Sheet]`.
   *
   * This enables the clean map syntax without requiring the `ref"..."` macro:
   * {{{
   * // All literals → returns Sheet (compile-time validated)
   * val sheet = Sheet("Demo").put(
   *   "A1" -> "Revenue",
   *   "B1" -> 100,
   *   "C1" -> fx"=A1+B1"
   * )
   *
   * // Runtime ref → returns XLResult[Sheet]
   * val col = "A"
   * val result = Sheet("Demo").put(s"$${col}1" -> "Dynamic")
   * }}}
   */
  @annotation.targetName("putStringTuples")
  transparent inline def put[A](inline updates: (String, A)*)(using
    inline cw: CellWriter[A]
  ): Sheet | XLResult[Sheet] =
    ${ com.tjclp.xl.macros.PutLiteral.putTuplesImpl('{ this }, 'updates, 'cw) }

  // GH-297: styles arrive pre-registered as (ref, styleId) — apply ids to the merged cell map
  // directly instead of folding withCellStyle (which re-registered every style and copied the
  // whole sheet once per styled cell). The id fold preserves the pre-existing duplicate-ref
  // semantics: the last value put wins, and any styled entry for that ref re-applies its id.
  private def applyBulkCells(
    builtCells: Iterable[Cell],
    styled: Iterable[(ARef, StyleId)],
    newRegistry: StyleRegistry
  ): Sheet =
    val baseCells = this.cells ++ builtCells.iterator.map(cell => cell.ref -> cell)
    val finalCells = styled.foldLeft(baseCells) { case (cs, (ref, styleId)) =>
      cs.get(ref).fold(cs)(cell => cs.updated(ref, cell.withStyle(styleId)))
    }
    copy(styleRegistry = newRegistry, cells = finalCells)

  /**
   * Apply a patch to this sheet.
   *
   * Patches enable declarative composition of updates (Put, SetStyle, Merge, etc.). This operation
   * is infallible since patches contain only validated references.
   *
   * Example:
   * {{{
   * val patch = (ref"A1" := "Title") ++ range"A1:C1".merge
   * val updated = sheet.put(patch)
   * }}}
   *
   * @param patch
   *   The patch to apply
   * @return
   *   The updated sheet
   */
  def put(patch: com.tjclp.xl.patch.Patch): Sheet =
    com.tjclp.xl.patch.Patch.applyPatch(this, patch)

  /** Remove cell at reference */
  def remove(ref: ARef): Sheet =
    copy(cells = cells.removed(ref))

  /** Remove all cells in range */
  def removeRange(range: CellRange): Sheet =
    // Use range.contains() instead of materializing to Set - O(n) where n = existing cells
    // This avoids iterating 1M+ cells for full column/row references like A:A or 1:1
    copy(cells = cells.filterNot((ref, _) => range.contains(ref)))

  /** Get all cells in a range */
  def getRange(range: CellRange): Iterable[Cell] =
    range.cells.flatMap(ref => cells.get(ref)).toSeq

  /** Merge cells in range */
  def merge(range: CellRange): Sheet =
    copy(mergedRanges = mergedRanges + range)

  /** Unmerge cells in range */
  def unmerge(range: CellRange): Sheet =
    copy(mergedRanges = mergedRanges - range)

  /** Check if cell is part of a merged range */
  def isMerged(ref: ARef): Boolean =
    mergedRanges.exists(_.contains(ref))

  /** Get merged range containing ref (if any) */
  def getMergedRange(ref: ARef): Option[CellRange] =
    mergedRanges.find(_.contains(ref))

  /** Add comment to cell */
  def comment(ref: ARef, comment: Comment): Sheet =
    copy(comments = comments.updated(ref, comment))

  /**
   * Add comment to cell (string variant).
   *
   * When called with a string literal, the cell reference is validated at compile time and returns
   * `Sheet` directly. Invalid literals fail to compile. Runtime strings return `XLResult[Sheet]`.
   */
  @annotation.targetName("commentString")
  transparent inline def comment(inline ref: String, cmt: Comment): Sheet | XLResult[Sheet] =
    ${ com.tjclp.xl.macros.PutLiteral.commentImpl('{ this }, 'ref, 'cmt) }

  /** Get comment at cell reference */
  def getComment(ref: ARef): Option[Comment] =
    comments.get(ref)

  /** Remove comment from cell */
  def removeComment(ref: ARef): Sheet =
    copy(comments = comments.removed(ref))

  /** Check if cell has a comment */
  def hasComment(ref: ARef): Boolean =
    comments.contains(ref)

  // ===== Drawings: embedded images (GH-221) =====

  /** Add an image with full anchor control (total). */
  def addImage(image: ImageData, anchor: DrawingAnchor): Sheet =
    copy(drawings = drawings :+ Drawing.Picture(anchor, image))

  /**
   * Add an image one-cell-anchored at `at`, sized to its natural pixel dimensions at 96 DPI.
   * Returns Left when the dimensions cannot be sniffed from the bytes (Tiff/Emf/Wmf or malformed
   * headers) — no silent size guessing; pass an explicit [[Extent]] instead.
   */
  def addImage(image: ImageData, at: ARef): XLResult[Sheet] =
    image.naturalExtent match
      case Some(extent) => Right(addImage(image, DrawingAnchor.at(at, extent)))
      case None =>
        Left(
          XLError.ParseError(
            "image bytes",
            s"cannot sniff natural size for ${image.format} image — use addImage(image, at, extent)"
          )
        )

  /** Add an image one-cell-anchored at `at` with an explicit extent (total). */
  def addImage(image: ImageData, at: ARef, extent: Extent): Sheet =
    addImage(image, DrawingAnchor.at(at, extent))

  /** Add an image two-cell-anchored over `range` (total). */
  def addImage(image: ImageData, range: CellRange, editAs: EditAs = EditAs.TwoCell): Sheet =
    addImage(image, DrawingAnchor.over(range, editAs))

  /** All typed pictures in z-order (Preserved fragments excluded). */
  def pictures: Vector[Drawing.Picture] =
    drawings.collect { case p: Drawing.Picture => p }

  // ===== Drawings: typed charts (GH-222) =====

  /** Add a chart with full anchor control (total). */
  def addChart(chart: Chart, anchor: DrawingAnchor): Sheet =
    copy(drawings = drawings :+ Drawing.ChartFrame(anchor, chart))

  /** Add a chart two-cell-anchored over `range` (total). */
  def addChart(chart: Chart, over: CellRange, editAs: EditAs = EditAs.TwoCell): Sheet =
    addChart(chart, DrawingAnchor.over(over, editAs))

  /** All typed charts in z-order (Preserved fragments excluded). */
  def charts: Vector[Drawing.ChartFrame] =
    drawings.collect { case c: Drawing.ChartFrame => c }

  /**
   * Shift this sheet's typed-chart data references that point at sheet `edited` through a
   * structural edit on that sheet (GH-222): `delta > 0` inserts `delta` rows/columns at 0-based
   * `at`, `delta < 0` deletes `-delta`. Anchors are NOT touched — they live on this sheet, whose
   * own geometry did not change. Used by xl-evaluator's StructuralEditor for cross-sheet chart
   * tracking; the edited sheet's own charts are handled inside the structural shift itself.
   */
  def shiftChartRefs(edited: String, isRow: Boolean, at: Int, delta: Int): Sheet =
    if delta == 0 then this
    else
      val deleting = delta < 0
      val count = math.abs(delta)
      val newDrawings = drawings.map {
        case Drawing.ChartFrame(anchor, chart, n) =>
          Drawing.ChartFrame(
            anchor,
            Sheet.shiftChartReferences(chart, edited, isRow, at, count, deleting),
            n
          )
        case other => other
      }
      copy(drawings = newDrawings)

  /** Remove the drawing at `index` (z-order position); identity when out of range. */
  def removeDrawing(index: Int): Sheet =
    if drawings.isDefinedAt(index) then copy(drawings = drawings.patch(index, Nil, 1))
    else this

  // ===== Conditional formatting (GH-136) =====

  /**
   * Append ONE conditional-formatting block applying `rule` (and `more`) to `range`.
   *
   * Authoring always appends a new block, never merges into existing blocks (Excel accepts and
   * itself writes multiple blocks). Rules with `priority <= 0` ([[CfRule.AutoPriority]]) are
   * assigned `maxExistingPriority + 1, +2, ...` in argument order — above every priority already on
   * the sheet, including those of Preserved rules/blocks. Allocation saturates at `Int.MaxValue` (a
   * wrapped negative priority would be schema-invalid; a collision at the ceiling is
   * Excel-tolerated). Explicit positive priorities pass through unvalidated; colliding priorities
   * are Excel-tolerated but discouraged.
   */
  def conditionalFormat(range: CellRange, rule: CfRule, more: CfRule*): Sheet =
    conditionalFormat(Vector(range), (rule +: more).toVector)

  /** Multi-range variant of [[conditionalFormat]]: one block, several sqref ranges. */
  def conditionalFormat(ranges: Vector[CellRange], rules: Vector[CfRule]): Sheet =
    if ranges.isEmpty || rules.isEmpty then this
    else
      val maxExisting = Sheet.maxCfPriority(conditionalFormats)
      val (stamped, _) = rules.foldLeft((Vector.empty[CfRule], maxExisting)) {
        case ((acc, cur), rule) =>
          CfRule.priorityOf(rule) match
            case Some(p) if p <= 0 =>
              val next = Sheet.nextCfPriority(cur)
              (acc :+ CfRule.withPriority(rule, next), next)
            case _ => (acc :+ rule, cur)
      }
      copy(conditionalFormats = conditionalFormats :+ ConditionalFormat.Rules(ranges, stamped))

  /** Remove the conditional-format block at `index`; identity when out of range. */
  def removeConditionalFormat(index: Int): Sheet =
    if conditionalFormats.isDefinedAt(index) then
      copy(conditionalFormats = conditionalFormats.patch(index, Nil, 1))
    else this

  /** All typed conditional-format blocks in document order (Preserved fragments excluded). */
  def typedConditionalFormats: Vector[ConditionalFormat.Rules] =
    conditionalFormats.collect { case r: ConditionalFormat.Rules => r }

  // ===== Data validation (GH-375) =====

  /**
   * Append ONE data validation applying `dv` to `range` — list dropdowns first:
   * {{{
   * sheet.withDataValidation(ref"B2:B10", DataValidation.listOf("Low", "Med", "High"))
   * sheet.withDataValidation(ref"C2:C10", DataValidation.list("$Z$1:$Z$3", allowBlank = false))
   * }}}
   * Authoring always appends a new entry (Excel accepts several validations per sheet); any ranges
   * already on `dv` are replaced by `range`.
   */
  def withDataValidation(range: CellRange, dv: DataValidation.Rules): Sheet =
    withDataValidation(Vector(range), dv)

  /** Multi-range variant of [[withDataValidation]]: one entry, several sqref ranges. */
  def withDataValidation(ranges: Vector[CellRange], dv: DataValidation.Rules): Sheet =
    if ranges.isEmpty then this
    else copy(dataValidations = dataValidations :+ dv.copy(ranges = ranges))

  /** Remove the data validation at `index` (document order); identity when out of range. */
  def removeDataValidation(index: Int): Sheet =
    if dataValidations.isDefinedAt(index) then
      copy(dataValidations = dataValidations.patch(index, Nil, 1))
    else this

  /** All typed data validations in document order (Preserved entries excluded). */
  def typedDataValidations: Vector[DataValidation.Rules] =
    dataValidations.collect { case r: DataValidation.Rules => r }

  /** Add or update table in sheet */
  def withTable(table: TableSpec): Sheet =
    copy(tables = tables.updated(table.name, table))

  /** Get table by name */
  def getTable(name: String): Option[TableSpec] =
    tables.get(name)

  /** Remove table by name */
  def removeTable(name: String): Sheet =
    copy(tables = tables.removed(name))

  /** Check if table exists */
  def hasTable(name: String): Boolean =
    tables.contains(name)

  /** Get all tables in sheet */
  def allTables: Iterable[TableSpec] =
    tables.values

  /** Set column properties */
  def setColumnProperties(col: Column, props: ColumnProperties): Sheet =
    copy(columnProperties = columnProperties.updated(col, props))

  /** Get column properties */
  def getColumnProperties(col: Column): ColumnProperties =
    columnProperties.getOrElse(col, ColumnProperties())

  /** Set row properties */
  def setRowProperties(row: Row, props: RowProperties): Sheet =
    copy(rowProperties = rowProperties.updated(row, props))

  /** Get row properties */
  def getRowProperties(row: Row): RowProperties =
    rowProperties.getOrElse(row, RowProperties())

  /**
   * Set a column-default style (GH-445): cells in this column with no explicit style render with
   * `style`, and cells typed into the column later inherit it — the `<col style=>` mechanism
   * professional workbooks use for a sheet-wide body font without touching the Normal style.
   *
   * The style is interned into the sheet's [[StyleRegistry]]; other column properties (width,
   * hidden, outline) already set on the column are preserved.
   */
  def withColumnStyle(col: Column, style: CellStyle): Sheet =
    val (newRegistry, styleId) = styleRegistry.register(style)
    copy(
      styleRegistry = newRegistry,
      columnProperties = columnProperties.updated(
        col,
        getColumnProperties(col).copy(styleId = Some(styleId))
      )
    )

  /**
   * Set a row-default style (GH-445): the `<row s= customFormat="1">` mechanism, used for
   * full-width spacer and rule rows. Interns the style; other row properties are preserved.
   */
  def withRowStyle(row: Row, style: CellStyle): Sheet =
    val (newRegistry, styleId) = styleRegistry.register(style)
    copy(
      styleRegistry = newRegistry,
      rowProperties = rowProperties.updated(
        row,
        getRowProperties(row).copy(styleId = Some(styleId))
      )
    )

  /** Get all non-empty cells */
  def nonEmptyCells: Iterable[Cell] =
    cells.values.filter(_.nonEmpty)

  /** Get used range (bounding box of all non-empty cells) */
  def usedRange: Option[CellRange] =
    val nonEmpty = nonEmptyCells
    if nonEmpty.isEmpty then None
    else
      // Single-pass fold to compute min/max for both col and row (75% faster than 4 passes)
      val (minCol, minRow, maxCol, maxRow) = nonEmpty
        .map(_.ref)
        .foldLeft((Int.MaxValue, Int.MaxValue, Int.MinValue, Int.MinValue)) {
          case ((minC, minR, maxC, maxR), ref) =>
            (
              math.min(minC, ref.col.index0),
              math.min(minR, ref.row.index0),
              math.max(maxC, ref.col.index0),
              math.max(maxR, ref.row.index0)
            )
        }
      Some(
        CellRange(
          ARef.from0(minCol, minRow),
          ARef.from0(maxCol, maxRow)
        )
      )

  /** Freeze panes at the given cell (rows above and columns left are frozen). */
  def freezeAt(ref: ARef): Sheet = copy(freezePane = Some(FreezePane.At(ref)))

  /**
   * Freeze panes at `ref` with the scrollable pane scrolled so `scrolledTo` is its top-left visible
   * cell (GH-382) — the OOXML `<pane topLeftCell=".."/>` attribute.
   */
  def freezeAt(ref: ARef, scrolledTo: ARef): Sheet =
    copy(freezePane = Some(FreezePane.At(ref, Some(scrolledTo))))

  /** Remove freeze panes. */
  def unfreeze: Sheet = copy(freezePane = Some(FreezePane.Remove))

  /**
   * Set sheet view settings (gridline visibility, zoom).
   *
   * Example:
   * {{{
   * sheet.withViewSettings(SheetView(showGridLines = false, zoomScale = Some(85)))
   * }}}
   */
  def withViewSettings(view: SheetView): Sheet = copy(viewSettings = Some(view))

  /**
   * Set page setup (print scale/orientation, margins, header/footer, print area, repeat rows).
   *
   * Example:
   * {{{
   * sheet.withPageSetup(PageSetup(
   *   orientation = Some("landscape"),
   *   headerFooter = Some(HeaderFooter(oddFooter = Some("Page &P of &N")))
   * ))
   * }}}
   */
  def withPageSetup(setup: PageSetup): Sheet = copy(pageSetup = Some(setup))

  /**
   * Set the sheet tab color (GH-358). Both RGB and theme colors are legal:
   * {{{
   * sheet.withTabColor(Color.Rgb(0xFF1F4E79))                 // navy
   * sheet.withTabColor(Color.Theme(ThemeSlot.Accent2, 0.25))  // theme + tint
   * }}}
   */
  def withTabColor(color: Color): Sheet = copy(tabColor = Some(color))

  /**
   * Clear the modeled tab color back to the passive default. NOTE: on write, `None` PRESERVES any
   * tabColor already present in the source XML (preserve-if-None, the viewSettings precedent) — it
   * does not actively strip a preserved color.
   */
  def withoutTabColor: Sheet = copy(tabColor = None)

  /** Count of non-empty cells */
  def cellCount: Int = cells.size

  /** Clear all cells */
  def clearCells: Sheet =
    copy(cells = Map.empty)

  /** Clear all merged ranges */
  def clearMerged: Sheet =
    copy(mergedRanges = Set.empty)

  /**
   * Clear styles from cells in range (set styleId to None).
   *
   * This resets cells in the range to the default style without affecting their contents or
   * comments. Cells outside the range are unchanged.
   *
   * @param range
   *   The cell range to clear styles from
   * @return
   *   A new Sheet with styles cleared from cells in the range
   */
  def clearStylesInRange(range: CellRange): Sheet =
    // Use filter + reconstruct pattern for performance:
    // Only modify cells within the range, avoiding reconstruction of unaffected entries
    val (inRange, outsideRange) = cells.partition((ref, _) => range.contains(ref))
    val clearedInRange = inRange.view.mapValues(cell => cell.copy(styleId = None)).toMap
    copy(cells = outsideRange ++ clearedInRange)

  /**
   * Clear comments from cells in range.
   *
   * Removes all comments from cells within the specified range. Cell contents and styles are not
   * affected.
   *
   * @param range
   *   The cell range to clear comments from
   * @return
   *   A new Sheet with comments removed from cells in the range
   */
  def clearCommentsInRange(range: CellRange): Sheet =
    copy(comments = comments.filterNot((ref, _) => range.contains(ref)))

object Sheet:

  /**
   * Span algebra shared by every range-shaped structure a structural edit moves: merged ranges,
   * cf/dv envelopes, print areas, repeat rows, tables, autoFilter, and chart data references.
   *
   * Maps an inclusive 0-based `[s, e]` span through an insert (`deleting = false`) or delete of
   * `count` indices at `at` on one axis, clamped to the axis maximum (GH-428): an insert pins the
   * span END at `axisMax` (Excel's behavior for full-height/width ranges — an unclamped shift emits
   * rows past 1048576 / columns past XFD and Excel refuses the file); a span whose START passes the
   * edge drops. A delete drops the overlap; `None` = the span vanished. Long intermediate math so a
   * pathological `count` cannot overflow.
   */
  private[xl] def shiftSpan(
    s: Int,
    e: Int,
    at: Int,
    count: Int,
    deleting: Boolean,
    axisMax: Int
  ): Option[(Int, Int)] =
    // A range that covers the complete edited axis models a whole-column/whole-row reference
    // (A:A / 1:1). Structural edits cannot make that reference partial: it must continue to cover
    // the complete axis. Besides preserving the reference's semantics, this keeps print areas and
    // preserved sqref tokens in their full-axis shape at the two cases the generic endpoint
    // algebra cannot distinguish: inserts at index 0 and any deletion.
    if s == 0 && e == axisMax then Some((0, axisMax))
    else if !deleting then
      val ns = if s >= at then s.toLong + count else s.toLong
      val ne = if e >= at then e.toLong + count else e.toLong
      if ns > axisMax then None
      else Some((ns.toInt, math.min(ne, axisMax.toLong).toInt))
    else
      val ns = if s < at then s else if s >= at + count then s - count else at
      val ne = if e < at then e else if e >= at + count then e - count else at - 1
      if ns > ne then None else Some((ns, ne))

  /**
   * Re-derive a table's column vector for a COLUMN-axis structural edit (GH-429). `newRange` is the
   * spec's range already mapped through [[shiftSpan]] (so `TableSpec.isValid` holds by
   * construction). Returns the updated spec plus the (absolute 0-based column index, name) of every
   * SYNTHESIZED column so the caller can stamp header cells; `None` drops the table.
   *
   *   - edit outside the span / pure translation: columns untouched
   *   - delete: drop the TableColumns at absolute indices in `[at, at+count)`; drop the table when
   *     none survive
   *   - insert strictly inside: splice `count` fresh `TableColumn(maxId+i, "ColumnN")` at the edit
   *     offset, names case-insensitively unique against the existing ones (Excel's scheme)
   */
  private[xl] def retabulateColumns(
    spec: TableSpec,
    newRange: CellRange,
    at: Int,
    count: Int,
    deleting: Boolean
  ): Option[(TableSpec, Vector[(Int, String)])] =
    val oldStart = spec.range.start.col.index0
    val oldEnd = spec.range.end.col.index0
    if deleting then
      val survivors = spec.columns.zipWithIndex.collect {
        case (c, i) if oldStart + i < at || oldStart + i >= at + count => c
      }
      Option.when(survivors.nonEmpty)(
        (spec.copy(range = newRange, columns = survivors), Vector.empty)
      )
    else if at <= oldStart || at > oldEnd then
      // left of the span (pure translation) or right of it (no-op): columns untouched
      Some((spec.copy(range = newRange), Vector.empty))
    else
      def freshName(taken: Set[String]): String =
        Iterator
          .from(1)
          .map(k => s"Column$k")
          .find(nm => !taken.contains(nm.toLowerCase))
          .getOrElse(s"Column${taken.size + 1}") // unreachable: the iterator is unbounded
      val maxId = spec.columns.map(_.id).maxOption.getOrElse(0L)
      val (fresh, _) = (0 until count).foldLeft(
        (Vector.empty[TableColumn], spec.columns.map(_.name.toLowerCase).toSet)
      ) { case ((acc, taken), i) =>
        val n = freshName(taken)
        (acc :+ TableColumn(maxId + i + 1, n), taken + n.toLowerCase)
      }
      val offset = at - newRange.start.col.index0
      val spliced = spec.columns.take(offset) ++ fresh ++ spec.columns.drop(offset)
      val stamps = fresh.zipWithIndex.map((c, i) => (at + i, c.name))
      Some((spec.copy(range = newRange, columns = spliced), stamps))

  /**
   * Highest priority present across the sheet's conditional formats (0 when none): typed rules'
   * priorities ∪ parsed `CfRule.Preserved` priorities ∪ priorities text-scanned from
   * `ConditionalFormat.Preserved` payloads (GH-136). Auto-priority allocation appends above this.
   */
  private[xl] def maxCfPriority(cfs: Vector[ConditionalFormat]): Int =
    cfs.foldLeft(0) {
      case (acc, ConditionalFormat.Rules(_, rules, _)) =>
        rules.foldLeft(acc)((a, r) => math.max(a, CfRule.priorityOf(r).getOrElse(0)))
      case (acc, ConditionalFormat.Preserved(xml)) =>
        ConditionalFormat.scanPriorities(xml).foldLeft(acc)(math.max)
    }

  /**
   * Saturating successor for auto-priority allocation (shared with the OOXML emitter's safety net):
   * `current + 1`, capped at `Int.MaxValue` so an existing ceiling priority can never wrap to a
   * schema-invalid negative.
   */
  private[xl] def nextCfPriority(current: Int): Int =
    if current == Int.MaxValue then current else current + 1

  /**
   * Shift a chart's data references that point at sheet `edited` (matched case-insensitively, the
   * FormulaShifter convention) through one structural axis edit (GH-222). Same clamp algebra as
   * merged ranges. Deletion semantics (total — no `#REF!` state in the typed model):
   *   - partial overlap → clamp to the surviving span
   *   - values fully deleted → DROP the series
   *   - categories fully deleted → `categories = None` (Excel falls back to 1..N)
   *   - name cell deleted → `name = None`
   */
  private[xl] def shiftChartReferences(
    chart: Chart,
    edited: String,
    rowAxis: Boolean,
    at: Int,
    count: Int,
    deleting: Boolean
  ): Chart =
    def axisStart(r: CellRange): Int =
      if rowAxis then r.start.row.index0 else r.start.col.index0
    def axisEnd(r: CellRange): Int =
      if rowAxis then r.end.row.index0 else r.end.col.index0
    def rebuild(ref: ARef, newIdx: Int): ARef =
      if rowAxis then ARef.from0(ref.col.index0, newIdx) else ARef.from0(newIdx, ref.row.index0)
    val axisMax = if rowAxis then Row.MaxIndex0 else Column.MaxIndex0
    // The shared shiftSpan clamp algebra: shift both endpoints; collapse-to-empty = None.
    def shiftRange(range: CellRange): Option[CellRange] =
      shiftSpan(axisStart(range), axisEnd(range), at, count, deleting, axisMax)
        .map((ns, ne) => CellRange(rebuild(range.start, ns), rebuild(range.end, ne)))
    def shiftCell(ref: ARef): Option[ARef] =
      val i = if rowAxis then ref.row.index0 else ref.col.index0
      shiftSpan(i, i, at, count, deleting, axisMax).map((ni, _) => rebuild(ref, ni))
    def matches(sheet: SheetName): Boolean = sheet.value.equalsIgnoreCase(edited)
    val shiftedSeries = chart.series.flatMap { series =>
      val newValues =
        if matches(series.values.sheet) then
          shiftRange(series.values.range).map(r => series.values.copy(range = r))
        else Some(series.values)
      newValues.map { values =>
        val newCats = series.categories.flatMap { cats =>
          if matches(cats.sheet) then shiftRange(cats.range).map(r => cats.copy(range = r))
          else Some(cats)
        }
        val newName = series.name.flatMap {
          case SeriesName.FromCell(sheet, ref) if matches(sheet) =>
            shiftCell(ref).map(SeriesName.FromCell(sheet, _))
          case other => Some(other)
        }
        // copy, not reconstruct: styling (fill, pointFills) survives the structural shift
        series.copy(values = values, categories = newCats, name = newName)
      }
    }
    chart.copy(series = shiftedSeries)

  /**
   * Create empty sheet with name.
   *
   * When called with a string literal, the name is validated at compile time and returns `Sheet`
   * directly. When called with a runtime expression, validation occurs at runtime and returns
   * `XLResult[Sheet]`.
   *
   * Validation rules (Excel sheet name constraints):
   *   - Cannot be empty
   *   - Maximum 31 characters
   *   - Cannot contain: : \ / ? * [ ]
   *
   * Examples:
   *   - `Sheet("Sales")` → `Sheet` (compile-time validated)
   *   - `Sheet(userInput)` → `XLResult[Sheet]` (runtime validated)
   *
   * The literal-vs-dynamic specialization is invisible at the call site — nothing in `Sheet(nm)`
   * hints that the return type changed. For names computed at runtime, prefer [[named]], which
   * spells the `XLResult` in its signature (GH-420).
   */
  @annotation.targetName("applyStringLiteral")
  transparent inline def apply(inline name: String): Sheet | XLResult[Sheet] =
    ${ com.tjclp.xl.macros.SheetLiteral.sheetImpl('name) }

  /**
   * Create an empty sheet from a name computed at runtime — the documented dynamic-name factory
   * (GH-420).
   *
   * Semantically identical to the dynamic branch of the union-typed `apply` (same [[SheetName]]
   * validation, same [[com.tjclp.xl.error.XLError.InvalidSheetName]] error), but non-inline with
   * the `XLResult` spelled in the signature, so the `.map`/`.unsafe` step is expected rather than a
   * surprise:
   *
   * {{{
   * val nm: String = config.sheetName
   * val sheet: XLResult[Sheet] = Sheet.named(nm).map(_.put(ref"A1", 1))
   * }}}
   */
  def named(name: String): XLResult[Sheet] =
    SheetName(name) match
      case Right(validName) => Right(Sheet(validName))
      case Left(err) => Left(XLError.InvalidSheetName(name, err))

  /** Create empty sheet with validated name */
  @annotation.targetName("applySheetName")
  def apply(name: SheetName): Sheet =
    Sheet(
      name,
      Map.empty,
      Set.empty,
      Map.empty,
      Map.empty,
      None,
      None,
      StyleRegistry.default,
      Map.empty,
      Map.empty
    )
