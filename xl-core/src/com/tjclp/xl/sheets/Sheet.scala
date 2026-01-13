package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{ARef, CellRange, Column, RefType, Row, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue, Comment}
import com.tjclp.xl.codec.{CellCodec, CellWritable, CellWriter}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.styles.{CellStyle, StyleRegistry}
import com.tjclp.xl.tables.TableSpec

import scala.collection.immutable.{Map, Set}
import scala.util.boundary, boundary.break

/**
 * A worksheet containing cells, merged ranges, and properties.
 *
 * Immutable design: all operations return new Sheet instances. Uses persistent data structures for
 * efficient updates.
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
  pageSetup: Option[PageSetup] = None
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

  /** Put cell at reference (always succeeds - Cell is pre-validated) */
  def put(cell: Cell): Sheet =
    copy(cells = cells.updated(cell.ref, cell))

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
  private def putSingle[A: CellWriter](ref: ARef, value: A): Sheet =
    import com.tjclp.xl.codec.given
    val (cellValue, styleOpt) = CellWriter[A].write(value)
    val existingCell = cells.get(ref)
    val updatedCell = existingCell match
      case Some(existing) => existing.withValue(cellValue)
      case None => Cell(ref, cellValue)
    val sheetWithCell = copy(cells = cells.updated(ref, updatedCell))
    styleOpt match
      case Some(codecStyle) =>
        val mergedStyle = existingCell.flatMap(_.styleId).flatMap(styleRegistry.get) match
          case Some(existingStyle) => mergeStyles(existingStyle, codecStyle)
          case None => codecStyle
        val (newRegistry, _) = styleRegistry.register(mergedStyle)
        import com.tjclp.xl.sheets.styleSyntax.withCellStyle
        sheetWithCell.copy(styleRegistry = newRegistry).withCellStyle(ref, mergedStyle)
      case None =>
        sheetWithCell

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

    // Single-pass: build cells and collect styles simultaneously
    val builtCells = scala.collection.mutable.ArrayBuffer[Cell]()
    val cellsWithStyles = scala.collection.mutable.ArrayBuffer[(ARef, CellStyle)]()
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
        val (newRegistry, _) = registry.register(mergedStyle)
        registry = newRegistry
        cellsWithStyles += ((ref, mergedStyle))
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
    val cellsWithStyles = scala.collection.mutable.ArrayBuffer[(ARef, CellStyle)]()
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
        val (newRegistry, _) = registry.register(mergedStyle)
        registry = newRegistry
        cellsWithStyles += ((ref, mergedStyle))
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

  private def applyBulkCells(
    builtCells: Iterable[Cell],
    styled: Iterable[(ARef, CellStyle)],
    newRegistry: StyleRegistry
  ): Sheet =
    val withCells = copy(
      styleRegistry = newRegistry,
      cells = this.cells ++ builtCells.iterator.map(cell => cell.ref -> cell)
    )
    import com.tjclp.xl.sheets.styleSyntax.withCellStyle
    styled.foldLeft(withCells) { case (s, (ref, style)) => s.withCellStyle(ref, style) }

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
   */
  @annotation.targetName("applyStringLiteral")
  transparent inline def apply(inline name: String): Sheet | XLResult[Sheet] =
    ${ com.tjclp.xl.macros.SheetLiteral.sheetImpl('name) }

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
