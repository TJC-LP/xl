package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{ARef, CellRange, Column, RefType, Row, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue, Comment}
import com.tjclp.xl.codec.CellCodec
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
  tables: Map[String, TableSpec] = Map.empty
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

  /** Put cell at reference */
  def put(cell: Cell): Sheet =
    copy(cells = cells.updated(cell.ref, cell))

  /** Put value at reference (preserves existing cell metadata) */
  def put(ref: ARef, value: CellValue): Sheet =
    val updatedCell = cells.get(ref) match
      case Some(existing) => existing.withValue(value) // Preserve styleId, comment, hyperlink
      case None => Cell(ref, value) // New cell with no metadata
    put(updatedCell)

  /**
   * Batch put with mixed value types and automatic style inference.
   *
   * Accepts (ARef, Any) pairs and uses runtime pattern matching to resolve the appropriate codec
   * for each value. Auto-infers styles based on value types (dates get date format, decimals get
   * number format, etc.). Formatted literals (money"", date"", percent"") preserve their NumFmt.
   *
   * Supported types: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime,
   * RichText, Formatted. Unsupported types return Left with error.
   *
   * Example:
   * {{{
   * sheet.put(
   *   ref"A1" -> "Revenue",
   *   ref"B1" -> LocalDate.of(2025, 11, 10),
   *   ref"C1" -> money"$$1,234.56"
   * ) match
   *   case Right(updated) => updated
   *   case Left(err) => handleError(err)
   * }}}
   *
   * For demos/REPLs, you can use .unsafe (requires explicit import):
   * {{{
   * import com.tjclp.xl.unsafe.*
   * sheet.put(ref"A1" -> "Hello").unsafe
   * }}}
   *
   * @param updates
   *   Varargs of (ARef, Any) pairs
   * @return
   *   Either error (if unsupported type) or updated sheet
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def put(updates: (ARef, Any)*): XLResult[Sheet] =
    import com.tjclp.xl.codec.{CellCodec, given}
    import java.time.{LocalDate, LocalDateTime}

    boundary:
      // NOTE: Local mutation for performance - buffers are private to this method
      // and never escape. The function remains pure (referentially transparent) because:
      // 1. All mutations are confined to this scope
      // 2. No shared mutable state is accessed
      // 3. Output depends only on inputs (deterministic)
      // This is a common FP optimization pattern for bulk operations (similar to Scala stdlib).

      // Single-pass: build cells and collect styles simultaneously
      val cells = scala.collection.mutable.ArrayBuffer[Cell]()
      val cellsWithStyles = scala.collection.mutable.ArrayBuffer[(ARef, CellStyle)]()
      var registry = styleRegistry

      // Helper to process a typed value (DRY)
      def processValue[A: CellCodec](ref: ARef, value: A): Unit =
        val (cellValue, styleOpt) = CellCodec[A].write(value)
        cells += Cell(ref, cellValue)
        styleOpt.foreach { style =>
          val (newRegistry, _) = registry.register(style)
          registry = newRegistry
          cellsWithStyles += ((ref, style))
        }

      // Pattern match on runtime type and delegate to helper
      updates.foreach { (ref, value) =>
        value match
          // Handle Formatted values (money"", date"", etc.) - preserve NumFmt metadata
          case formatted: com.tjclp.xl.formatted.Formatted =>
            cells += Cell(ref, formatted.value)
            val style = CellStyle.default.withNumFmt(formatted.numFmt)
            val (newRegistry, _) = registry.register(style)
            registry = newRegistry
            cellsWithStyles += ((ref, style))

          case v: String => processValue(ref, v)
          case v: Int => processValue(ref, v)
          case v: Long => processValue(ref, v)
          case v: Double => processValue(ref, v)
          case v: BigDecimal => processValue(ref, v)
          case v: Boolean => processValue(ref, v)
          case v: LocalDate => processValue(ref, v)
          case v: LocalDateTime => processValue(ref, v)
          case v: com.tjclp.xl.richtext.RichText => processValue(ref, v)
          case unsupported =>
            break(Left(XLError.UnsupportedType(ref.toA1, unsupported.getClass.getName)))
      }

      Right(applyBulkCells(cells, cellsWithStyles, registry))

  /**
   * Type-safe variant of [[put]] that requires a single `CellCodec` for all values.
   *
   * This avoids runtime type inspection and lets the compiler inline the codec implementation,
   * ensuring zero-overhead writes when the value type is known statically.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def putTyped[A](updates: (ARef, A)*)(using CellCodec[A]): Sheet =
    val cells = scala.collection.mutable.ArrayBuffer[Cell]()
    val cellsWithStyles = scala.collection.mutable.ArrayBuffer[(ARef, CellStyle)]()
    var registry = styleRegistry
    val codec = summon[CellCodec[A]]

    updates.foreach { (ref, value) =>
      val (cellValue, styleOpt) = codec.write(value)
      cells += Cell(ref, cellValue)
      styleOpt.foreach { style =>
        val (newRegistry, _) = registry.register(style)
        registry = newRegistry
        cellsWithStyles += ((ref, style))
      }
    }

    applyBulkCells(cells, cellsWithStyles, registry)

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
   * Patches enable declarative composition of updates (Put, SetStyle, Merge, etc.). Returns Either
   * for operations that can fail (e.g., merge overlaps, invalid ranges).
   *
   * Example:
   * {{{
   * val patch = (ref"A1" := "Title") ++ range"A1:C1".merge
   * sheet.put(patch) match
   *   case Right(updated) => updated
   *   case Left(err) => handleError(err)
   * }}}
   *
   * @param patch
   *   The patch to apply
   * @return
   *   Either an updated sheet or an error
   */
  def put(patch: com.tjclp.xl.patch.Patch): XLResult[Sheet] =
    com.tjclp.xl.patch.Patch.applyPatch(this, patch)

  /** Remove cell at reference */
  def remove(ref: ARef): Sheet =
    copy(cells = cells.removed(ref))

  /** Remove all cells in range */
  def removeRange(range: CellRange): Sheet =
    val toRemove = range.cells.toSet
    copy(cells = cells.filterNot((ref, _) => toRemove.contains(ref)))

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

object Sheet:
  /** Create empty sheet with name */
  def apply(name: String): XLResult[Sheet] =
    SheetName(name).left
      .map(err => XLError.InvalidSheetName(name, err))
      .map(sn => Sheet(sn))

  /** Create empty sheet with validated name */
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
