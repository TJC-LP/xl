package com.tjclp.xl

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue, Comment}
import com.tjclp.xl.codec.CellWriter
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.workbooks.Workbook

/**
 * String-based extensions for ergonomic Easy Mode API.
 *
 * Extensions return [[XLResult]] for chainable error handling. Use `.unsafe` once at the end of
 * chains to unwrap.
 *
 * '''Example: Chainable operations'''
 * {{{
 * import com.tjclp.xl.*
 *
 * val sheet = Sheet("Sales")
 *   .put("A1", "Product")
 *   .put("B1", "Revenue", headerStyle)
 *   .style("A1:B1", boldStyle)
 *   .unsafe  // Single boundary point
 * }}}
 *
 * '''Styling approaches:'''
 *   - `.style(range, style)` - Template pattern (format before data)
 *   - `.put(ref, value, style)` - Inline styling
 *   - RichText - Intra-cell formatting
 *
 * @since 0.3.0
 */
@SuppressWarnings(Array("org.wartremover.warts.Throw"))
object extensions:

  // Re-export CellWriter given instances for type class-based put() methods
  export com.tjclp.xl.codec.CellCodec.given

  // Helper to convert Either[String, A] to XLResult[A]
  private def toXLResult[A](
    either: Either[String, A],
    ref: String,
    errorType: String
  ): XLResult[A] =
    either.left.map(msg => XLError.InvalidCellRef(ref, s"$errorType: $msg"))

  // ========== Sheet Extensions: Style Operations ==========

  extension (sheet: Sheet)
    /**
     * Apply style to cell or range (template pattern).
     *
     * @param ref
     *   Cell ("A1") or range ("A1:B10")
     * @param cellStyle
     *   CellStyle to apply
     * @return
     *   XLResult[Sheet] for chaining
     */
    @annotation.targetName("styleSheet")
    def style(ref: String, cellStyle: CellStyle): XLResult[Sheet] =
      if ref.contains(":") then
        toXLResult(CellRange.parse(ref), ref, "Invalid range")
          .map(range => sheet.withRangeStyle(range, cellStyle))
      else
        toXLResult(ARef.parse(ref), ref, "Invalid cell reference")
          .map(aref => sheet.withCellStyle(aref, cellStyle))

    /**
     * Apply style to cell (compile-time validated ref).
     *
     * Universal style method that works with typed ARef from ref"A1" macro.
     *
     * @param ref
     *   Cell reference (ARef from ref"A1" macro)
     * @param cellStyle
     *   CellStyle to apply
     * @return
     *   XLResult[Sheet] for chaining
     */
    @annotation.targetName("styleSheetARef")
    def style(ref: com.tjclp.xl.addressing.ARef, cellStyle: CellStyle): XLResult[Sheet] =
      Right(sheet.withCellStyle(ref, cellStyle))

    /**
     * Apply style to range (compile-time validated ref).
     *
     * Universal style method that works with typed CellRange from ref"A1:B10" macro.
     *
     * @param range
     *   Cell range (CellRange from ref"A1:B10" macro)
     * @param cellStyle
     *   CellStyle to apply
     * @return
     *   XLResult[Sheet] for chaining
     */
    @annotation.targetName("styleSheetRange")
    def style(range: com.tjclp.xl.addressing.CellRange, cellStyle: CellStyle): XLResult[Sheet] =
      Right(sheet.withRangeStyle(range, cellStyle))

  // ========== Sheet Extensions: Lookup Operations (Safe) ==========

  extension (sheet: Sheet)
    /**
     * Get cell at reference (safe lookup).
     *
     * Returns None for invalid references or missing cells. No exceptions thrown.
     *
     * @param cellRef
     *   Cell reference like "A1"
     * @return
     *   Some(cell) if valid ref and exists, None otherwise
     */
    def cell(cellRef: String): Option[Cell] =
      ARef.parse(cellRef).toOption.flatMap(ref => sheet.cells.get(ref))

    /**
     * Get cells in range (safe lookup).
     *
     * Returns empty list for invalid ranges. Only includes existing cells.
     *
     * @param rangeRef
     *   Range like "A1:B10"
     * @return
     *   List of cells (only existing cells)
     */
    def range(rangeRef: String): List[Cell] =
      CellRange
        .parse(rangeRef)
        .toOption
        .map(r => r.cells.flatMap(sheet.cells.get).toList)
        .getOrElse(List.empty)

    /**
     * Get cell(s) at reference (auto-detects cell vs range).
     *
     * Convenience method that handles both single cells and ranges uniformly. Returns List[Cell]
     * for consistent handling.
     *
     * @param ref
     *   Cell ("A1") or range ("A1:B10")
     * @return
     *   List of cells (empty if invalid ref or no cells exist)
     */
    def get(ref: String): List[Cell] =
      if ref.contains(":") then range(ref) // Range → List[Cell]
      else cell(ref).toList // Cell → List[0 or 1]

  // ========== Sheet Extensions: Merge Operations ==========

  extension (sheet: Sheet)
    /**
     * Merge cells in range.
     *
     * @param rangeRef
     *   Range like "A1:B1"
     * @return
     *   XLResult[Sheet] for chaining
     */
    def merge(rangeRef: String): XLResult[Sheet] =
      toXLResult(CellRange.parse(rangeRef), rangeRef, "Invalid range")
        .map(range => sheet.merge(range))

  // ========== XLResult[Sheet] Extensions: Chainable Operations ==========

  extension (result: XLResult[Sheet])
    /**
     * Put typed value (chainable).
     *
     * Chains after previous XLResult[Sheet] operations. All types supported.
     */
    @annotation.targetName("putGenericChainable")
    def put[A: CellWriter](cellRef: String, value: A): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value))

    /**
     * Put typed value with style (chainable).
     *
     * Chains after previous XLResult[Sheet] operations with inline styling.
     */
    @annotation.targetName("putGenericStyledChainable")
    def put[A: CellWriter](cellRef: String, value: A, cellStyle: CellStyle): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value, cellStyle))

    /**
     * Put value at ARef (chainable).
     *
     * Universal put that accepts any supported type at compile-time validated ref.
     */
    @annotation.targetName("putARefChainable")
    def put[A: CellWriter](ref: com.tjclp.xl.addressing.ARef, value: A): XLResult[Sheet] =
      result.map(_.put(ref, value))

    /**
     * Put value at ARef with style (chainable).
     */
    @annotation.targetName("putARefStyledChainable")
    def put[A: CellWriter](
      ref: com.tjclp.xl.addressing.ARef,
      value: A,
      cellStyle: CellStyle
    ): XLResult[Sheet] =
      result.map(_.put(ref, value, cellStyle))

    /**
     * Apply style (chainable, string ref).
     */
    @annotation.targetName("styleSheetChainable")
    def style(ref: String, cellStyle: CellStyle): XLResult[Sheet] =
      result.flatMap(_.style(ref, cellStyle))

    /**
     * Apply style (chainable, typed ref).
     */
    @annotation.targetName("styleSheetARefChainable")
    def style(ref: com.tjclp.xl.addressing.ARef, cellStyle: CellStyle): XLResult[Sheet] =
      result.flatMap(_.style(ref, cellStyle))

    /**
     * Apply style (chainable, range).
     */
    @annotation.targetName("styleSheetRangeChainable")
    def style(range: com.tjclp.xl.addressing.CellRange, cellStyle: CellStyle): XLResult[Sheet] =
      result.flatMap(_.style(range, cellStyle))

    /** Merge range (chainable). */
    def merge(rangeRef: String): XLResult[Sheet] =
      result.flatMap(_.merge(rangeRef))

    /** Apply patch (chainable). */
    @annotation.targetName("putPatchChainable")
    def put(patch: com.tjclp.xl.patch.Patch): XLResult[Sheet] =
      result.flatMap(_.put(patch))

    /**
     * Batch put (chainable).
     *
     * Chain batch puts after XLResult[Sheet] operations.
     */
    @annotation.targetName("putBatchChainable")
    def put[A: CellWriter](updates: (com.tjclp.xl.addressing.ARef, A)*): XLResult[Sheet] =
      result.map(_.put(updates*))

    /** Add comment to cell (chainable). */
    @annotation.targetName("commentChainable")
    def comment(
      ref: com.tjclp.xl.addressing.ARef,
      cmt: com.tjclp.xl.cells.Comment
    ): XLResult[Sheet] =
      result.map(_.comment(ref, cmt))

  // ========== XLResult[Workbook] Extensions: Chainable Operations ==========

  extension (result: XLResult[Workbook])
    /** Put sheet (add-or-replace, chainable). */
    @annotation.targetName("putSheetChainable")
    def put(sheet: Sheet): XLResult[Workbook] =
      result.flatMap(_.put(sheet))

    /** Update sheet by name (chainable). */
    @annotation.targetName("updateSheetChainable")
    def update(name: String, f: Sheet => Sheet): XLResult[Workbook] =
      result.flatMap(_.update(name, f))

    /** Remove sheet by name (chainable). */
    @annotation.targetName("removeSheetChainable")
    def remove(name: String): XLResult[Workbook] =
      result.flatMap(wb =>
        toXLResult(SheetName(name), name, "Invalid sheet name").flatMap(wb.remove)
      )

  // ========== Workbook Extensions: String-Based Lookups ==========

  extension (workbook: Workbook)
    /**
     * Get sheet by name (safe lookup).
     *
     * Returns None if sheet not found. No exceptions thrown.
     *
     * @param name
     *   Sheet name
     * @return
     *   Some(sheet) if found, None otherwise
     */
    def get(name: String): Option[Sheet] =
      workbook(name).toOption
