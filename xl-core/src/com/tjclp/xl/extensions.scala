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
     * When called with a string literal, the reference is validated at compile time and returns
     * `Sheet` directly. When called with a runtime string, validation is deferred and returns
     * `XLResult[Sheet]`.
     *
     * @param ref
     *   Cell ("A1") or range ("A1:B10")
     * @param cellStyle
     *   CellStyle to apply
     * @return
     *   `Sheet` for literal refs (compile-time validated), `XLResult[Sheet]` for runtime refs
     */
    @annotation.targetName("styleSheet")
    transparent inline def style(
      inline ref: String,
      cellStyle: CellStyle
    ): Sheet | XLResult[Sheet] =
      ${ com.tjclp.xl.macros.PutLiteral.styleImpl('{ sheet }, 'ref, 'cellStyle) }

    /**
     * Apply style to cell (compile-time validated ref).
     *
     * Universal style method that works with typed ARef from ref"A1" macro. This operation is
     * infallible since the reference is already validated.
     *
     * @param ref
     *   Cell reference (ARef from ref"A1" macro)
     * @param cellStyle
     *   CellStyle to apply
     * @return
     *   Updated Sheet
     */
    @annotation.targetName("styleSheetARef")
    def style(ref: com.tjclp.xl.addressing.ARef, cellStyle: CellStyle): Sheet =
      sheet.withCellStyle(ref, cellStyle)

    /**
     * Apply style to range (compile-time validated ref).
     *
     * Universal style method that works with typed CellRange from ref"A1:B10" macro. This operation
     * is infallible since the range is already validated.
     *
     * @param range
     *   Cell range (CellRange from ref"A1:B10" macro)
     * @param cellStyle
     *   CellStyle to apply
     * @return
     *   Updated Sheet
     */
    @annotation.targetName("styleSheetRange")
    def style(range: com.tjclp.xl.addressing.CellRange, cellStyle: CellStyle): Sheet =
      sheet.withRangeStyle(range, cellStyle)

  // ========== Sheet Extensions: Lookup Operations (Safe) ==========

  extension (sheet: Sheet)
    /**
     * Get cell at reference (safe lookup).
     *
     * When called with a string literal, the reference is validated at compile time. Invalid
     * literals like "INVALID" fail to compile. When called with a runtime string, validation is
     * deferred and returns None for invalid refs.
     *
     * @param cellRef
     *   Cell reference like "A1"
     * @return
     *   Some(cell) if valid ref and exists, None otherwise
     */
    @annotation.targetName("cellString")
    transparent inline def cell(inline cellRef: String): Option[Cell] =
      ${ com.tjclp.xl.macros.CellLookupMacros.cellImpl('{ sheet }, 'cellRef) }

    /**
     * Get cells in range (safe lookup).
     *
     * When called with a string literal, the range is validated at compile time. Invalid literals
     * fail to compile. When called with a runtime string, validation is deferred and returns empty
     * list for invalid ranges.
     *
     * @param rangeRef
     *   Range like "A1:B10"
     * @return
     *   List of cells (only existing cells)
     */
    @annotation.targetName("rangeString")
    transparent inline def range(inline rangeRef: String): List[Cell] =
      ${ com.tjclp.xl.macros.CellLookupMacros.rangeImpl('{ sheet }, 'rangeRef) }

    /**
     * Get cell(s) at reference (auto-detects cell vs range).
     *
     * When called with a string literal, the reference is validated at compile time and
     * auto-detected as cell or range. Invalid literals fail to compile.
     *
     * @param ref
     *   Cell ("A1") or range ("A1:B10")
     * @return
     *   List of cells (empty if invalid ref or no cells exist)
     */
    @annotation.targetName("getString")
    transparent inline def get(inline ref: String): List[Cell] =
      ${ com.tjclp.xl.macros.CellLookupMacros.getImpl('{ sheet }, 'ref) }

  // ========== Sheet Extensions: Merge Operations ==========

  extension (sheet: Sheet)
    /**
     * Merge cells in range.
     *
     * When called with a string literal, the reference is validated at compile time and returns
     * `Sheet` directly. When called with a runtime string, validation is deferred and returns
     * `XLResult[Sheet]`.
     *
     * @param rangeRef
     *   Range like "A1:B1"
     * @return
     *   `Sheet` for literal refs (compile-time validated), `XLResult[Sheet]` for runtime refs
     */
    @annotation.targetName("mergeSheet")
    transparent inline def merge(inline rangeRef: String): Sheet | XLResult[Sheet] =
      ${ com.tjclp.xl.macros.PutLiteral.mergeImpl('{ sheet }, 'rangeRef) }

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
      result.map(_.style(ref, cellStyle))

    /**
     * Apply style (chainable, range).
     */
    @annotation.targetName("styleSheetRangeChainable")
    def style(range: com.tjclp.xl.addressing.CellRange, cellStyle: CellStyle): XLResult[Sheet] =
      result.map(_.style(range, cellStyle))

    /** Merge range (chainable). */
    def merge(rangeRef: String): XLResult[Sheet] =
      result.flatMap(_.merge(rangeRef))

    /** Apply patch (chainable). */
    @annotation.targetName("putPatchChainable")
    def put(patch: com.tjclp.xl.patch.Patch): XLResult[Sheet] =
      result.map(_.put(patch))

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
      result.map(_.put(sheet))

    /**
     * Update sheet by name (chainable).
     *
     * When called with a string literal, the name format is validated at compile time. Invalid
     * literals fail to compile.
     */
    @annotation.targetName("updateSheetChainable")
    inline def update(inline name: String, f: Sheet => Sheet): XLResult[Workbook] =
      result.flatMap(_.update(name, f))

    /**
     * Remove sheet by name (chainable).
     *
     * When called with a string literal, the name format is validated at compile time. Invalid
     * literals fail to compile.
     */
    @annotation.targetName("removeSheetChainable")
    inline def remove(inline name: String): XLResult[Workbook] =
      result.flatMap(_.remove(name))

  // ========== Workbook Extensions: String-Based Lookups ==========

  extension (workbook: Workbook)
    /**
     * Get sheet by name (safe lookup).
     *
     * When called with a string literal, the name format is validated at compile time. Invalid
     * literals like "Invalid:Name" fail to compile. Returns None if sheet not found.
     *
     * @param name
     *   Sheet name
     * @return
     *   Some(sheet) if found, None otherwise
     */
    @annotation.targetName("getSheetByName")
    transparent inline def get(inline name: String): Option[Sheet] =
      ${ com.tjclp.xl.macros.WorkbookMacros.getImpl('{ workbook }, 'name) }
