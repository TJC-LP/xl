package com.tjclp.xl

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cell.{Cell, CellValue}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.sheet.Sheet
import com.tjclp.xl.style.CellStyle
import com.tjclp.xl.unsafe.*
import com.tjclp.xl.workbook.Workbook

import java.time.{LocalDate, LocalDateTime}

/**
 * String-based extensions for ergonomic Easy Mode API.
 *
 * These extensions accept string cell references and throw [[com.tjclp.xl.error.XLException]] on
 * parse errors, enabling a more script-friendly API without explicit Either handling.
 *
 * '''Philosophy:''' `.unsafe` is an implementation detail. Users never type `.unsafe` in their code -
 * it's hidden within these extension methods.
 *
 * '''Example: Easy Mode usage'''
 * {{{
 * import com.tjclp.xl.*  // Includes extensions automatically
 *
 * val sheet = Sheet("Sales")
 *   .put("A1", "Product")               // String ref, throws on invalid ref
 *   .put("B1", "Revenue", headerStyle)  // With inline styling
 *   .applyStyle("A1:B1", boldStyle)     // Template-first: style then data
 * }}}
 *
 * '''Two styling approaches:'''
 *   - `.applyStyle(range, style)` - Format structure before data (template pattern)
 *   - `.put(ref, value, style)` - Inline styling for one-offs
 *   - RichText - Intra-cell formatting (multiple styles in one cell)
 *
 * @since 0.3.0
 */
@SuppressWarnings(Array("org.wartremover.warts.Throw"))
object extensions:

  // Helper to convert Either[String, A] to XLResult[A]
  private def toXLResult[A](
    either: Either[String, A],
    ref: String,
    errorType: String
  ): XLResult[A] =
    either.left.map(msg => XLError.InvalidCellRef(ref, s"$errorType: $msg"))

  // ========== Sheet Extensions: Data Operations ==========

  extension (sheet: Sheet)
    /**
     * Put a String value at cell reference (string).
     *
     * @param cellRef
     *   Cell reference like "A1", "B2", etc.
     * @param value
     *   String value to write
     * @return
     *   Updated sheet
     * @throws com.tjclp.xl.error.XLException
     *   if cellRef is invalid
     */
    def put(cellRef: String, value: String): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      sheet.put(ref, CellValue.Text(value))

    /** Put an Int value at cell reference (string). */
    def put(cellRef: String, value: Int): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      sheet.put(ref, CellValue.Number(BigDecimal(value)))

    /** Put a Long value at cell reference (string). */
    def put(cellRef: String, value: Long): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      sheet.put(ref, CellValue.Number(BigDecimal(value)))

    /** Put a Double value at cell reference (string). */
    def put(cellRef: String, value: Double): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      sheet.put(ref, CellValue.Number(BigDecimal(value)))

    /** Put a BigDecimal value at cell reference (string). */
    def put(cellRef: String, value: BigDecimal): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      sheet.put(ref, CellValue.Number(value))

    /** Put a Boolean value at cell reference (string). */
    def put(cellRef: String, value: Boolean): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      sheet.put(ref, CellValue.Bool(value))

    /** Put a LocalDate value at cell reference (string). */
    def put(cellRef: String, value: LocalDate): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      sheet.put(ref, CellValue.DateTime(value.atStartOfDay))

    /** Put a LocalDateTime value at cell reference (string). */
    def put(cellRef: String, value: LocalDateTime): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      sheet.put(ref, CellValue.DateTime(value))

    /** Put a RichText value at cell reference (string). */
    def put(cellRef: String, value: RichText): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      sheet.put(ref, CellValue.RichText(value))

  // ========== Sheet Extensions: Styled Data Operations ==========

  extension (sheet: Sheet)
    /**
     * Put a String value with inline styling.
     *
     * @param cellRef
     *   Cell reference like "A1"
     * @param value
     *   String value
     * @param style
     *   CellStyle to apply
     * @return
     *   Updated sheet with value and style
     * @throws com.tjclp.xl.error.XLException
     *   if cellRef is invalid
     */
    def put(cellRef: String, value: String, style: CellStyle): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      val updated = sheet.put(ref, CellValue.Text(value))
      updated.withCellStyle(ref, style)

    /** Put an Int value with inline styling. */
    def put(cellRef: String, value: Int, style: CellStyle): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      val updated = sheet.put(ref, CellValue.Number(BigDecimal(value)))
      updated.withCellStyle(ref, style)

    /** Put a Long value with inline styling. */
    def put(cellRef: String, value: Long, style: CellStyle): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      val updated = sheet.put(ref, CellValue.Number(BigDecimal(value)))
      updated.withCellStyle(ref, style)

    /** Put a Double value with inline styling. */
    def put(cellRef: String, value: Double, style: CellStyle): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      val updated = sheet.put(ref, CellValue.Number(BigDecimal(value)))
      updated.withCellStyle(ref, style)

    /** Put a BigDecimal value with inline styling. */
    def put(cellRef: String, value: BigDecimal, style: CellStyle): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      val updated = sheet.put(ref, CellValue.Number(value))
      updated.withCellStyle(ref, style)

    /** Put a Boolean value with inline styling. */
    def put(cellRef: String, value: Boolean, style: CellStyle): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      val updated = sheet.put(ref, CellValue.Bool(value))
      updated.withCellStyle(ref, style)

    /** Put a LocalDate value with inline styling. */
    def put(cellRef: String, value: LocalDate, style: CellStyle): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      val updated = sheet.put(ref, CellValue.DateTime(value.atStartOfDay))
      updated.withCellStyle(ref, style)

    /** Put a LocalDateTime value with inline styling. */
    def put(cellRef: String, value: LocalDateTime, style: CellStyle): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      val updated = sheet.put(ref, CellValue.DateTime(value))
      updated.withCellStyle(ref, style)

    /** Put a RichText value with inline styling. */
    def put(cellRef: String, value: RichText, style: CellStyle): Sheet =
      val ref = toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference").unsafe
      val updated = sheet.put(ref, CellValue.RichText(value))
      updated.withCellStyle(ref, style)

  // ========== Sheet Extensions: Style Operations ==========

  extension (sheet: Sheet)
    /**
     * Apply style to a cell or range (template pattern).
     *
     * This method enables template-first workflows: format structure before filling data.
     *
     * '''Example: Template-first workflow'''
     * {{{
     * val template = Sheet("Q1 Sales")
     *   .applyStyle("A1:E1", titleStyle)       // Title row
     *   .applyStyle("A2:E2", headerStyle)      // Header row
     *   .applyStyle("A3:E100", dataStyle)      // Data rows
     *
     * val report = template
     *   .put("A1", "Q1 2025 Sales Report")
     *   .put("A2", "Product")
     * }}}
     *
     * Automatically detects whether ref is a single cell or range (by presence of ":").
     *
     * @param ref
     *   Cell reference ("A1") or range ("A1:B10")
     * @param cellStyle
     *   CellStyle to apply
     * @return
     *   Updated sheet with style applied
     * @throws com.tjclp.xl.error.XLException
     *   if ref is invalid
     */
    def applyStyle(ref: String, cellStyle: CellStyle): Sheet =
      if ref.contains(":") then
        // Range styling
        val range = toXLResult(CellRange.parse(ref), ref, "Invalid range").unsafe
        sheet.withRangeStyle(range, cellStyle)
      else
        // Single cell styling
        val aref = toXLResult(ARef.parse(ref), ref, "Invalid cell reference").unsafe
        sheet.withCellStyle(aref, cellStyle)

  // ========== Sheet Extensions: Lookup Operations ==========

  extension (sheet: Sheet)
    /**
     * Get cell at reference (string), returning Option.
     *
     * This is naturally safe - returns None for missing cells or invalid refs, no exceptions.
     *
     * @param cellRef
     *   Cell reference like "A1"
     * @return
     *   Some(cell) if present and valid ref, None otherwise
     */
    def getCell(cellRef: String): Option[Cell] =
      ARef.parse(cellRef).toOption.flatMap(ref => sheet.cells.get(ref))

    /**
     * Get all cells in range (string).
     *
     * @param rangeRef
     *   Range like "A1:B10"
     * @return
     *   List of cells in range (only cells that exist)
     * @throws com.tjclp.xl.error.XLException
     *   if rangeRef is invalid
     */
    def getCells(rangeRef: String): List[Cell] =
      val range = toXLResult(CellRange.parse(rangeRef), rangeRef, "Invalid range").unsafe
      range.cells.flatMap(sheet.cells.get).toList

  // ========== Sheet Extensions: Merge Operations ==========

  extension (sheet: Sheet)
    /**
     * Merge cells in range (string).
     *
     * @param rangeRef
     *   Range like "A1:B1"
     * @return
     *   Updated sheet with merged range
     * @throws com.tjclp.xl.error.XLException
     *   if rangeRef is invalid
     */
    def merge(rangeRef: String): Sheet =
      val range = toXLResult(CellRange.parse(rangeRef), rangeRef, "Invalid range").unsafe
      sheet.merge(range)

  // ========== XLResult[Workbook] Extensions: Chainable Operations ==========

  extension (result: XLResult[Workbook])
    /**
     * Add sheet to workbook (chainable).
     *
     * Enables chaining workbook operations without intermediate .unsafe calls:
     * {{{
     * val wb = Workbook.empty
     *   .addSheet(report)
     *   .addSheet(quickSheet)
     *   .unsafe  // Single unwrap at the end
     * }}}
     *
     * @param sheet
     *   Sheet to add
     * @return
     *   XLResult[Workbook] for further chaining
     */
    @annotation.targetName("addSheetChainable")
    def addSheet(sheet: Sheet): XLResult[Workbook] =
      result.flatMap(_.put(sheet))

    /**
     * Put sheet (add-or-replace, chainable).
     *
     * @param sheet
     *   Sheet to add or replace
     * @return
     *   XLResult[Workbook] for further chaining
     */
    @annotation.targetName("putSheetChainable")
    def put(sheet: Sheet): XLResult[Workbook] =
      result.flatMap(_.put(sheet))

    /**
     * Update sheet by name (chainable).
     *
     * @param name
     *   Sheet name
     * @param f
     *   Transform function
     * @return
     *   XLResult[Workbook] for further chaining
     */
    @annotation.targetName("updateSheetChainable")
    def update(name: String, f: Sheet => Sheet): XLResult[Workbook] =
      result.flatMap(_.update(name, f))

    /**
     * Remove sheet by name (chainable).
     *
     * @param name
     *   Sheet name
     * @return
     *   XLResult[Workbook] for further chaining
     */
    @annotation.targetName("removeSheetChainable")
    def remove(name: String): XLResult[Workbook] =
      result.flatMap(wb =>
        toXLResult(SheetName(name), name, "Invalid sheet name").flatMap(wb.remove)
      )

  // ========== Workbook Extensions: String-Based Lookups ==========

  extension (workbook: Workbook)
    /**
     * Get sheet by name (string).
     *
     * @param name
     *   Sheet name
     * @return
     *   Some(sheet) if found, None otherwise
     */
    def getSheet(name: String): Option[Sheet] =
      workbook(name).toOption
