package com.tjclp.xl

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.style.CellStyle
import com.tjclp.xl.workbook.Workbook

import java.time.{LocalDate, LocalDateTime}

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
 *   .applyStyle("A1:B1", boldStyle)
 *   .unsafe  // Single boundary point
 * }}}
 *
 * '''Styling approaches:'''
 *   - `.applyStyle(range, style)` - Template pattern (format before data)
 *   - `.put(ref, value, style)` - Inline styling
 *   - RichText - Intra-cell formatting
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
    /** Put String value at cell reference (returns XLResult for chaining). */
    def put(cellRef: String, value: String): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map(ref => sheet.put(ref, CellValue.Text(value)))

    /** Put Int value at cell reference. */
    def put(cellRef: String, value: Int): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map(ref => sheet.put(ref, CellValue.Number(BigDecimal(value))))

    /** Put Long value at cell reference. */
    def put(cellRef: String, value: Long): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map(ref => sheet.put(ref, CellValue.Number(BigDecimal(value))))

    /** Put Double value at cell reference. */
    def put(cellRef: String, value: Double): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map(ref => sheet.put(ref, CellValue.Number(BigDecimal(value))))

    /** Put BigDecimal value at cell reference. */
    def put(cellRef: String, value: BigDecimal): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map(ref => sheet.put(ref, CellValue.Number(value)))

    /** Put Boolean value at cell reference. */
    def put(cellRef: String, value: Boolean): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map(ref => sheet.put(ref, CellValue.Bool(value)))

    /** Put LocalDate value at cell reference. */
    def put(cellRef: String, value: LocalDate): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map(ref => sheet.put(ref, CellValue.DateTime(value.atStartOfDay)))

    /** Put LocalDateTime value at cell reference. */
    def put(cellRef: String, value: LocalDateTime): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map(ref => sheet.put(ref, CellValue.DateTime(value)))

    /** Put RichText value at cell reference. */
    def put(cellRef: String, value: RichText): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map(ref => sheet.put(ref, CellValue.RichText(value)))

  // ========== Sheet Extensions: Styled Data Operations ==========

  extension (sheet: Sheet)
    /** Put String value with inline styling. */
    def put(cellRef: String, value: String, cellStyle: CellStyle): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map { ref =>
          val updated = sheet.put(ref, CellValue.Text(value))
          updated.withCellStyle(ref, cellStyle)
        }

    /** Put Int value with inline styling. */
    def put(cellRef: String, value: Int, cellStyle: CellStyle): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map { ref =>
          val updated = sheet.put(ref, CellValue.Number(BigDecimal(value)))
          updated.withCellStyle(ref, cellStyle)
        }

    /** Put Long value with inline styling. */
    def put(cellRef: String, value: Long, cellStyle: CellStyle): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map { ref =>
          val updated = sheet.put(ref, CellValue.Number(BigDecimal(value)))
          updated.withCellStyle(ref, cellStyle)
        }

    /** Put Double value with inline styling. */
    def put(cellRef: String, value: Double, cellStyle: CellStyle): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map { ref =>
          val updated = sheet.put(ref, CellValue.Number(BigDecimal(value)))
          updated.withCellStyle(ref, cellStyle)
        }

    /** Put BigDecimal value with inline styling. */
    def put(cellRef: String, value: BigDecimal, cellStyle: CellStyle): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map { ref =>
          val updated = sheet.put(ref, CellValue.Number(value))
          updated.withCellStyle(ref, cellStyle)
        }

    /** Put Boolean value with inline styling. */
    def put(cellRef: String, value: Boolean, cellStyle: CellStyle): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map { ref =>
          val updated = sheet.put(ref, CellValue.Bool(value))
          updated.withCellStyle(ref, cellStyle)
        }

    /** Put LocalDate value with inline styling. */
    def put(cellRef: String, value: LocalDate, cellStyle: CellStyle): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map { ref =>
          val updated = sheet.put(ref, CellValue.DateTime(value.atStartOfDay))
          updated.withCellStyle(ref, cellStyle)
        }

    /** Put LocalDateTime value with inline styling. */
    def put(cellRef: String, value: LocalDateTime, cellStyle: CellStyle): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map { ref =>
          val updated = sheet.put(ref, CellValue.DateTime(value))
          updated.withCellStyle(ref, cellStyle)
        }

    /** Put RichText value with inline styling. */
    def put(cellRef: String, value: RichText, cellStyle: CellStyle): XLResult[Sheet] =
      toXLResult(ARef.parse(cellRef), cellRef, "Invalid cell reference")
        .map { ref =>
          val updated = sheet.put(ref, CellValue.RichText(value))
          updated.withCellStyle(ref, cellStyle)
        }

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
    @annotation.targetName("applyStyleSheet")
    def applyStyle(ref: String, cellStyle: CellStyle): XLResult[Sheet] =
      if ref.contains(":") then
        toXLResult(CellRange.parse(ref), ref, "Invalid range")
          .map(range => sheet.withRangeStyle(range, cellStyle))
      else
        toXLResult(ARef.parse(ref), ref, "Invalid cell reference")
          .map(aref => sheet.withCellStyle(aref, cellStyle))

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
    def getCell(cellRef: String): Option[Cell] =
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
    def getCells(rangeRef: String): List[Cell] =
      CellRange
        .parse(rangeRef)
        .toOption
        .map(r => r.cells.flatMap(sheet.cells.get).toList)
        .getOrElse(List.empty)

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
    /** Put String value (chainable). */
    @annotation.targetName("putStringChainable")
    def put(cellRef: String, value: String): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value))

    /** Put Int value (chainable). */
    @annotation.targetName("putIntChainable")
    def put(cellRef: String, value: Int): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value))

    /** Put Long value (chainable). */
    @annotation.targetName("putLongChainable")
    def put(cellRef: String, value: Long): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value))

    /** Put Double value (chainable). */
    @annotation.targetName("putDoubleChainable")
    def put(cellRef: String, value: Double): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value))

    /** Put BigDecimal value (chainable). */
    @annotation.targetName("putBigDecimalChainable")
    def put(cellRef: String, value: BigDecimal): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value))

    /** Put Boolean value (chainable). */
    @annotation.targetName("putBooleanChainable")
    def put(cellRef: String, value: Boolean): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value))

    /** Put LocalDate value (chainable). */
    @annotation.targetName("putLocalDateChainable")
    def put(cellRef: String, value: LocalDate): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value))

    /** Put LocalDateTime value (chainable). */
    @annotation.targetName("putLocalDateTimeChainable")
    def put(cellRef: String, value: LocalDateTime): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value))

    /** Put RichText value (chainable). */
    @annotation.targetName("putRichTextChainable")
    def put(cellRef: String, value: RichText): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value))

    /** Put String with style (chainable). */
    @annotation.targetName("putStringStyledChainable")
    def put(cellRef: String, value: String, cellStyle: CellStyle): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value, cellStyle))

    /** Put Int with style (chainable). */
    @annotation.targetName("putIntStyledChainable")
    def put(cellRef: String, value: Int, cellStyle: CellStyle): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value, cellStyle))

    /** Put Long with style (chainable). */
    @annotation.targetName("putLongStyledChainable")
    def put(cellRef: String, value: Long, cellStyle: CellStyle): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value, cellStyle))

    /** Put Double with style (chainable). */
    @annotation.targetName("putDoubleStyledChainable")
    def put(cellRef: String, value: Double, cellStyle: CellStyle): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value, cellStyle))

    /** Put BigDecimal with style (chainable). */
    @annotation.targetName("putBigDecimalStyledChainable")
    def put(cellRef: String, value: BigDecimal, cellStyle: CellStyle): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value, cellStyle))

    /** Put Boolean with style (chainable). */
    @annotation.targetName("putBooleanStyledChainable")
    def put(cellRef: String, value: Boolean, cellStyle: CellStyle): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value, cellStyle))

    /** Put LocalDate with style (chainable). */
    @annotation.targetName("putLocalDateStyledChainable")
    def put(cellRef: String, value: LocalDate, cellStyle: CellStyle): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value, cellStyle))

    /** Put LocalDateTime with style (chainable). */
    @annotation.targetName("putLocalDateTimeStyledChainable")
    def put(cellRef: String, value: LocalDateTime, cellStyle: CellStyle): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value, cellStyle))

    /** Put RichText with style (chainable). */
    @annotation.targetName("putRichTextStyledChainable")
    def put(cellRef: String, value: RichText, cellStyle: CellStyle): XLResult[Sheet] =
      result.flatMap(_.put(cellRef, value, cellStyle))

    /**
     * Apply style (chainable).
     */
    @annotation.targetName("applyStyleSheetChainable")
    def applyStyle(ref: String, cellStyle: CellStyle): XLResult[Sheet] =
      result.flatMap(_.applyStyle(ref, cellStyle))

    /** Merge range (chainable). */
    def merge(rangeRef: String): XLResult[Sheet] =
      result.flatMap(_.merge(rangeRef))

  // ========== XLResult[Workbook] Extensions: Chainable Operations ==========

  extension (result: XLResult[Workbook])
    /** Add sheet to workbook (chainable). */
    @annotation.targetName("addSheetChainable")
    def addSheet(sheet: Sheet): XLResult[Workbook] =
      result.flatMap(_.put(sheet))

    /** Put sheet (chainable). */
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
    def getSheet(name: String): Option[Sheet] =
      workbook(name).toOption
