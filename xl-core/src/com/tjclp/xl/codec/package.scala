package com.tjclp.xl.codec

import com.tjclp.xl.*
import java.time.{LocalDate, LocalDateTime}

/** Extension methods for type-safe cell operations using codecs */
extension (sheet: Sheet)

  /**
   * Read a typed value from a cell using CellCodec.
   *
   * Returns Right(None) if cell is empty, Right(Some(value)) if successfully decoded, or Left(error) if there's a
   * type mismatch.
   *
   * @tparam A
   *   The type to decode to (must have a CellCodec instance)
   * @param ref
   *   The cell reference
   * @return
   *   Either[CodecError, Option[A]] - Right(None) if empty, Right(Some(value)) if success, Left(error) if type
   *   mismatch
   */
  def readTyped[A: CellCodec](ref: ARef): Either[CodecError, Option[A]] =
    sheet.cells.get(ref) match
      case None => Right(None) // Cell doesn't exist
      case Some(c) => CellCodec[A].read(c)

  /**
   * Put a typed value to a cell with auto-inferred style.
   *
   * Uses CellCodec to write the value and apply appropriate formatting (e.g., date format for LocalDate, decimal
   * format for BigDecimal).
   *
   * @tparam A
   *   The type to encode (must have a CellCodec instance)
   * @param ref
   *   The cell reference
   * @param value
   *   The value to write
   * @return
   *   Updated sheet with the cell value and style applied
   */
  def putTyped[A: CellCodec](ref: ARef, value: A): Sheet =
    val (cellValue, styleOpt) = CellCodec[A].write(value)
    val updated = sheet.put(ref, cellValue)
    styleOpt match
      case Some(style) => updated.withCellStyle(ref, style)
      case None => updated

  /**
   * Put multiple cells with mixed types, leveraging existing putAll.
   *
   * Accepts (ARef, Any) pairs and uses runtime pattern matching to resolve the appropriate CellCodec for each value.
   * Auto-infers styles based on value types (dates get date format, decimals get number format, etc.).
   *
   * Example:
   * {{{
   * sheet.putMixed(
   *   cell"A1" -> "Revenue",
   *   cell"B1" -> LocalDate.of(2025, 11, 10),
   *   cell"C1" -> BigDecimal("123.45")
   * )
   * }}}
   *
   * @param updates
   *   Varargs of (ARef, Any) pairs
   * @return
   *   Updated sheet with all cells and styles applied
   */
  def putMixed(updates: (ARef, Any)*): Sheet =
    // Convert (ARef, Any) pairs to Cell objects using CellCodec pattern matching
    val cells = scala.collection.mutable.ArrayBuffer[Cell]()
    var currentSheet = sheet

    updates.foreach { (ref, value) =>
      value match
        case v: String =>
          val (cellValue, styleOpt) = CellCodec[String].write(v)
          cells += Cell(ref, cellValue)
          styleOpt.foreach(style => currentSheet = currentSheet.copy(styleRegistry = currentSheet.styleRegistry.register(style)._1))

        case v: Int =>
          val (cellValue, styleOpt) = CellCodec[Int].write(v)
          cells += Cell(ref, cellValue)
          styleOpt.foreach(style => currentSheet = currentSheet.copy(styleRegistry = currentSheet.styleRegistry.register(style)._1))

        case v: Long =>
          val (cellValue, styleOpt) = CellCodec[Long].write(v)
          cells += Cell(ref, cellValue)
          styleOpt.foreach(style => currentSheet = currentSheet.copy(styleRegistry = currentSheet.styleRegistry.register(style)._1))

        case v: Double =>
          val (cellValue, styleOpt) = CellCodec[Double].write(v)
          cells += Cell(ref, cellValue)
          styleOpt.foreach(style => currentSheet = currentSheet.copy(styleRegistry = currentSheet.styleRegistry.register(style)._1))

        case v: BigDecimal =>
          val (cellValue, styleOpt) = CellCodec[BigDecimal].write(v)
          cells += Cell(ref, cellValue)
          styleOpt.foreach(style => currentSheet = currentSheet.copy(styleRegistry = currentSheet.styleRegistry.register(style)._1))

        case v: Boolean =>
          val (cellValue, styleOpt) = CellCodec[Boolean].write(v)
          cells += Cell(ref, cellValue)
          styleOpt.foreach(style => currentSheet = currentSheet.copy(styleRegistry = currentSheet.styleRegistry.register(style)._1))

        case v: LocalDate =>
          val (cellValue, styleOpt) = CellCodec[LocalDate].write(v)
          cells += Cell(ref, cellValue)
          styleOpt.foreach(style => currentSheet = currentSheet.copy(styleRegistry = currentSheet.styleRegistry.register(style)._1))

        case v: LocalDateTime =>
          val (cellValue, styleOpt) = CellCodec[LocalDateTime].write(v)
          cells += Cell(ref, cellValue)
          styleOpt.foreach(style => currentSheet = currentSheet.copy(styleRegistry = currentSheet.styleRegistry.register(style)._1))

        case v: RichText =>
          val (cellValue, styleOpt) = CellCodec[RichText].write(v)
          cells += Cell(ref, cellValue)
          styleOpt.foreach(style => currentSheet = currentSheet.copy(styleRegistry = currentSheet.styleRegistry.register(style)._1))

        case _ => () // Ignore unsupported types
    }

    // Use existing putAll + apply styles
    val withCells = currentSheet.putAll(cells)

    // Apply styles for cells that need them
    updates.foldLeft(withCells) { case (s, (ref, value)) =>
      value match
        case _: String | _: Boolean | _: RichText => s // No cell-level style (RichText has run-level formatting)
        case v: Int =>
          val (_, styleOpt) = CellCodec[Int].write(v)
          styleOpt.fold(s)(style => s.withCellStyle(ref, style))
        case v: Long =>
          val (_, styleOpt) = CellCodec[Long].write(v)
          styleOpt.fold(s)(style => s.withCellStyle(ref, style))
        case v: Double =>
          val (_, styleOpt) = CellCodec[Double].write(v)
          styleOpt.fold(s)(style => s.withCellStyle(ref, style))
        case v: BigDecimal =>
          val (_, styleOpt) = CellCodec[BigDecimal].write(v)
          styleOpt.fold(s)(style => s.withCellStyle(ref, style))
        case v: LocalDate =>
          val (_, styleOpt) = CellCodec[LocalDate].write(v)
          styleOpt.fold(s)(style => s.withCellStyle(ref, style))
        case v: LocalDateTime =>
          val (_, styleOpt) = CellCodec[LocalDateTime].write(v)
          styleOpt.fold(s)(style => s.withCellStyle(ref, style))
        case _ => s
    }
