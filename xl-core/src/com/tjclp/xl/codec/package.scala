package com.tjclp.xl.codec

import com.tjclp.xl.*
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cell.Cell
import com.tjclp.xl.style.CellStyle

import java.time.{LocalDate, LocalDateTime}

/** Extension methods for type-safe cell operations using codecs */
extension (sheet: Sheet)

  /**
   * Read a typed value from a cell using CellCodec.
   *
   * Returns Right(None) if cell is empty, Right(Some(value)) if successfully decoded, or
   * Left(error) if there's a type mismatch.
   *
   * @tparam A
   *   The type to decode to (must have a CellCodec instance)
   * @param ref
   *   The cell reference
   * @return
   *   Either[CodecError, Option[A]] - Right(None) if empty, Right(Some(value)) if success,
   *   Left(error) if type mismatch
   */
  def readTyped[A: CellCodec](ref: ARef): Either[CodecError, Option[A]] =
    sheet.cells.get(ref) match
      case None => Right(None) // Cell doesn't exist
      case Some(c) => CellCodec[A].read(c)

  /**
   * Put a typed value to a cell with auto-inferred style.
   *
   * Uses CellCodec to write the value and apply appropriate formatting (e.g., date format for
   * LocalDate, decimal format for BigDecimal).
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
   * Accepts (ARef, Any) pairs and uses runtime pattern matching to resolve the appropriate
   * CellCodec for each value. Auto-infers styles based on value types (dates get date format,
   * decimals get number format, etc.).
   *
   * Note: Unsupported types are silently skipped. Supported types: String, Int, Long, Double,
   * BigDecimal, Boolean, LocalDate, LocalDateTime, RichText.
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
    // Single-pass: build cells and collect styles simultaneously
    val cells = scala.collection.mutable.ArrayBuffer[Cell]()
    val cellsWithStyles = scala.collection.mutable.ArrayBuffer[(ARef, CellStyle)]()
    var registry = sheet.styleRegistry

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
        case v: String => processValue(ref, v)
        case v: Int => processValue(ref, v)
        case v: Long => processValue(ref, v)
        case v: Double => processValue(ref, v)
        case v: BigDecimal => processValue(ref, v)
        case v: Boolean => processValue(ref, v)
        case v: LocalDate => processValue(ref, v)
        case v: LocalDateTime => processValue(ref, v)
        case v: RichText => processValue(ref, v)
        case _ => () // Silently skip unsupported types (documented in scaladoc)
    }

    // Update sheet with new registry and cells
    val withCells = sheet.copy(styleRegistry = registry).putAll(cells)

    // Apply styles in batch
    cellsWithStyles.foldLeft(withCells) { case (s, (ref, style)) =>
      s.withCellStyle(ref, style)
    }
