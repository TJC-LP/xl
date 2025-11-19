package com.tjclp.xl.codecs

import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.Cell
import com.tjclp.xl.codecs.{CellCodec, CodecError} // Explicit import for companion object
import com.tjclp.xl.sheets.syntax.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

import java.time.{LocalDate, LocalDateTime}

/** Extension methods for type-safe cell operations using codecs */
object syntax:
  extension (sheet: Sheet)

    /**
     * Read a typed value from a cell using CellCodec.
     *
     * Returns Right(None) if cell is empty, Right(Some(value)) if successfully decoded, or
     * Left(errors) if there's a type mismatch.
     *
     * @tparam A
     *   The type to decode to (must have a CellCodec instance)
     * @param ref
     *   The cell reference
     * @return
     *   Either[CodecError, Option[A]] - Right(None) if empty, Right(Some(value)) if success,
     *   Left(errors) if type mismatch
     */
    def readTyped[A: CellCodec](ref: ARef): Either[CodecError, Option[A]] =
      sheet.cells.get(ref) match
        case None => Right(None) // Cell doesn't exist
        case Some(c) => CellCodec[A].read(c)

export syntax.*
