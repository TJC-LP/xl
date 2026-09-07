package com.tjclp.xl.codec

import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.Cell
import com.tjclp.xl.codec.{CellCodec, CodecError} // Explicit import for companion object
import com.tjclp.xl.sheets.syntax.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

import java.time.{LocalDate, LocalDateTime}

/** Extension methods for type-safe cell operations using codecs */
object syntax:
  // Records as rows (GH-590): readRows / readRowsByHeader / headers / column / putRows / putTable
  export rowSyntax.*

  extension (sheet: Sheet)

    /**
     * Read a typed value from a cell using CellCodec.
     *
     * Returns Right(None) if cell is empty, Right(Some(value)) if successfully decoded, or
     * Left(error) if there's a type mismatch.
     *
     * Formula cells decode through their cached value (GH-477): after `recalculate()`,
     * `writeRecalculated`, or reading a book Excel saved, `Formula("A1*3", Some(Number(6)), _)`
     * reads as `Right(Some(6))` — no manual `CellValue.Formula(_, Some(v), _)` unwrapping. A
     * formula with no cached value has nothing to read and is
     * `Left(TypeMismatch(expected, formula))`. Use [[readTypedStrict]] when any formula cell must
     * be rejected regardless of its cache.
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
     * Strict typed read: like [[readTyped]], but a formula cell is ALWAYS
     * `Left(TypeMismatch(expected, formula))`, cached or not — the pre-GH-477 semantics.
     *
     * Reach for it when "is this cell a formula?" matters more than the formula's result: auditing
     * hand-entered constants against computed cells, or refusing to trust a cache that may be
     * stale. Everywhere else prefer `readTyped` / `readTypedOr` / `readTypedOpt`, which see the
     * cached value a recalculated or Excel-saved book carries.
     */
    def readTypedStrict[A: CellCodec](ref: ARef): Either[CodecError, Option[A]] =
      sheet.cells.get(ref) match
        case None => Right(None)
        case Some(c) => CellCodec[A].readStrict(c)

    /**
     * Read a typed value, falling back to a default on missing cell, empty cell, or type mismatch
     * (an uncached formula included). Total — the scripting counterpart of `readTyped`; like it,
     * sees a formula's cached value (GH-477).
     *
     * Naming convention: verb + totality-strategy suffix (`readTyped` / `readTypedOr` /
     * `readTypedOpt`; `readTypedStrict` is the formula-rejecting variant of `readTyped`).
     */
    def readTypedOr[A: CellCodec](ref: ARef, default: => A): A =
      readTyped[A](ref).toOption.flatten.getOrElse(default)

    /**
     * Read a typed value as a flat Option: None on missing cell, empty cell, or type mismatch (an
     * uncached formula included); a cached formula is `Some(cached)` (GH-477). Total. Use
     * `readTyped` when decode errors must be distinguished from absence.
     */
    def readTypedOpt[A: CellCodec](ref: ARef): Option[A] =
      readTyped[A](ref).toOption.flatten

export syntax.*
