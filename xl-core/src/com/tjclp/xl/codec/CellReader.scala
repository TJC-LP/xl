package com.tjclp.xl.codec

import com.tjclp.xl.cells.Cell

/**
 * Read a typed value from a Cell.
 *
 * Returns Right(None) if the cell is empty, Right(Some(value)) if successfully decoded, or
 * Left(error) if there's a type mismatch or parse error.
 *
 * Formula cells (GH-477): [[read]] decodes a formula's cached value — `Formula(_, Some(v), _)`
 * reads exactly as a plain cell holding `v` would, which is what a recalculated or Excel-saved book
 * carries and what `view`/`eval` display. A formula with no cached value has nothing to decode and
 * is a `CodecError.TypeMismatch` whose `actual` is the formula itself. [[readStrict]] is the escape
 * hatch that treats EVERY formula cell as a mismatch, cached or not.
 */
trait CellReader[A]:
  def read(cell: Cell): Either[CodecError, Option[A]]

  /**
   * Strict read: never looks through a formula's cached value (the pre-GH-477 semantics).
   *
   * The default delegates to [[read]] so readers implemented outside this library keep compiling
   * and behaving as before; the built-in `CellCodec` instances override it to reject any
   * `CellValue.Formula` with `TypeMismatch(<expected type>, formula)` and otherwise decode as
   * [[read]] does.
   */
  def readStrict(cell: Cell): Either[CodecError, Option[A]] = read(cell)
