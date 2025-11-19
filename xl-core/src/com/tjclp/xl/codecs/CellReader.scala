package com.tjclp.xl.codecs

import com.tjclp.xl.cells.Cell

/**
 * Read a typed value from a Cell.
 *
 * Returns Right(None) if the cell is empty, Right(Some(value)) if successfully decoded, or
 * Left(errors) if there's a type mismatch or parse errors.
 */
trait CellReader[A]:
  def read(cell: Cell): Either[CodecError, Option[A]]
