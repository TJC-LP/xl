package com.tjclp.xl.formula

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.ArrayResult

/**
 * Test-only: the whole [[ArrayResult]] as rows of cells, whose storage is not part of its API.
 * Fails the test past [[ArrayResult.MaxCells]].
 */
extension (ar: ArrayResult)
  def dense: Vector[Vector[CellValue]] =
    ar.toDense("test").fold(e => munit.Assertions.fail(s"ArrayResult.toDense: $e"), identity)
