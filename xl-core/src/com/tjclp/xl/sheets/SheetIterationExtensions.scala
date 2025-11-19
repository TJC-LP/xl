package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{Column, Row}
import com.tjclp.xl.cells.Cell

// ========== Deterministic Iteration Helpers ==========

object iterationSyntax:
  extension (sheet: Sheet)

    /**
     * Get all cells sorted in row-major order (left-to-right, top-to-bottom).
     *
     * Provides deterministic iteration order matching the canonical write order.
     *
     * @return
     *   Vector of cells sorted by (row, column)
     */
    def cellsSorted: Vector[Cell] =
      sheet.cells.values.toVector.sortBy(c => (c.row.index0, c.col.index0))

    /**
     * Get cells grouped by row, sorted by row index.
     *
     * Each row's cells are sorted left-to-right.
     *
     * @return
     *   Vector of (row, cells) pairs sorted by row index
     */
    def rowsSorted: Vector[(Row, Vector[Cell])] =
      sheet.cells.values
        .groupBy(_.row)
        .toVector
        .sortBy(_._1.index0)
        .map { case (row, cells) => (row, cells.toVector.sortBy(_.col.index0)) }

    /**
     * Get cells grouped by column, sorted by column index.
     *
     * Each column's cells are sorted top-to-bottom.
     *
     * @return
     *   Vector of (column, cells) pairs sorted by column index
     */
    def columnsSorted: Vector[(Column, Vector[Cell])] =
      sheet.cells.values
        .groupBy(_.col)
        .toVector
        .sortBy(_._1.index0)
        .map { case (col, cells) => (col, cells.toVector.sortBy(_.row.index0)) }
