package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.Cell
import com.tjclp.xl.styles.CellStyle

// ========== Style Application Extensions ==========

object styleSyntax:
  extension (sheet: Sheet)

    /**
     * Apply a CellStyle to a cell, registering it automatically.
     *
     * Registers the style in the sheet's styleRegistry and applies the resulting index to the cell.
     * If the style is already registered, reuses the existing index.
     */
    @annotation.targetName("withCellStyleExt")
    def withCellStyle(ref: ARef, style: CellStyle): Sheet =
      val (newRegistry, styleId) = sheet.styleRegistry.register(style)
      val cell = sheet(ref).withStyle(styleId)
      sheet.copy(
        styleRegistry = newRegistry,
        cells = sheet.cells.updated(ref, cell)
      )

    /** Apply a CellStyle to all cells in a range. */
    @annotation.targetName("withRangeStyleExt")
    def withRangeStyle(range: CellRange, style: CellStyle): Sheet =
      val (newRegistry, styleId) = sheet.styleRegistry.register(style)
      val updatedCells = range.cells.foldLeft(sheet.cells) { (cells, ref) =>
        val cell = cells.getOrElse(ref, Cell.empty(ref)).withStyle(styleId)
        cells.updated(ref, cell)
      }
      sheet.copy(
        styleRegistry = newRegistry,
        cells = updatedCells
      )

    /** Get the CellStyle for a cell (if it has one). */
    @annotation.targetName("getCellStyleExt")
    def getCellStyle(ref: ARef): Option[CellStyle] =
      sheet(ref).styleId.flatMap(sheet.styleRegistry.get)
