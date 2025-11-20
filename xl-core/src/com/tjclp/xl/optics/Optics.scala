package com.tjclp.xl.optics

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.units.StyleId

/**
 * Predefined optics for common Sheet/Cell operations.
 *
 * Provides Lens and Optional types for functional, law-governed transformations without external
 * dependencies.
 *
 * Laws:
 *   - Lens: get-put, put-get, put-put
 *   - Optional: get-put, put-get (when value present)
 */
object Optics:

  /** Lens for accessing all cells in a sheet */
  val sheetCells: Lens[Sheet, Map[ARef, Cell]] =
    Lens(
      get = s => s.cells,
      set = (cells, s) => s.copy(cells = cells)
    )

  /** Optional for accessing a specific cell (returns empty cell if not present) */
  def cellAt(ref: ARef): Optional[Sheet, Cell] =
    Optional(
      getOption = s => s.cells.get(ref).orElse(Some(Cell.empty(ref))),
      set = (c, s) => s.put(c)
    )

  /** Lens for accessing cell value */
  val cellValue: Lens[Cell, CellValue] =
    Lens(
      get = c => c.value,
      set = (v, c) => c.withValue(v)
    )

  /** Lens for accessing cell styleId */
  val cellStyleId: Lens[Cell, Option[StyleId]] =
    Lens(
      get = c => c.styleId,
      set = (sid, c) =>
        sid match
          case Some(id) => c.withStyle(id)
          case None => c.clearStyle
    )

  /** Lens for accessing sheet name */
  val sheetName: Lens[Sheet, SheetName] =
    Lens(
      get = s => s.name,
      set = (n, s) => s.copy(name = n)
    )

  /** Lens for accessing merged ranges */
  val mergedRanges: Lens[Sheet, Set[CellRange]] =
    Lens(
      get = s => s.mergedRanges,
      set = (ranges, s) => s.copy(mergedRanges = ranges)
    )
