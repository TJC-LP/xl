package com.tjclp.xl.optics

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheet.Sheet
import com.tjclp.xl.style.units.StyleId

/**
 * Focus DSL extensions for Sheet.
 *
 * Usage:
 * {{{
 *   import com.tjclp.xl.optics.*
 *
 *   sheet.modifyValue(ref"A1") {
 *     case CellValue.Text(s) => CellValue.Text(s.toUpperCase)
 *     case other => other
 *   }
 * }}}
 */
object syntax:
  extension (sheet: Sheet)
    /**
     * Focus on a specific cell for composable updates.
     *
     * Returns an Optional that can be used to query or modify the cell.
     */
    def focus(ref: ARef): Optional[Sheet, Cell] = Optics.cellAt(ref)

    /**
     * Modify a cell using a transformation function.
     *
     * If the cell doesn't exist, creates an empty cell and applies the transformation.
     *
     * Example:
     * {{{
     *   sheet.modifyCell(ref"A1")(_.withValue(CellValue.Text("Updated")))
     * }}}
     */
    def modifyCell(ref: ARef)(f: Cell => Cell): Sheet =
      focus(ref).modify(f)(sheet)

    /**
     * Modify a cell's value using a transformation function.
     *
     * Example:
     * {{{
     *   sheet.modifyValue(ref"A1") {
     *     case CellValue.Text(s) => CellValue.Text(s.toUpperCase)
     *     case other => other
     *   }
     * }}}
     */
    def modifyValue(ref: ARef)(f: CellValue => CellValue): Sheet =
      focus(ref).modify(c => Optics.cellValue.update(f)(c))(sheet)

    /**
     * Modify a cell's style ID.
     *
     * Example:
     * {{{
     *   sheet.modifyStyleId(ref"A1")(_.map(id => StyleId(id.value + 1)))
     * }}}
     */
    def modifyStyleId(ref: ARef)(f: Option[StyleId] => Option[StyleId]): Sheet =
      focus(ref).modify(c => Optics.cellStyleId.update(f)(c))(sheet)

export syntax.*
