package com.tjclp.xl

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cell.{Cell, CellValue}
import com.tjclp.xl.style.StyleId

/**
 * Lightweight optics library for composable updates.
 *
 * Provides Lens and Optional types for functional, law-governed transformations without external
 * dependencies.
 *
 * Laws:
 *   - Lens: get-put, put-get, put-put
 *   - Optional: get-put, put-get (when value present)
 *
 * Usage:
 * {{{
 *   import com.tjclp.xl.optics.*
 *
 *   sheet.modifyValue(cell"A1") {
 *     case CellValue.Text(s) => CellValue.Text(s.toUpperCase)
 *     case other => other
 *   }
 * }}}
 */

/** Lens: Total getter and setter for a field */
final case class Lens[S, A](get: S => A, set: (A, S) => S):
  /** Modify the focused field using a function */
  def modify(f: A => A)(s: S): S = set(f(get(s)), s)

  /** Update the focused field using a function */
  def update(f: A => A)(s: S): S = set(f(get(s)), s)

  /** Compose with another lens */
  def andThen[B](other: Lens[A, B]): Lens[S, B] =
    Lens(
      get = s => other.get(this.get(s)),
      set = (b, s) => this.set(other.set(b, this.get(s)), s)
    )

/** Optional: Partial getter and setter (may not have a value) */
final case class Optional[S, A](getOption: S => Option[A], set: (A, S) => S):
  /** Modify the focused field if it exists */
  def modify(f: A => A)(s: S): S =
    getOption(s) match
      case Some(a) => set(f(a), s)
      case None => s

  /** Get the value or use a default */
  def getOrElse(default: A)(s: S): A =
    getOption(s).getOrElse(default)

/** Predefined optics for common Sheet/Cell operations */
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

/** Focus DSL extensions for Sheet */
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
   *   sheet.modifyCell(cell"A1")(_.withValue(CellValue.Text("Updated")))
   * }}}
   */
  def modifyCell(ref: ARef)(f: Cell => Cell): Sheet =
    focus(ref).modify(f)(sheet)

  /**
   * Modify a cell's value using a transformation function.
   *
   * Example:
   * {{{
   *   sheet.modifyValue(cell"A1") {
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
   *   sheet.modifyStyleId(cell"A1")(_.map(id => StyleId(id.value + 1)))
   * }}}
   */
  def modifyStyleId(ref: ARef)(f: Option[StyleId] => Option[StyleId]): Sheet =
    focus(ref).modify(c => Optics.cellStyleId.update(f)(c))(sheet)
