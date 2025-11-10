package com.tjclp.xl

import cats.Monoid

/** Patch ADT for Sheet updates with monoid semantics.
  *
  * Patches can be composed using the Monoid instance, allowing
  * batch operations to be built declaratively.
  *
  * Laws:
  * - Associativity: (p1 |+| p2) |+| p3 == p1 |+| (p2 |+| p3)
  * - Identity: Patch.empty |+| p == p == p |+| Patch.empty
  * - Idempotence: Applying the same patch twice yields the same result
  */
enum Patch:
  /** Put a cell value at a reference */
  case Put(ref: ARef, value: CellValue)

  /** Set style for a cell */
  case SetStyle(ref: ARef, styleId: Int)

  /** Clear style for a cell */
  case ClearStyle(ref: ARef)

  /** Merge cells in a range */
  case Merge(range: CellRange)

  /** Unmerge cells in a range */
  case Unmerge(range: CellRange)

  /** Set column properties */
  case SetColumnProperties(col: Column, props: ColumnProperties)

  /** Set row properties */
  case SetRowProperties(row: Row, props: RowProperties)

  /** Remove cell at reference */
  case Remove(ref: ARef)

  /** Remove all cells in range */
  case RemoveRange(range: CellRange)

  /** Batch multiple patches together */
  case Batch(patches: Vector[Patch])

object Patch:
  /** Empty patch (identity element) */
  val empty: Patch = Batch(Vector.empty)

  /** Combine two patches into a batch.
    *
    * Flattens nested batches to maintain a flat structure.
    * Later patches override earlier ones for the same reference.
    */
  def combine(p1: Patch, p2: Patch): Patch = (p1, p2) match
    case (Batch(ps1), Batch(ps2)) => Batch(ps1 ++ ps2)
    case (Batch(ps1), p2) => Batch(ps1 :+ p2)
    case (p1, Batch(ps2)) => Batch(p1 +: ps2)
    case (p1, p2) => Batch(Vector(p1, p2))

  /** Monoid instance for Patch composition */
  given Monoid[Patch] with
    def empty: Patch = Patch.empty
    def combine(x: Patch, y: Patch): Patch = Patch.combine(x, y)

  /** Apply a patch to a sheet, returning the modified sheet.
    *
    * Patches are applied left-to-right. Later patches override earlier ones
    * for conflicting operations (e.g., two Puts to the same reference).
    *
    * @param sheet The sheet to modify
    * @param patch The patch to apply
    * @return Either an error or the modified sheet
    */
  def applyPatch(sheet: Sheet, patch: Patch): XLResult[Sheet] = patch match
    case Put(ref, value) =>
      Right(sheet.put(ref, value))

    case SetStyle(ref, styleId) =>
      val cell = sheet(ref).withStyle(styleId)
      Right(sheet.put(cell))

    case ClearStyle(ref) =>
      val cell = sheet(ref).clearStyle
      Right(sheet.put(cell))

    case Merge(range) =>
      Right(sheet.merge(range))

    case Unmerge(range) =>
      Right(sheet.unmerge(range))

    case SetColumnProperties(col, props) =>
      Right(sheet.setColumnProperties(col, props))

    case SetRowProperties(row, props) =>
      Right(sheet.setRowProperties(row, props))

    case Remove(ref) =>
      Right(sheet.remove(ref))

    case RemoveRange(range) =>
      Right(sheet.removeRange(range))

    case Batch(patches) =>
      patches.foldLeft[XLResult[Sheet]](Right(sheet)) { (result, p) =>
        result.flatMap(s => applyPatch(s, p))
      }

  /** Apply multiple patches in sequence */
  def applyPatches(sheet: Sheet, patches: Iterable[Patch]): XLResult[Sheet] =
    applyPatch(sheet, Batch(patches.toVector))

  /** Convenience extensions for Sheet */
  extension (sheet: Sheet)
    /** Apply a patch to this sheet */
    def applyPatch(patch: Patch): XLResult[Sheet] =
      Patch.applyPatch(sheet, patch)

    /** Apply multiple patches to this sheet */
    def applyPatches(patches: Patch*): XLResult[Sheet] =
      Patch.applyPatches(sheet, patches)

/** Syntax for patch composition using cats Monoid */
object syntax:
  export cats.syntax.monoid.given
  export cats.syntax.semigroup.given
