package com.tjclp.xl.ops

import com.tjclp.xl.addressing.{CellRange, SheetName}
import com.tjclp.xl.workbooks.Workbook

/**
 * One edit resolved against the workbook it ran on: its 1-based position, the sheet THE rule gave
 * it (`None` for a workbook-level edit), and the ranges whose CONTENT it may have changed — the
 * seed of an after-edit recalculation cone. Structural and name edits report no ranges; their
 * [[EditSpec.structural]] flag says the whole book must be recalculated.
 */
final case class Planned(
  index: Int,
  edit: Edit,
  sheet: Option[SheetName],
  touched: Vector[CellRange]
) derives CanEqual

/**
 * The result of [[Edit.applyAll]]: the edited workbook (every changed sheet written back through
 * `Workbook.put`, so `modifiedSheets` is right), the [[Planned]] row of every edit, and the scope
 * after the sequence (rule 4: a rename of the default sheet moves it).
 */
final case class Applied(workbook: Workbook, planned: Vector[Planned], scope: Scope)
    derives CanEqual:

  /** The content writes grouped by sheet, in edit order. */
  def touchedBySheet: Map[SheetName, Vector[CellRange]] =
    planned
      .collect { case Planned(_, _, Some(sheet), ranges) if ranges.nonEmpty => sheet -> ranges }
      .groupMapReduce(_._1)(_._2)(_ ++ _)

  /** Whether any edit changed structure or names — the whole book needs recalculating. */
  def structural: Boolean = planned.exists(p => EditSchema.specOf(p.edit).structural)
