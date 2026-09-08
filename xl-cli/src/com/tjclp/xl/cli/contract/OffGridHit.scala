package com.tjclp.xl.cli.contract

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.ops.OffGridRef

/**
 * GH-628: one cell whose dragged, filled or copied formula gained a `#REF!` because a reference
 * left the grid — the sheet it is on and the core's report of the references voided.
 */
final case class OffGridHit(sheet: String, hit: OffGridRef) derives CanEqual:
  def cell: ARef = hit.cell
  def references: Vector[String] = hit.references

object OffGridHit:

  /** The core's report for `sheet`, qualified. */
  def of(sheet: SheetName, hits: Vector[OffGridRef]): Vector[OffGridHit] =
    hits.map(OffGridHit(sheet.value, _))

  private val MaxListed = 10

  /** `Data!B1 (A1)` — the cell and the source references that had nowhere to go. */
  private def describe(hit: OffGridHit): String =
    val refs = if hit.references.isEmpty then "" else s" (${hit.references.mkString(", ")})"
    s"${SheetName.quoteForFormula(hit.sheet)}!${hit.cell.toA1}$refs"

  /** The `--strict` reason, when there is one. */
  def strictReason(hits: Vector[OffGridHit]): Option[String] =
    Option.when(hits.nonEmpty)(
      s"${hits.size} formula${if hits.size == 1 then "" else "s"} gained #REF! for a reference " +
        "dragged off the grid"
    )

  /**
   * The `OFF_GRID_REF` warning listing the cells (the first ten, then a count), located at the
   * first of them; None when nothing was voided.
   */
  def warning(hits: Vector[OffGridHit]): Option[Warning] =
    hits.headOption.map { first =>
      val listed = hits.take(MaxListed).map(describe).mkString(", ")
      val more = if hits.size > MaxListed then s", … (+${hits.size - MaxListed} more)" else ""
      Warning(
        WarningCode.OFF_GRID_REF,
        s"${hits.size} formula${if hits.size == 1 then "" else "s"} gained #REF! for a reference " +
          s"dragged off the grid: $listed$more; Excel writes it silently, --strict makes this exit 1",
        Some(Location(None, Some(first.sheet), Some(first.cell.toA1), None))
      )
    }
