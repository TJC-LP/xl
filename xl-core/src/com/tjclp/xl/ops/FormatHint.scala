package com.tjclp.xl.ops

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Where a value's number format came from (ADR-017 invariant 4, GH-560). Only [[Explicit]] — the
 * agent asked for it — may override a non-General numFmt; [[Inferred]] — a codec's hint for a
 * `LocalDate`, a detected currency string — keeps `Sheet.put`'s General-only merge rule on every
 * surface.
 */
enum FormatHint derives CanEqual:
  /** Codec inference: applied only when the existing numFmt is General. */
  case Inferred(fmt: NumFmt)

  /** The caller asked for it: replaces the numFmt, keeping font, fill, border and alignment. */
  case Explicit(fmt: NumFmt)

  /** The format the hint carries, whichever way it arose. */
  def numFmt: NumFmt = this match
    case Inferred(f) => f
    case Explicit(f) => f

  /** The style the hint yields on top of `existing` (the cell's current style, or the default). */
  def resolve(existing: CellStyle): CellStyle = this match
    case Explicit(f) => existing.withNumFmt(f)
    case Inferred(f) =>
      if existing.numFmt == NumFmt.General && f != NumFmt.General then existing.withNumFmt(f)
      else existing

/**
 * The scope an edit sequence runs under: the default sheet for unqualified targets (`-s` at the
 * CLI, `Scope.of(sheet)` for `sheet.edit`). Rule 4 (ADR-017 §2.6): a rename of the default sheet
 * retargets the edits that follow it.
 */
final case class Scope(defaultSheet: Option[SheetName]) derives CanEqual:
  /** The scope after `edit` has run. */
  def after(edit: Edit): Scope = edit match
    case Edit.RenameSheet(from, to) if defaultSheet.contains(from) => Scope(Some(to))
    case _ => this

object Scope:
  /** No default: every unqualified target needs a single-sheet book or fails `SheetRequired`. */
  val none: Scope = Scope(None)

  def of(sheet: SheetName): Scope = Scope(Some(sheet))

/** What [[Edit.Clear]] removes. At least one flag must be set. */
final case class ClearWhat(contents: Boolean, styles: Boolean, comments: Boolean) derives CanEqual:
  def isEmpty: Boolean = !contents && !styles && !comments

object ClearWhat:
  val contents: ClearWhat = ClearWhat(contents = true, styles = false, comments = false)
  val styles: ClearWhat = ClearWhat(contents = false, styles = true, comments = false)
  val comments: ClearWhat = ClearWhat(contents = false, styles = false, comments = true)
  val all: ClearWhat = ClearWhat(contents = true, styles = true, comments = true)

  /**
   * The CLI flag rule: `--all` clears everything; no flag clears contents; else the named parts.
   */
  def fromFlags(all: Boolean, styles: Boolean, comments: Boolean): ClearWhat =
    ClearWhat(
      contents = all || (!styles && !comments),
      styles = all || styles,
      comments = all || comments
    )

/** How [[Edit.Style]] combines with the cells' current styles. */
enum StyleMode derives CanEqual:
  /** Overlay onto each cell's existing style (the CLI default). */
  case Merge

  /** Replace the whole style: the overlay applied to `CellStyle.default`. */
  case Replace
