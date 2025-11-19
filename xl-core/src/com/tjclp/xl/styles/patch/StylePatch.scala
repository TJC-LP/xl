package com.tjclp.xl.styles.patch

import cats.Monoid
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.alignment.Align
import com.tjclp.xl.styles.border.Border
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Patch operations for CellStyle with Monoid semantics.
 *
 * Allows declarative composition of styles updates.
 */

/** Patch operations for CellStyle with Monoid semantics */
enum StylePatch:
  case SetFont(font: Font)
  case SetFill(fill: Fill)
  case SetBorder(border: Border)
  case SetNumFmt(numFmt: NumFmt)
  case SetAlign(align: Align)
  case Batch(patches: Vector[StylePatch])

object StylePatch:
  val empty: StylePatch = Batch(Vector.empty)

  def combine(p1: StylePatch, p2: StylePatch): StylePatch = (p1, p2) match
    case (Batch(ps1), Batch(ps2)) => Batch(ps1 ++ ps2)
    case (Batch(ps1), p2) => Batch(ps1 :+ p2)
    case (p1, Batch(ps2)) => Batch(p1 +: ps2)
    case (p1, p2) => Batch(Vector(p1, p2))

  given Monoid[StylePatch] with
    def empty: StylePatch = StylePatch.empty
    def combine(x: StylePatch, y: StylePatch): StylePatch = StylePatch.combine(x, y)

  /** Apply a styles patch to create a new styles */
  def applyPatch(style: CellStyle, patch: StylePatch): CellStyle = patch match
    case SetFont(font) => style.withFont(font)
    case SetFill(fill) => style.withFill(fill)
    case SetBorder(border) => style.withBorder(border)
    case SetNumFmt(numFmt) => style.withNumFmt(numFmt)
    case SetAlign(align) => style.withAlign(align)
    case Batch(patches) =>
      patches.foldLeft(style)((s, p) => applyPatch(s, p))

  /** Apply multiple patches in sequence */
  def applyPatches(style: CellStyle, patches: Iterable[StylePatch]): CellStyle =
    applyPatch(style, Batch(patches.toVector))

  // ========== Composition Operator (No Type Ascription Required) ==========

  extension (p1: StylePatch)
    /**
     * Compose two StylePatches without requiring Cats Monoid syntax.
     *
     * Eliminates the need for type ascription on enum cases, enabling clean composition:
     * {{{
     *   val patch = StylePatch.SetFont(font) ++ StylePatch.SetFill(fill)
     *   // No need for: (StylePatch.SetFont(font): StylePatch) |+| ...
     * }}}
     */
    infix def ++(p2: StylePatch): StylePatch = StylePatch.combine(p1, p2)

  extension (style: CellStyle)
    @annotation.targetName("applyStylePatchExt")
    def applyPatch(patch: StylePatch): CellStyle =
      StylePatch.applyPatch(style, patch)

    @annotation.targetName("applyStylePatchesExt")
    def applyPatches(patches: StylePatch*): CellStyle =
      StylePatch.applyPatches(style, patches)
