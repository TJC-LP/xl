package com.tjclp.xl.cf

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.styles.Dxf
import com.tjclp.xl.styles.color.Color

/**
 * What conditional formatting paints on one cell (GH-497): the differential format its true rules
 * compose to (see [[Dxf.orElse]]; a colour scale contributes a solid fill already resolved to RGB)
 * plus at most one data bar. The renderers lay `dxf` over the cell's own style ([[Dxf.applyTo]])
 * and draw strike, which `Font` cannot carry, from `dxf.font`.
 */
final case class CfPaint(dxf: Dxf, bar: Option[CfBar]) derives CanEqual

/**
 * An Excel 2007 data bar: it fills `fraction` of the cell's inner width (Excel's 10%..90% band)
 * with a gradient from `color` to white. `showValue = false` hides the cell's text.
 */
final case class CfBar(fraction: Double, color: Color, showValue: Boolean) derives CanEqual

/**
 * The evaluated conditional formatting of one render window, the input of the renderers' overlay
 * overloads (`sheet.toSvg(range, overlay)`, `sheet.toHtml(range, overlay)`). A cell without an
 * entry is unpainted. xl-core only paints it; the rules are evaluated by xl-evaluator's
 * `CfEvaluator` (`sheet.conditionalFormatOverlay(range)`), since most of them are formulas.
 */
final case class CfOverlay(cells: Map[ARef, CfPaint]) derives CanEqual:
  /** The paint of the cell at `ref`, if any rule painted it. */
  def at(ref: ARef): Option[CfPaint] = cells.get(ref)

  /** Whether no cell is painted. */
  def isEmpty: Boolean = cells.isEmpty

object CfOverlay:
  /** Nothing painted: the renderers draw every cell from its own style alone. */
  val empty: CfOverlay = CfOverlay(Map.empty)
