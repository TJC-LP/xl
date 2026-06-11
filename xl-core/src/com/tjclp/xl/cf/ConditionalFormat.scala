package com.tjclp.xl.cf

import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.styles.Dxf
import com.tjclp.xl.styles.color.Color

/**
 * Conditional formatting (GH-136): one worksheet `<conditionalFormatting>` block.
 *
 * Document order in `Sheet.conditionalFormats` is emission order. The envelope (sqref ranges +
 * pivot flag) is ALWAYS typed when parseable — ranges shift under structural edits even when
 * individual rules ride through as [[CfRule.Preserved]].
 */
enum ConditionalFormat derives CanEqual:
  /**
   * One block with a parsed envelope: the rules apply to every range in `ranges`.
   *
   * A block with empty `ranges` or `rules` is unexpressible in OOXML and is dropped at emission
   * (the authoring API cannot construct one).
   */
  case Rules(ranges: Vector[CellRange], rules: Vector[CfRule], pivot: Boolean = false)

  /**
   * Whole-block fallback ONLY for an unparseable/unknown envelope (sqref that fails to parse,
   * envelope attrs outside {sqref, pivot}). CONTRACT = [[com.tjclp.xl.drawings.Drawing.Preserved]]:
   * constructed only by xl-ooxml's reader; the payload is the scope-self-contained canonical XML of
   * one whole element, re-emitted verbatim. Users must not construct this case; a hand-built
   * payload that is not canonical XML is silently dropped at emission.
   */
  case Preserved(xml: String)

object ConditionalFormat:
  /**
   * Priorities regex-scanned from a block-Preserved payload — the rare corrupt-envelope case where
   * rule priorities are not individually parsed. A documented wart: pure text scan, no XML parse;
   * values that overflow Int are skipped.
   */
  def scanPriorities(xml: String): Vector[Int] =
    raw"""priority="(\d+)"""".r.findAllMatchIn(xml).flatMap(_.group(1).toIntOption).toVector

/** Comparison operator for [[CfRule.CellIs]] (OOXML ST_ConditionalFormattingOperator subset). */
enum CfOperator derives CanEqual:
  case LessThan, LessThanOrEqual, Equal, NotEqual,
    GreaterThanOrEqual, GreaterThan, Between, NotBetween

/** Text-match operator for [[CfRule.Text]] rules. */
enum CfTextOp derives CanEqual:
  case Contains, NotContains, BeginsWith, EndsWith

/**
 * Conditional format value object (`<cfvo>`): an axis point for color scales and data bars. Formula
 * text is stored WITHOUT a leading '='.
 */
enum Cfvo derives CanEqual:
  case Min, Max
  case Num(value: BigDecimal)
  case Percent(value: BigDecimal)
  case Percentile(value: BigDecimal)
  case Formula(formula: String)

/** One color-scale interpolation point: a [[Cfvo]] position plus its color. */
final case class CfPoint(cfvo: Cfvo, color: Color) derives CanEqual

/**
 * One typed `<cfRule>` (GH-136).
 *
 * Formula semantics (CellIs/Expression): Excel evaluates rule formulas as if entered for the
 * range's top-left cell, relative references adjusting per cell (fill semantics). XL stores the
 * text verbatim — without the leading '=' (the OOXML `<formula>` child has none) — and does not
 * rewrite anchors at author time. Sheet-qualified references are the author's literal text.
 *
 * Lower `priority` wins in Excel; `CfRule.AutoPriority` (0) means "assign at append" — see
 * `Sheet.conditionalFormat`.
 */
enum CfRule derives CanEqual:
  /** Cell-value comparison (`type="cellIs"`); `formula2` only for Between/NotBetween. */
  case CellIs(
    op: CfOperator,
    formula1: String,
    formula2: Option[String],
    dxf: Option[Dxf],
    priority: Int,
    stopIfTrue: Boolean = false
  )

  /** Custom formula rule (`type="expression"`). */
  case Expression(formula: String, dxf: Option[Dxf], priority: Int, stopIfTrue: Boolean = false)

  /** 2- or 3-point color scale; the optional `mid` makes illegal point counts unrepresentable. */
  case ColorScale(min: CfPoint, mid: Option[CfPoint], max: CfPoint, priority: Int)

  /** Data bar with inline color (no dxf — Excel stores data-bar colors inline). */
  case DataBar(min: Cfvo, max: Cfvo, color: Color, showValue: Boolean = true, priority: Int)

  /** Top/bottom N (or N percent) rule (`type="top10"`). */
  case Top10(
    rank: Int,
    percent: Boolean,
    bottom: Boolean,
    dxf: Option[Dxf],
    priority: Int,
    stopIfTrue: Boolean = false
  )

  /**
   * Text-match rule (containsText / notContainsText / beginsWith / endsWith). The OOXML `<formula>`
   * child is DERIVED at emission from `text` and the block's first range — never stored — so it
   * stays correct under structural shifts.
   */
  case Text(
    op: CfTextOp,
    text: String,
    dxf: Option[Dxf],
    priority: Int,
    stopIfTrue: Boolean = false
  )

  /**
   * Unmodeled `<cfRule>` (iconSet, timePeriod, aboveAverage, duplicate/uniqueValues,
   * containsBlanks/Errors, any rule with a child extLst, any rule whose dxf is outside the modeled
   * subset, any unknown attr/child). Reader-constructed only; canonical XML re-emitted verbatim.
   * `priority` is parsed read-only and feeds the allocator's max — never re-stamped.
   */
  case Preserved(xml: String, priority: Option[Int])

object CfRule:
  /** Documented sentinel: "assign at append" (see `Sheet.conditionalFormat`). */
  val AutoPriority: Int = 0

  private def strip(formula: String): String =
    if formula.startsWith("=") then formula.drop(1) else formula

  /**
   * Cell-value rule for non-Between operators. For Between/NotBetween use [[between]] /
   * [[notBetween]] (this constructor leaves formula2 empty, which emits a degenerate single-operand
   * rule).
   */
  def cellIs(op: CfOperator, formula: String, dxf: Dxf): CfRule =
    CellIs(op, strip(formula), None, Some(dxf), AutoPriority)

  /** Cell-value Between rule (inclusive bounds). */
  def between(lo: String, hi: String, dxf: Dxf): CfRule =
    CellIs(CfOperator.Between, strip(lo), Some(strip(hi)), Some(dxf), AutoPriority)

  /** Cell-value NotBetween rule. */
  def notBetween(lo: String, hi: String, dxf: Dxf): CfRule =
    CellIs(CfOperator.NotBetween, strip(lo), Some(strip(hi)), Some(dxf), AutoPriority)

  /** Custom formula rule. */
  def expression(formula: String, dxf: Dxf): CfRule =
    Expression(strip(formula), Some(dxf), AutoPriority)

  /** Two-point color scale. */
  def colorScale2(min: CfPoint, max: CfPoint): CfRule =
    ColorScale(min, None, max, AutoPriority)

  /** Three-point color scale. */
  def colorScale3(min: CfPoint, mid: CfPoint, max: CfPoint): CfRule =
    ColorScale(min, Some(mid), max, AutoPriority)

  /** Data bar over the full value range by default. */
  def dataBar(color: Color, min: Cfvo = Cfvo.Min, max: Cfvo = Cfvo.Max): CfRule =
    DataBar(min, max, color, showValue = true, AutoPriority)

  /** Top/bottom-N rule. */
  def top10(rank: Int, dxf: Dxf, percent: Boolean = false, bottom: Boolean = false): CfRule =
    Top10(rank, percent, bottom, Some(dxf), AutoPriority)

  /** Highlight cells containing `text`. */
  def containsText(text: String, dxf: Dxf): CfRule =
    Text(CfTextOp.Contains, text, Some(dxf), AutoPriority)

  /** Highlight cells NOT containing `text`. */
  def notContainsText(text: String, dxf: Dxf): CfRule =
    Text(CfTextOp.NotContains, text, Some(dxf), AutoPriority)

  /** Highlight cells beginning with `text`. */
  def beginsWith(text: String, dxf: Dxf): CfRule =
    Text(CfTextOp.BeginsWith, text, Some(dxf), AutoPriority)

  /** Highlight cells ending with `text`. */
  def endsWith(text: String, dxf: Dxf): CfRule =
    Text(CfTextOp.EndsWith, text, Some(dxf), AutoPriority)

  /** The rule's priority; `None` for a Preserved rule whose priority did not parse. */
  def priorityOf(rule: CfRule): Option[Int] = rule match
    case r: CellIs => Some(r.priority)
    case r: Expression => Some(r.priority)
    case r: ColorScale => Some(r.priority)
    case r: DataBar => Some(r.priority)
    case r: Top10 => Some(r.priority)
    case r: Text => Some(r.priority)
    case Preserved(_, p) => p

  /** Restamp a typed rule's priority; Preserved rules are NEVER renumbered (identity). */
  def withPriority(rule: CfRule, priority: Int): CfRule = rule match
    case r: CellIs => r.copy(priority = priority)
    case r: Expression => r.copy(priority = priority)
    case r: ColorScale => r.copy(priority = priority)
    case r: DataBar => r.copy(priority = priority)
    case r: Top10 => r.copy(priority = priority)
    case r: Text => r.copy(priority = priority)
    case p: Preserved => p
