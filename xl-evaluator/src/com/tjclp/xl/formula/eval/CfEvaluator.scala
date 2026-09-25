package com.tjclp.xl.formula.eval

import scala.math.BigDecimal.RoundingMode

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cf.{
  CfBar,
  CfOperator,
  CfOverlay,
  CfPaint,
  CfPoint,
  CfRule,
  CfTextOp,
  Cfvo,
  ConditionalFormat
}
import com.tjclp.xl.error.XLError
import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.graph.DependencyGraph
import com.tjclp.xl.formula.printer.FormulaShifter
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.Dxf
import com.tjclp.xl.styles.color.{Color, ThemePalette}
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.workbooks.Workbook

/**
 * The conditional formatting of one render window, evaluated (GH-497): the paint the renderers'
 * overlay overloads draw, and every rule that intersects the window but could not be painted, or
 * was painted only in part.
 */
final case class CfEvaluation(overlay: CfOverlay, unevaluated: Vector[CfUnevaluated])
    derives CanEqual

/**
 * A conditional-format rule the evaluation could not paint, or painted only in part (a formula rule
 * that failed at some cells, [[CfUnevaluated.Reason.FailedAt]]): its priority (None for a rule
 * whose priority did not parse), its OOXML kind (`cellIs`, `iconSet`, ... — `block` for a whole
 * block whose range could not be read), the block's ranges, and why.
 */
final case class CfUnevaluated(
  priority: Option[Int],
  kind: String,
  ranges: Vector[CellRange],
  reason: CfUnevaluated.Reason
) derives CanEqual:

  /** One line naming the rule, where it applies and why it was not (or only partly) painted. */
  def message: String =
    val at = priority.fold("")(p => s" (priority $p)")
    val sqref = ranges.map(r => if r.start == r.end then r.start.toA1 else r.toA1)
    val where = if ranges.isEmpty then "" else s" on ${sqref.mkString(" ")}"
    val verb = reason match
      case CfUnevaluated.Reason.FailedAt(_, _, true) => "partly rendered"
      case _ => "not rendered"
    s"conditional format $kind$at$where $verb: ${reason.describe}"

object CfUnevaluated:
  /** Why a rule was not painted, or not everywhere. */
  enum Reason derives CanEqual:
    /** A kind (or formatting) xl does not model for rendering yet: icon sets, x14 data bars, ... */
    case NotModeled

    /**
     * A rule or value-object formula that did not parse, or failed for a reason of the host; or a
     * block cell whose formula a numeric rule needed and xl could not compute. A failure met while
     * evaluating cells names the first one, row-major.
     */
    case FormulaFailed(formula: String, message: String)

    /** A rule no evaluation can paint: a one-operand between, a percentile outside 0..100. */
    case InvalidRule(message: String)

    /**
     * A cell-value, formula or text rule that could not be evaluated for some target cells: the
     * first one, row-major, and why. `evaluatedElsewhere` is whether it evaluated for another cell
     * of the window — then it is painted wherever it did, and only these cells go without it.
     */
    case FailedAt(cell: ARef, message: String, evaluatedElsewhere: Boolean)

    def describe: String = this match
      case NotModeled => "xl does not evaluate this rule (its kind or its formatting) yet"
      case FormulaFailed(formula, message) =>
        s"formula '${quoted(formula, 200)}' failed: ${quoted(message, 400)}"
      case InvalidRule(message) => message
      case FailedAt(cell, message, elsewhere) =>
        val scope =
          if elsewhere then "; it is painted on the cells where it evaluates" else ""
        s"it could not be evaluated at ${cell.toA1}: ${quoted(message, 400)}$scope"

  /** At most `max` characters of `text` (a formula may run to Excel's 8,192), never half a pair. */
  private def quoted(text: String, max: Int): String =
    if text.length <= max then text
    else
      val cut = if Character.isHighSurrogate(text.charAt(max - 1)) then max - 1 else max
      text.take(cut) + "…"

/**
 * Excel's conditional-formatting semantics over a [[Sheet]] (GH-497), for the renderers.
 *
 * Rules are evaluated for every window cell their block covers, in Excel's precedence — priority
 * ascending, then document order (block, rule) — and composed property by property
 * ([[Dxf.orElse]]): the highest-precedence TRUE rule wins each fill, font attribute, border side
 * and number format, non-conflicting properties from every true rule combine, a colour scale's fill
 * competes at its own priority, and `stopIfTrue` on a true rule (with or without a format) stops
 * every lower rule for that cell, scales and bars included. At most one data bar per cell: the
 * highest-precedence one.
 *
 * Formula rules are evaluated as Excel stores them, relative to the top-left of the block's
 * bounding box (MS-XLS refBound): each is parsed once and shifted to every target cell, which is
 * also the current cell (`ROW()` banding works). Like Excel, they evaluate as array formulas (a
 * range in `=OR(A1=$X$1:$X$5)` is every cell of it, not the one in the rule cell's row), an array
 * result reading at its top-left. A cell-value rule is lowered to one comparison formula, so
 * blanks, text, case and errors compare exactly as the evaluator's `<`/`=` do: a blank matches
 * `less than 5`, text sorts above numbers, an error never matches. A text rule is the SEARCH / LEFT
 * / RIGHT formula Excel stores ([[CfTextOp.formula]]). An expression is true for TRUE and for a
 * non-zero number or date.
 *
 * Top-N, colour scales and data bars consider numbers and dates only, over every populated cell of
 * the whole block — every range, hidden rows and cells outside the window included, each once. A
 * formula cell reads its cached value; an uncached one is computed, row-major through a memo of the
 * block's own, so a chain filled down or across within the block computes whatever its length, and
 * what one block cannot compute never reaches another. A block cell xl cannot compute (a chain read
 * from its far end, past the evaluator's recursion guard, among them) leaves the statistics
 * unknown: the block's numeric rules paint nothing and are reported, with the first such cell.
 * Colour-scale colours are resolved to RGB with the workbook's theme (Office without a workbook).
 *
 * A formula rule evaluates each target cell on its own, so an uncached chain it reads is bounded by
 * the recursion guard like any single evaluation.
 *
 * Total: an Excel error value is simply no match. A formula that does not parse paints nothing and
 * is reported once. A formula that fails for a reason of the host (an unknown sheet, no workbook
 * for a cross-sheet reference, the recursion guard) at some target cells paints nothing there,
 * still paints wherever it evaluates, and is reported once with the first failing cell in row-major
 * order — `partly rendered` when it evaluated elsewhere, `not rendered` when nowhere. Randomness is
 * seeded, so a render is a function of the sheet, the workbook and the clock.
 */
object CfEvaluator:

  // The only public members are extensions with no default arguments: the object is
  // wildcard-exported (formulaExports, the scripting prelude).
  extension (sheet: Sheet)
    /**
     * Evaluate the sheet's conditional formatting for the cells of `window`: the paint, and every
     * rule intersecting the window that could not be painted. `workbook` resolves cross-sheet
     * references (and supplies the theme and the 1904 date system); `clock` answers TODAY/NOW.
     */
    def evaluateConditionalFormats(
      window: CellRange,
      workbook: Option[Workbook],
      clock: Clock
    ): CfEvaluation =
      Engine(sheet, workbook, clock).evaluate(window)

    /**
     * The paint of the sheet's conditional formatting over `window`, with no workbook (no
     * cross-sheet references) and the system clock: `sheet.toSvg(range,
     * sheet.conditionalFormatOverlay(range))`.
     */
    def conditionalFormatOverlay(window: CellRange): CfOverlay =
      Engine(sheet, None, Clock.system).evaluate(window).overlay

  /**
   * What the paint of `window` reads on `sheet` ([[Engine.reads]]): `view --eval` evaluates these
   * cells with the window, so the paint comes from the values the picture draws.
   */
  private[eval] def reads(
    sheet: Sheet,
    window: CellRange,
    workbook: Option[Workbook] = None
  ): Vector[CellRange] =
    Engine(sheet, workbook, Clock.system).reads(window)

  /** A rule and where it sits: its precedence key and its block. */
  private final case class Entry(block: Block, ruleIndex: Int, rule: CfRule):
    val priority: Option[Int] = CfRule.priorityOf(rule)
    val key: (Int, Int, Int) = (priority.getOrElse(Int.MaxValue), block.index, ruleIndex)

  /** One Rules block intersecting the window. */
  private final case class Block(index: Int, ranges: Vector[CellRange], rules: Vector[CfRule]):
    /** The top-left of the bounding box of every range: the cell rule formulas are written for. */
    val anchor: ARef = ARef.from0(
      ranges.map(_.start.col.index0).minOption.getOrElse(0),
      ranges.map(_.start.row.index0).minOption.getOrElse(0)
    )

    def covers(ref: ARef): Boolean = ranges.exists(_.contains(ref))

  /** A rule ready to paint. */
  private enum Compiled:
    /** A formula rule (cell value, expression, text): paints `dxf` where the formula is true. */
    case Formula(expr: TExpr[?], text: String, dxf: Option[Dxf], stop: Boolean)

    /** Top/bottom N: paints where the value reaches the threshold (None matches nothing). */
    case Top(threshold: Option[BigDecimal], bottom: Boolean, dxf: Option[Dxf], stop: Boolean)

    /** A colour scale over resolved points (RGB); empty paints nothing. */
    case Scale(points: Vector[(BigDecimal, Int)])

    /** A data bar spanning `lo..hi`. */
    case Bar(lo: BigDecimal, hi: BigDecimal, color: Color, showValue: Boolean)

    /** Paints nothing (reported, or an empty population). */
    case Inert

  /** One cell's composition so far. */
  private final case class Acc(dxf: Dxf, bar: Option[CfBar], stopped: Boolean):
    def matched(rule: Option[Dxf], stop: Boolean): Acc =
      copy(dxf = dxf.orElse(rule.getOrElse(Dxf())), stopped = stop)

  private object Acc:
    val start: Acc = Acc(Dxf(), None, stopped = false)

  /**
   * What a block's numeric rules read: the number (or date serial) of every populated cell it
   * covers, and why the first cell xl could not compute (row-major) failed, if one did.
   */
  private final case class Population(
    numbers: Map[ARef, BigDecimal],
    failure: Option[CfUnevaluated.Reason]
  ):
    /** The numbers ascending, duplicates kept. */
    val sorted: Vector[BigDecimal] = numbers.values.toVector.sorted

  private object Population:
    val empty: Population = Population(Map.empty, None)

  private final class Engine(sheet: Sheet, workbook: Option[Workbook], clock: Clock):
    // Excel evaluates conditional-format formulas as array formulas, not as a plain cell's
    // legacy formula: `=SUM(($A$1:$A$10=A1)*1)>1` and `=OR(A1=$X$1:$X$5)` test every row
    private val evaluator: Evaluator = Evaluator.arrayInstance(Rng.seeded(0L))

    private val theme: ThemePalette = workbook.fold(ThemePalette.office)(_.metadata.theme)
    private val date1904: Boolean = workbook.exists(_.metadata.date1904)

    /** The Rules blocks intersecting `window`, in document order. */
    private def blocksOver(window: CellRange): Vector[Block] =
      sheet.conditionalFormats.zipWithIndex.collect {
        case (ConditionalFormat.Rules(ranges, rules, _), i)
            if ranges.exists(_.intersects(window)) =>
          Block(i, ranges, rules)
      }

    /**
     * The cells of the sheet the paint of `window` reads, as ranges never expanded (a whole-column
     * reference stays one range): every range of a block with a numeric rule, whose statistics span
     * the block; what each formula rule reads at each window cell it covers, the cell itself
     * included; and what each value-object formula reads at the block's anchor. Defined names
     * resolve against the workbook, including aliases and sheet-scoped names. References to other
     * sheets are left out, as is a formula that does not parse.
     */
    def reads(window: CellRange): Vector[CellRange] =
      def read(expr: TExpr[?]): Vector[CellRange] =
        val (cells, ranges) = DependencyGraph.localReads(expr, sheet.name, workbook)
        cells.toVector.map(ref => CellRange(ref, ref)) ++ ranges.map(r => CellRange(r.start, r.end))
      def parsed(formula: String): Option[TExpr[?]] = SheetEvaluator.parseFormula(formula).toOption
      blocksOver(window).flatMap { b =>
        val anchorA1 = b.anchor.toA1
        val statistics = if b.rules.exists(numericRule) then b.ranges else Vector.empty
        val targets = b.ranges.flatMap(_.intersect(window)).flatMap(_.cellsRowMajor).distinct
        val ruleReads = b.rules.flatMap(ruleFormula(_, anchorA1)).flatMap(parsed).flatMap { expr =>
          targets.flatMap { ref =>
            val dc = ref.col.index0 - b.anchor.col.index0
            val dr = ref.row.index0 - b.anchor.row.index0
            read(FormulaShifter.shift(expr, dc, dr))
          }
        }
        val cfvoReads = b.rules.flatMap(cfvoFormulas).flatMap(parsed).flatMap(read)
        statistics ++ ruleReads ++ cfvoReads
      }.distinct

    /** The formula a cell-value, expression or text rule tests, written for `anchorA1`. */
    private def ruleFormula(rule: CfRule, anchorA1: String): Option[String] = rule match
      case CfRule.CellIs(op, f1, f2, _, _, _) => cellIsFormula(op, f1, f2, anchorA1).toOption
      case CfRule.Expression(formula, _, _, _) => Some(formula)
      case CfRule.Text(op, text, _, _, _) => Some(CfTextOp.formula(op, text, anchorA1))
      case _ => None

    /** The formulas of a colour scale's or data bar's value objects. */
    private def cfvoFormulas(rule: CfRule): Vector[String] =
      val cfvos = rule match
        case CfRule.ColorScale(min, mid, max, _) =>
          Vector(Some(min), mid, Some(max)).flatten.map(_.cfvo)
        case CfRule.DataBar(min, max, _, _, _) => Vector(min, max)
        case _ => Vector.empty
      cfvos.collect { case Cfvo.Formula(f) => f }

    def evaluate(window: CellRange): CfEvaluation =
      val indexed = sheet.conditionalFormats.zipWithIndex
      val blocks = blocksOver(window)
      // A block whose envelope did not parse has no known range: reported, never painted
      val unreadable = indexed.collect { case (ConditionalFormat.Preserved(xml), i) =>
        val priority = ConditionalFormat.scanPriorities(xml).minOption
        (priority.getOrElse(Int.MaxValue), i, 0) ->
          CfUnevaluated(priority, "block", Vector.empty, CfUnevaluated.Reason.NotModeled)
      }
      // Each block's statistics, read once, for the blocks with a numeric rule
      val populations = blocks.collect {
        case b if b.rules.exists(numericRule) => b.index -> population(b)
      }.toMap
      val compiled = blocks
        .flatMap(b => b.rules.zipWithIndex.map((rule, i) => Entry(b, i, rule)))
        .sortBy(_.key)
        .map(e => (e, compile(e, populations.getOrElse(e.block.index, Population.empty))))
      val rules = compiled.map { case (e, (rule, _)) => (e, rule) }
      val refused = compiled.collect { case (e, (_, Some(reason))) => e.key -> report(e, reason) }
      val (paints, failures, evaluated) = targets(blocks, window).foldLeft(
        (
          Map.empty[ARef, CfPaint],
          Map.empty[(Int, Int, Int), CfUnevaluated],
          Set.empty[(Int, Int, Int)]
        )
      ) { case ((painted, failed, ran), ref) =>
        val (acc, failedNow, ranNow) = paintCell(ref, rules, populations, failed, ran)
        val next =
          if acc.dxf == Dxf() && acc.bar.isEmpty then painted
          else painted.updated(ref, CfPaint(acc.dxf, acc.bar))
        (next, failedNow, ranNow)
      }
      // a rule that failed for some cells says whether it evaluated for any other
      val settled = failures.map { case (key, report) =>
        report.reason match
          case CfUnevaluated.Reason.FailedAt(cell, message, _) =>
            key -> report.copy(reason =
              CfUnevaluated.Reason.FailedAt(cell, message, evaluated(key))
            )
          case _ => key -> report
      }
      val unevaluated = (unreadable ++ refused ++ settled.toVector).sortBy(_._1).map(_._2)
      CfEvaluation(CfOverlay(paints), unevaluated)

    private def report(e: Entry, reason: CfUnevaluated.Reason): CfUnevaluated =
      CfUnevaluated(e.priority, kindOf(e.rule), e.block.ranges, reason)

    /** The window cells some block covers, row-major, each once. */
    private def targets(blocks: Vector[Block], window: CellRange): Vector[ARef] =
      blocks
        .flatMap(_.ranges.flatMap(_.intersect(window)))
        .flatMap(_.cellsRowMajor)
        .distinct
        .sortBy(ref => (ref.row.index0, ref.col.index0))

    /** Fold every rule covering `ref`, in precedence order, into the cell's paint. */
    private def paintCell(
      ref: ARef,
      rules: Vector[(Entry, Compiled)],
      populations: Map[Int, Population],
      failed: Map[(Int, Int, Int), CfUnevaluated],
      evaluated: Set[(Int, Int, Int)]
    ): (Acc, Map[(Int, Int, Int), CfUnevaluated], Set[(Int, Int, Int)]) =
      rules.foldLeft((Acc.start, failed, evaluated)) { case ((acc, failures, ran), (entry, rule)) =>
        val (nextAcc, nextFailures, ok) =
          paintRule(ref, entry, rule, populations, acc, failures)
        (nextAcc, nextFailures, if ok then ran + entry.key else ran)
      }

    /**
     * One rule's contribution to `ref`'s paint (the body of [[paintCell]]'s fold), and whether a
     * formula rule evaluated for this cell.
     */
    private def paintRule(
      ref: ARef,
      entry: Entry,
      rule: Compiled,
      populations: Map[Int, Population],
      acc: Acc,
      failures: Map[(Int, Int, Int), CfUnevaluated]
    ): (Acc, Map[(Int, Int, Int), CfUnevaluated], Boolean) =
      if acc.stopped || !entry.block.covers(ref) then (acc, failures, false)
      else
        // the number a numeric rule sees here: the block's own reading of the cell
        def value: Option[BigDecimal] =
          populations.get(entry.block.index).flatMap(_.numbers.get(ref))
        rule match
          case Compiled.Formula(expr, text, dxf, stop) =>
            test(expr, text, entry.block.anchor, ref) match
              case Right(true) => (acc.matched(dxf, stop), failures, true)
              case Right(false) => (acc, failures, true)
              case Left(message) =>
                // the first failure of each rule, row-major; the rest of its cells still run
                val first =
                  if failures.contains(entry.key) then failures
                  else
                    val reason = CfUnevaluated.Reason.FailedAt(ref, message, false)
                    failures.updated(entry.key, report(entry, reason))
                (acc, first, false)
          case Compiled.Top(threshold, bottom, dxf, stop) =>
            val hit = (value, threshold) match
              case (Some(v), Some(t)) => if bottom then v <= t else v >= t
              case _ => false
            (if hit then acc.matched(dxf, stop) else acc, failures, false)
          case Compiled.Scale(points) =>
            val painted = value.flatMap(scaleColor(points, _)).fold(acc) { argb =>
              acc.copy(dxf = acc.dxf.orElse(Dxf(fill = Some(Fill.Solid(Color.Rgb(argb))))))
            }
            (painted, failures, false)
          case Compiled.Bar(lo, hi, color, showValue) =>
            val barred =
              if acc.bar.isDefined then acc
              else
                value.fold(acc)(v =>
                  acc.copy(bar = Some(CfBar(barFraction(v, lo, hi), color, showValue)))
                )
            (barred, failures, false)
          case Compiled.Inert => (acc, failures, false)

    // ========== Compilation ==========

    /** The rule ready to paint, and why it was not (reported) when it cannot be. */
    private def compile(
      e: Entry,
      population: Population
    ): (Compiled, Option[CfUnevaluated.Reason]) =
      val anchorA1 = e.block.anchor.toA1
      val pop = population.sorted
      e.rule match
        // statistics missing a cell would paint the right values in the wrong colours
        case rule if numericRule(rule) && population.failure.isDefined =>
          (Compiled.Inert, population.failure)
        case CfRule.CellIs(op, f1, f2, dxf, _, stop) =>
          cellIsFormula(op, f1, f2, anchorA1) match
            case Left(reason) => (Compiled.Inert, Some(reason))
            case Right(formula) => formulaRule(formula, (f1 +: f2.toList).mkString(", "), dxf, stop)
        case CfRule.Expression(formula, dxf, _, stop) =>
          formulaRule(formula, formula, dxf, stop)
        case CfRule.Text(op, text, dxf, _, stop) =>
          val formula = CfTextOp.formula(op, text, anchorA1)
          formulaRule(formula, formula, dxf, stop)
        case CfRule.Top10(rank, percent, bottom, dxf, _, stop) =>
          (
            Compiled.Top(
              topThreshold(pop, rank, percent, bottom),
              bottom,
              dxf,
              stop
            ),
            None
          )
        case CfRule.ColorScale(min, mid, max, _) =>
          if pop.isEmpty then (Compiled.Inert, None)
          else
            val points =
              Vector(Some(min), mid, Some(max)).flatten.map(point(_, pop, e.block.anchor))
            points.collectFirst { case Left(reason) => reason } match
              case Some(reason) => (Compiled.Inert, Some(reason))
              case None => (Compiled.Scale(points.collect { case Right(p) => p }), None)
        case CfRule.DataBar(min, max, color, showValue, _) =>
          if pop.isEmpty then (Compiled.Inert, None)
          else
            (cfvoValue(min, pop, e.block.anchor), cfvoValue(max, pop, e.block.anchor)) match
              case (Right(lo), Right(hi)) => (Compiled.Bar(lo, hi, color, showValue), None)
              case (Left(reason), _) => (Compiled.Inert, Some(reason))
              case (_, Left(reason)) => (Compiled.Inert, Some(reason))
        case _: CfRule.Preserved => (Compiled.Inert, Some(CfUnevaluated.Reason.NotModeled))

    private def formulaRule(
      formula: String,
      display: String,
      dxf: Option[Dxf],
      stop: Boolean
    ): (Compiled, Option[CfUnevaluated.Reason]) =
      SheetEvaluator.parseFormula(formula) match
        case Right(expr) => (Compiled.Formula(expr, display, dxf, stop), None)
        case Left(err) =>
          (Compiled.Inert, Some(CfUnevaluated.Reason.FormulaFailed(display, reasonOf(err))))

    /**
     * A cell-value rule as ONE comparison formula for the anchor `t`: `t<(f1)` and its kin; between
     * is inclusive and indifferent to operand order (`OR(AND(t>=a,t<=b),AND(t>=b,t<=a))`), not
     * between its negation — so an error in the cell or an operand never matches either.
     */
    private def cellIsFormula(
      op: CfOperator,
      f1: String,
      f2: Option[String],
      t: String
    ): Either[CfUnevaluated.Reason, String] =
      val a = s"(${CellValue.canonicalFormulaText(f1)})"
      def between(b: String): String = s"OR(AND($t>=$a,$t<=$b),AND($t>=$b,$t<=$a))"
      op match
        case CfOperator.LessThan => Right(s"$t<$a")
        case CfOperator.LessThanOrEqual => Right(s"$t<=$a")
        case CfOperator.Equal => Right(s"$t=$a")
        case CfOperator.NotEqual => Right(s"$t<>$a")
        case CfOperator.GreaterThanOrEqual => Right(s"$t>=$a")
        case CfOperator.GreaterThan => Right(s"$t>$a")
        case CfOperator.Between | CfOperator.NotBetween =>
          f2.map(f => s"(${CellValue.canonicalFormulaText(f)})") match
            case None =>
              Left(CfUnevaluated.Reason.InvalidRule("a between rule needs two operands"))
            case Some(b) =>
              Right(if op == CfOperator.Between then between(b) else s"NOT(${between(b)})")

    /**
     * The threshold a top/bottom-N rule paints from: the k-th largest (smallest) value, duplicates
     * counted, so ties at the threshold all paint. A percent rule takes at least one value — by
     * assumption, as Excel highlights the top cell of a short range for "top 10%" — and k never
     * exceeds the population. A rank below 1 or an empty population paints nothing.
     */
    private def topThreshold(
      pop: Vector[BigDecimal],
      rank: Int,
      percent: Boolean,
      bottom: Boolean
    ): Option[BigDecimal] =
      val n = pop.size
      if rank <= 0 || n == 0 then None
      else
        val wanted = if percent then math.max(1L, n.toLong * rank / 100) else rank.toLong
        val k = math.min(wanted, n.toLong).toInt
        if bottom then pop.lift(k - 1) else pop.lift(n - k)

    /** A colour-scale point: its position over the population and its colour in RGB. */
    private def point(
      p: CfPoint,
      pop: Vector[BigDecimal],
      anchor: ARef
    ): Either[CfUnevaluated.Reason, (BigDecimal, Int)] =
      cfvoValue(p.cfvo, pop, anchor).map(v => (v, opaque(p.color.toResolvedArgb(theme))))

    /**
     * Where a value object sits over a non-empty population (sorted ascending): the extremes, a
     * number, a percent of the span, a PERCENTILE.INC, or a formula evaluated once at the anchor.
     */
    private def cfvoValue(
      cfvo: Cfvo,
      pop: Vector[BigDecimal],
      anchor: ARef
    ): Either[CfUnevaluated.Reason, BigDecimal] =
      val lo = pop.headOption.getOrElse(BigDecimal(0))
      val hi = pop.lastOption.getOrElse(BigDecimal(0))
      def inRange(kind: String, p: BigDecimal)(
        value: => BigDecimal
      ): Either[CfUnevaluated.Reason, BigDecimal] =
        if p < 0 || p > 100 then
          Left(CfUnevaluated.Reason.InvalidRule(s"a $kind of $p is outside 0..100"))
        else Right(value)
      cfvo match
        case Cfvo.Min => Right(lo)
        case Cfvo.Max => Right(hi)
        case Cfvo.Num(v) => Right(v)
        case Cfvo.Percent(p) => inRange("percent", p)(lo + (hi - lo) * p / 100)
        case Cfvo.Percentile(p) => inRange("percentile", p)(percentile(pop, p))
        case Cfvo.Formula(f) =>
          SheetEvaluator
            .parseFormula(f)
            .flatMap(expr =>
              SheetEvaluator
                .evaluateParsedWith(sheet, f, expr, evaluator, clock, workbook, Some(anchor))
            ) match
            case Left(err) => Left(CfUnevaluated.Reason.FormulaFailed(f, reasonOf(err)))
            case Right(v) =>
              numeric(v).toRight(
                CfUnevaluated.Reason.FormulaFailed(f, s"returned $v, not a number")
              )

    /** PERCENTILE.INC over a sorted, non-empty population. */
    private def percentile(pop: Vector[BigDecimal], p: BigDecimal): BigDecimal =
      val h = BigDecimal(pop.size - 1) * p / 100
      val below = h.setScale(0, RoundingMode.FLOOR)
      val i = below.toInt
      val x = pop.lift(i).getOrElse(BigDecimal(0))
      x + (h - below) * (pop.lift(i + 1).getOrElse(x) - x)

    // ========== Evaluation ==========

    /**
     * The formula `expr`, written for `anchor`, evaluated for `ref`: true, false, or a host
     * failure.
     */
    private def test(
      expr: TExpr[?],
      text: String,
      anchor: ARef,
      ref: ARef
    ): Either[String, Boolean] =
      val dc = ref.col.index0 - anchor.col.index0
      val dr = ref.row.index0 - anchor.row.index0
      SheetEvaluator
        .evaluateParsedWith(
          sheet,
          text,
          FormulaShifter.shift(expr, dc, dr),
          evaluator,
          clock,
          workbook,
          Some(ref)
        )
        .fold(err => Left(reasonOf(err)), v => Right(truthy(v)))

    /** Excel's truth for a rule result: TRUE, a non-zero number, a non-zero date. */
    private def truthy(value: CellValue): Boolean = value match
      case CellValue.Bool(b) => b
      case CellValue.Formula(_, Some(cached), _) => truthy(cached)
      case other => numeric(other).exists(_.signum != 0)

    /** A number or a date's serial; everything else (text, logicals, blanks, errors) is not. */
    private def numeric(value: CellValue): Option[BigDecimal] = value match
      case CellValue.Number(n) => Some(n)
      case CellValue.DateTime(dt) => Some(BigDecimal(CellValue.dateTimeToExcelSerial(dt, date1904)))
      case _ => None

    /**
     * The number a numeric rule (top-N, scale, bar) sees in the cell at `ref`, read through `read`
     * — a number or a date's serial; an Excel error, like text, is none — or why xl could not
     * compute it.
     */
    private def numberAt(
      read: ARef => Either[EvalError, CellValue],
      ref: ARef
    ): Either[CfUnevaluated.Reason, Option[BigDecimal]] =
      read(ref) match
        case Right(value) => Right(numeric(value))
        case Left(err) =>
          val formula = sheet(ref).value match
            case CellValue.Formula(expression, _, _) => expression
            case _ => ref.toA1
          // the reader words a parse failure as a cross-sheet one: parse again for the plain words
          val why = SheetEvaluator
            .parseFormula(formula)
            .fold(reasonOf, _ => reasonOf(EvalError.toXLError(err, Some(formula))))
          Left(CfUnevaluated.Reason.FormulaFailed(formula, s"$why (cell ${ref.toA1})"))

    /**
     * Every populated cell the block covers, read once: statistics span the whole block, not the
     * window. Visits the sheet's cells, never a range's, so a whole-column block is cheap; reads
     * them row-major through one memo, so a chain filled down or across finds each precedent
     * computed instead of recursing through it.
     *
     * The memo and the seed are the block's own: the memo keeps failures (a depth-guard failure
     * reached from another block's cell would otherwise stop this one) and the draws advance, so
     * sharing either would make a block's paint depend on the blocks read before it.
     */
    private def population(block: Block): Population =
      val read =
        Evaluator.cellValueReader(
          sheet,
          clock,
          workbook,
          0,
          Rng.seeded(0L),
          new Evaluator.EvalMemo,
          None,
          None
        )
      val readings = sheet.cells.keysIterator
        .filter(block.covers)
        .toVector
        .sortBy(ref => (ref.row.index0, ref.col.index0))
        .map(ref => ref -> numberAt(read, ref))
      Population(
        readings.collect { case (ref, Right(Some(v))) => ref -> v }.toMap,
        readings.collectFirst { case (_, Left(reason)) => reason }
      )

    /**
     * A colour scale's colour for `v`: the first colour at or below the first point, the last at or
     * above the last point, else interpolated in the bracketing segment (a 3-point scale's middle
     * is clamped between its ends). All-equal values therefore take the first colour.
     */
    private def scaleColor(points: Vector[(BigDecimal, Int)], v: BigDecimal): Option[Int] =
      (points.headOption, points.lastOption) match
        case (Some((lo, loC)), Some((hi, hiC))) =>
          if v <= lo then Some(loC)
          else if v >= hi then Some(hiC)
          else
            points.lift(1).filter(_ => points.size == 3) match
              case None => Some(lerp(loC, hiC, fraction(v, lo, hi)))
              case Some((m0, midC)) =>
                val mid = m0.max(lo).min(hi)
                if v <= mid then Some(lerp(loC, midC, fraction(v, lo, mid)))
                else Some(lerp(midC, hiC, fraction(v, mid, hi)))
        case _ => None

    /** Where `v` sits in `lo..hi`, as 0..1; callers guarantee `lo < v < hi` or `lo <= v <= hi`. */
    private def fraction(v: BigDecimal, lo: BigDecimal, hi: BigDecimal): Double =
      if hi <= lo then 0.0 else ((v - lo) / (hi - lo)).toDouble

    /** Excel 2007's bar length: 10% of the cell at the minimum, 90% at the maximum. */
    private def barFraction(v: BigDecimal, lo: BigDecimal, hi: BigDecimal): Double =
      val t = if v <= lo then 0.0 else if v >= hi then 1.0 else fraction(v, lo, hi)
      0.1 + 0.8 * t

  // ========== Shared helpers ==========

  /** Per-channel linear interpolation between two ARGB colours, half-up, opaque. */
  private def lerp(from: Int, to: Int, t: Double): Int =
    def channel(shift: Int): Int =
      val a = (from >> shift) & 0xff
      val b = (to >> shift) & 0xff
      math.round(a + (b - a) * t).toInt & 0xff
    0xff000000 | (channel(16) << 16) | (channel(8) << 8) | channel(0)

  private def opaque(argb: Int): Int = 0xff000000 | (argb & 0xffffff)

  /** The host failure's reason, without the envelope `XLError.message` puts around it. */
  private def reasonOf(err: XLError): String = err match
    case XLError.FormulaError(_, reason) => reason
    case other => other.message

  /** The `type` of a payload's `<cfRule>` tag (a payload may open with an XML declaration). */
  private val typeToken = raw"""<(?:[\w.-]+:)?cfRule\b[^>]*?\btype="([^"]*)"""".r

  /** A rule that reads its block's statistics. */
  private def numericRule(rule: CfRule): Boolean = rule match
    case _: CfRule.Top10 | _: CfRule.ColorScale | _: CfRule.DataBar => true
    case _ => false

  /** The OOXML kind of a rule; a Preserved rule's is read off its payload's opening tag. */
  private def kindOf(rule: CfRule): String = rule match
    case _: CfRule.CellIs => "cellIs"
    case _: CfRule.Expression => "expression"
    case _: CfRule.ColorScale => "colorScale"
    case _: CfRule.DataBar => "dataBar"
    case _: CfRule.Top10 => "top10"
    case CfRule.Text(op, _, _, _, _) =>
      op match
        case CfTextOp.Contains => "containsText"
        case CfTextOp.NotContains => "notContainsText"
        case CfTextOp.BeginsWith => "beginsWith"
        case CfTextOp.EndsWith => "endsWith"
    case CfRule.Preserved(xml, _, _) =>
      typeToken.findFirstMatchIn(xml).map(_.group(1)).getOrElse("unknown")
