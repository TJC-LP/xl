package com.tjclp.xl.cli.helpers

import com.tjclp.xl.cf.{CfOperator, CfPoint, CfRule, CfTextOp, Cfvo}
import com.tjclp.xl.cli.ColorParser
import com.tjclp.xl.styles.{Dxf, DxfFont}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill

/**
 * Parser for the `cf add` colon rule DSL (GH-324), mapping rule strings to the xl-core CfRule
 * builders. All rules are built with [[CfRule.AutoPriority]] — priorities are stamped by
 * `Sheet.conditionalFormat` at append, never by the CLI.
 *
 * Families:
 *   - `cellIs:<op>:<value>` — op: lessThan/lt, lessThanOrEqual/lte, equal/eq, notEqual/ne,
 *     greaterThanOrEqual/gte, greaterThan/gt
 *   - `between:<lo>:<hi>` / `notBetween:<lo>:<hi>` — inclusive bounds
 *   - `expression:<formula>` — the formula may itself contain ':' (ranges)
 *   - `colorScale:<c1>:<c2>[:<c3>]` — 2- or 3-point scale (3-point mid at the 50th percentile)
 *   - `dataBar:<color>`
 *   - `top10:<n>[:percent]` / `bottom10:<n>[:percent]`
 *   - `text:<op>:<s>` — op: contains, notContains, beginsWith, endsWith; `<s>` may contain ':'
 *
 * Color tokens inside colorScale/dataBar accept named, #hex, and rgb(r,g,b) colors; `theme:` colors
 * are not representable there (the ':' separator conflicts) — dxf flag colors (`--bg`/`--fg`) have
 * no such restriction.
 */
object CfRuleParser:

  /** Rule families whose format comes from a dxf (highlight-style rules). */
  private val dxfFamilies =
    Set("cellis", "between", "notbetween", "expression", "top10", "bottom10", "text")

  private val operators: Map[String, CfOperator] = Map(
    "lessthan" -> CfOperator.LessThan,
    "lt" -> CfOperator.LessThan,
    "lessthanorequal" -> CfOperator.LessThanOrEqual,
    "lte" -> CfOperator.LessThanOrEqual,
    "equal" -> CfOperator.Equal,
    "eq" -> CfOperator.Equal,
    "notequal" -> CfOperator.NotEqual,
    "ne" -> CfOperator.NotEqual,
    "greaterthanorequal" -> CfOperator.GreaterThanOrEqual,
    "gte" -> CfOperator.GreaterThanOrEqual,
    "greaterthan" -> CfOperator.GreaterThan,
    "gt" -> CfOperator.GreaterThan
  )

  private val textOps: Map[String, CfTextOp] = Map(
    "contains" -> CfTextOp.Contains,
    "notcontains" -> CfTextOp.NotContains,
    "beginswith" -> CfTextOp.BeginsWith,
    "startswith" -> CfTextOp.BeginsWith,
    "endswith" -> CfTextOp.EndsWith
  )

  /**
   * Build a Dxf from the cf command's formatting flags. Boolean flag presence means force-on
   * (`Some(true)`) — DxfFont is tri-state and absent flags inherit from the base cell style. Colors
   * accept the full ColorParser syntax (named, #hex, rgb, theme:slot[:tint]). Returns None when no
   * flag is set.
   */
  def buildDxf(
    bold: Boolean,
    italic: Boolean,
    underline: Boolean,
    strike: Boolean,
    bg: Option[String],
    fg: Option[String]
  ): Either[String, Option[Dxf]] =
    for
      bgColor <- parseColorOpt(bg)
      fgColor <- parseColorOpt(fg)
    yield
      val font =
        if bold || italic || underline || strike || fgColor.isDefined then
          Some(
            DxfFont(
              bold = if bold then Some(true) else None,
              italic = if italic then Some(true) else None,
              strike = if strike then Some(true) else None,
              underline = if underline then Some(true) else None,
              color = fgColor
            )
          )
        else None
      val fill = bgColor.map(c => Fill.Solid(c))
      if font.isEmpty && fill.isEmpty then None else Some(Dxf(font = font, fill = fill))

  private def parseColorOpt(s: Option[String]): Either[String, Option[Color]] =
    s match
      case None => Right(None)
      case Some(str) => ColorParser.parse(str).map(Some(_))

  /**
   * Parse a rule string into a typed CfRule. Highlight families require a dxf (at least one
   * formatting flag); colorScale/dataBar carry inline colors and reject dxf flags.
   */
  def parse(rule: String, dxf: Option[Dxf]): Either[String, CfRule] =
    val family = rule.takeWhile(_ != ':').toLowerCase
    if dxfFamilies.contains(family) then
      dxf match
        case None =>
          Left(
            s"$family rules need a format: add at least one of --bold, --italic, --underline, " +
              "--strike, --bg <color>, --fg <color>"
          )
        case Some(d) => parseFamily(rule, family, d)
    else if family == "colorscale" || family == "databar" then
      dxf match
        case Some(_) =>
          Left(
            s"$family rules carry inline colors; formatting flags (--bold/--bg/--fg/...) " +
              "have no effect — remove them"
          )
        case None => parseFamily(rule, family, Dxf())
    else
      Left(
        s"Unknown rule family: '${rule.takeWhile(_ != ':')}'. Valid: cellIs, between, notBetween, " +
          "expression, colorScale, dataBar, top10, bottom10, text"
      )

  @SuppressWarnings(Array("org.wartremover.warts.SeqApply"))
  private def parseFamily(rule: String, family: String, dxf: Dxf): Either[String, CfRule] =
    family match
      case "cellis" =>
        rule.split(":", 3).toList match
          case _ :: opStr :: value :: Nil if value.nonEmpty =>
            operators
              .get(opStr.toLowerCase)
              .toRight(
                s"Unknown cellIs operator: '$opStr'. " +
                  "Valid: lessThan, lessThanOrEqual, equal, notEqual, greaterThanOrEqual, " +
                  "greaterThan (or lt, lte, eq, ne, gte, gt)"
              )
              .map(op => CfRule.cellIs(op, value, dxf))
          case _ =>
            Left("cellIs rule format: cellIs:<operator>:<value>, e.g. cellIs:greaterThan:100")

      case "between" | "notbetween" =>
        rule.split(":", 3).toList match
          case _ :: lo :: hi :: Nil if lo.nonEmpty && hi.nonEmpty =>
            if family == "between" then Right(CfRule.between(lo, hi, dxf))
            else Right(CfRule.notBetween(lo, hi, dxf))
          case _ =>
            Left(s"$family rule format: $family:<lo>:<hi>, e.g. between:10:100")

      case "expression" =>
        rule.split(":", 2).toList match
          case _ :: formula :: Nil if formula.nonEmpty =>
            Right(CfRule.expression(formula, dxf))
          case _ =>
            Left("expression rule format: expression:<formula>, e.g. expression:MOD(ROW(),2)=0")

      case "colorscale" =>
        val tokens = rule.split(":").toList.drop(1)
        if tokens.exists(_.equalsIgnoreCase("theme")) then
          Left(
            "theme colors are not supported inside colorScale/dataBar rule strings " +
              "(the ':' separator conflicts); use named, #hex, or rgb(r,g,b) colors"
          )
        else
          tokens match
            case c1 :: c2 :: Nil =>
              for
                min <- ColorParser.parse(c1)
                max <- ColorParser.parse(c2)
              yield CfRule.colorScale2(CfPoint(Cfvo.Min, min), CfPoint(Cfvo.Max, max))
            case c1 :: c2 :: c3 :: Nil =>
              for
                min <- ColorParser.parse(c1)
                mid <- ColorParser.parse(c2)
                max <- ColorParser.parse(c3)
              yield CfRule.colorScale3(
                CfPoint(Cfvo.Min, min),
                CfPoint(Cfvo.Percentile(BigDecimal(50)), mid),
                CfPoint(Cfvo.Max, max)
              )
            case _ =>
              Left(
                "colorScale rule format: colorScale:<c1>:<c2>[:<c3>], e.g. colorScale:red:white:green"
              )

      case "databar" =>
        rule.split(":", 2).toList match
          case _ :: colorStr :: Nil if colorStr.nonEmpty =>
            if colorStr.toLowerCase.startsWith("theme") then
              Left(
                "theme colors are not supported inside colorScale/dataBar rule strings " +
                  "(the ':' separator conflicts); use named, #hex, or rgb(r,g,b) colors"
              )
            else ColorParser.parse(colorStr).map(c => CfRule.dataBar(c))
          case _ =>
            Left("dataBar rule format: dataBar:<color>, e.g. dataBar:#638EC6")

      case "top10" | "bottom10" =>
        val bottom = family == "bottom10"
        rule.split(":").toList.drop(1) match
          case nStr :: rest if rest.isEmpty || rest == List("percent") =>
            nStr.toIntOption.filter(_ >= 1) match
              case Some(n) =>
                Right(CfRule.top10(n, dxf, percent = rest.nonEmpty, bottom = bottom))
              case None => Left(s"$family rank must be a positive integer, got: '$nStr'")
          case _ =>
            Left(s"$family rule format: $family:<n>[:percent], e.g. top10:5 or top10:10:percent")

      case "text" =>
        rule.split(":", 3).toList match
          case _ :: opStr :: text :: Nil if text.nonEmpty =>
            textOps
              .get(opStr.toLowerCase)
              .toRight(
                s"Unknown text operator: '$opStr'. Valid: contains, notContains, beginsWith, endsWith"
              )
              .map {
                case CfTextOp.Contains => CfRule.containsText(text, dxf)
                case CfTextOp.NotContains => CfRule.notContainsText(text, dxf)
                case CfTextOp.BeginsWith => CfRule.beginsWith(text, dxf)
                case CfTextOp.EndsWith => CfRule.endsWith(text, dxf)
              }
          case _ =>
            Left("text rule format: text:<operator>:<text>, e.g. text:contains:overdue")

      case other =>
        Left(s"Unknown rule family: '$other'") // unreachable — parse() gates families

  /** Short human-readable description of a rule (for command output and `cf list`). */
  def describe(rule: CfRule): String = rule match
    case CfRule.CellIs(op, f1, f2, _, _, _) =>
      f2 match
        case Some(hi) =>
          val name = if op == CfOperator.Between then "between" else "notBetween"
          s"cellIs $name $f1 and $hi"
        case None => s"cellIs ${opName(op)} $f1"
    case CfRule.Expression(formula, _, _, _) => s"expression $formula"
    case CfRule.ColorScale(_, mid, _, _) =>
      s"colorScale (${if mid.isDefined then 3 else 2}-point)"
    case CfRule.DataBar(_, _, _, _, _) => "dataBar"
    case CfRule.Top10(rank, percent, bottom, _, _, _) =>
      s"${if bottom then "bottom" else "top"} $rank${if percent then "%" else ""}"
    case CfRule.Text(op, text, _, _, _) =>
      val name = op match
        case CfTextOp.Contains => "contains"
        case CfTextOp.NotContains => "notContains"
        case CfTextOp.BeginsWith => "beginsWith"
        case CfTextOp.EndsWith => "endsWith"
      s"text $name '$text'"
    case CfRule.Preserved(_, _) => "(preserved rule)"

  private def opName(op: CfOperator): String = op match
    case CfOperator.LessThan => "lessThan"
    case CfOperator.LessThanOrEqual => "lessThanOrEqual"
    case CfOperator.Equal => "equal"
    case CfOperator.NotEqual => "notEqual"
    case CfOperator.GreaterThanOrEqual => "greaterThanOrEqual"
    case CfOperator.GreaterThan => "greaterThan"
    case CfOperator.Between => "between"
    case CfOperator.NotBetween => "notBetween"
