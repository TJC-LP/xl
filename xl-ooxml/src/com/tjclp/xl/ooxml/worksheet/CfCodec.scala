package com.tjclp.xl.ooxml.worksheet

import scala.xml.*

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cf.{CfOperator, CfPoint, CfRule, CfTextOp, Cfvo, ConditionalFormat}
import com.tjclp.xl.ooxml.{XmlSecurity, XmlUtil}
import com.tjclp.xl.ooxml.XmlUtil.{elem, elemOrdered}
import com.tjclp.xl.ooxml.style.DxfCodec
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.Dxf

/**
 * Total conditional-formatting codec (GH-136): `<conditionalFormatting>` blocks ↔ the typed
 * [[ConditionalFormat]] model.
 *
 * Parse is TOTAL with two preservation levels: an unparseable/unknown ENVELOPE (attrs outside
 * {sqref, pivot}, corrupt sqref, non-cfRule children) falls back to
 * [[ConditionalFormat.Preserved]]; an unmodeled RULE (iconSet, timePeriod, unknown attr/child,
 * child extLst, out-of-range or untypeable dxfId, text-template mismatch) falls back to
 * [[CfRule.Preserved]] while its typed envelope keeps shifting under structural edits. Preserved
 * payloads are captured as scope-self-contained canonical XML (the DrawingReader pattern) and
 * re-emitted verbatim.
 *
 * NOTE for future evaluator work: typed rule formulas are stored as text and never re-parsed on
 * read; xl-evaluator's StructuralEditor rewrites them through the formula parser on structural
 * edits. If a future version starts parsing cf formulas at read time, revisit that contract.
 */
object CfCodec:

  // ===== shared tokens =====

  private val operatorToken: Map[CfOperator, String] = Map(
    CfOperator.LessThan -> "lessThan",
    CfOperator.LessThanOrEqual -> "lessThanOrEqual",
    CfOperator.Equal -> "equal",
    CfOperator.NotEqual -> "notEqual",
    CfOperator.GreaterThanOrEqual -> "greaterThanOrEqual",
    CfOperator.GreaterThan -> "greaterThan",
    CfOperator.Between -> "between",
    CfOperator.NotBetween -> "notBetween"
  )
  private val operatorByToken: Map[String, CfOperator] = operatorToken.map(_.swap)

  /** Text families: (type token, operator token) — NB NotContains' operator differs from type. */
  private val textTokens: Map[CfTextOp, (String, String)] = Map(
    CfTextOp.Contains -> ("containsText" -> "containsText"),
    CfTextOp.NotContains -> ("notContainsText" -> "notContains"),
    CfTextOp.BeginsWith -> ("beginsWith" -> "beginsWith"),
    CfTextOp.EndsWith -> ("endsWith" -> "endsWith")
  )
  private val textOpByType: Map[String, CfTextOp] = textTokens.map((op, t) => t._1 -> op)

  /**
   * The derived `<formula>` for a text rule against top-left cell `tl` (relative form). Quotes in
   * the text are doubled per formula string-literal rules. Derived at emission, verified at parse —
   * never stored, so it auto-corrects under structural shifts.
   */
  private[ooxml] def textFormula(op: CfTextOp, text: String, tl: String): String =
    val q = text.replace("\"", "\"\"")
    op match
      case CfTextOp.Contains => s"""NOT(ISERROR(SEARCH("$q",$tl)))"""
      case CfTextOp.NotContains => s"""ISERROR(SEARCH("$q",$tl))"""
      case CfTextOp.BeginsWith => s"""LEFT($tl,LEN("$q"))="$q""""
      case CfTextOp.EndsWith => s"""RIGHT($tl,LEN("$q"))="$q""""

  // ===== parse =====

  /** Parse every `<conditionalFormatting>` block (total — never fails the read). */
  def parseAll(blocks: Seq[Elem], dxfs: Vector[Elem]): Vector[ConditionalFormat] =
    blocks.toVector.map(parseBlock(_, dxfs))

  /**
   * Scope-self-contained canonical capture (the DrawingReader.Preserved pattern). `inheritedScope`
   * supplies bindings the reader's BLOCK-level rebind severed from the element's own chain:
   * WorksheetReader hoists used prefixes onto the block root and cleans descendant scopes, so a
   * RULE captured in isolation must resolve x14/xm (the dataBar extLst pairing) against its parent
   * block's scope or the stored payload is unbound XML that parsePreserved silently drops at dirty
   * emission.
   */
  private def preservedXml(e: Elem, inheritedScope: NamespaceBinding): String =
    XmlUtil.compact(rebindUsedNamespaces(e, includeDefault = true, inheritedScope))

  private def attrKeys(e: Elem): Set[String] = e.attributes.asAttrMap.keySet

  /** Element children; Left when any non-whitespace non-element content is present. */
  private def childElems(e: Elem): Either[Unit, Vector[Elem]] =
    val nonWs = e.child.filterNot {
      case Text(t) => t.forall(_.isWhitespace)
      case _ => false
    }
    val elems = nonWs.collect { case c: Elem => c }.toVector
    if elems.sizeIs == nonWs.size then Right(elems) else Left(())

  private def parseBool(value: String): Option[Boolean] = value match
    case "1" | "true" => Some(true)
    case "0" | "false" => Some(false)
    case _ => None

  private def parseBlock(block: Elem, dxfs: Vector[Elem]): ConditionalFormat =
    val typed: Option[ConditionalFormat.Rules] =
      for
        _ <- Option.when(attrKeys(block).subsetOf(Set("sqref", "pivot")))(())
        sqref <- block.attribute("sqref").map(_.text)
        ranges <- parseSqref(sqref)
        pivot <- block.attribute("pivot").map(_.text) match
          case None => Some(false)
          case Some(v) => parseBool(v)
        children <- childElems(block).toOption
        rules <- Option.when(children.nonEmpty && children.forall(_.label == "cfRule"))(
          children.map(parseRule(_, block.scope, ranges, dxfs))
        )
      yield ConditionalFormat.Rules(ranges, rules, pivot)
    typed.getOrElse(ConditionalFormat.Preserved(preservedXml(block, TopScope)))

  /** Space-split sqref; a single-cell token becomes a 1×1 range. None on any corrupt token. */
  private def parseSqref(sqref: String): Option[Vector[CellRange]] =
    val tokens = sqref.trim.split("\\s+").toVector.filter(_.nonEmpty)
    if tokens.isEmpty then None
    else
      val parsed = tokens.map { token =>
        CellRange
          .parse(token)
          .toOption
          .orElse(ARef.parse(token).toOption.map(r => CellRange(r, r)))
      }
      if parsed.forall(_.isDefined) then Some(parsed.flatten) else None

  /** Per-rule typed parse with the family whitelists; falls back to [[CfRule.Preserved]]. */
  private def parseRule(
    rule: Elem,
    blockScope: NamespaceBinding,
    ranges: Vector[CellRange],
    dxfs: Vector[Elem]
  ): CfRule =
    val parsedPriority = rule.attribute("priority").map(_.text).flatMap(_.toIntOption)
    val typed = childElems(rule).toOption.flatMap { children =>
      if children.exists(_.label == "extLst") then None // protects x14 GUID pairings
      else
        val attrs = rule.attributes.asAttrMap
        for
          // schema requires priority >= 1; anything else rides Preserved un-renumbered
          priority <- parsedPriority.filter(_ >= 1)
          typeToken <- attrs.get("type")
          parsed <- parseFamily(rule, typeToken, attrs, children, priority, ranges, dxfs)
        yield parsed
    }
    typed.getOrElse(CfRule.Preserved(preservedXml(rule, blockScope), parsedPriority))

  private def parseStopIfTrue(attrs: Map[String, String]): Option[Boolean] =
    attrs.get("stopIfTrue") match
      case None => Some(false)
      case Some(v) => parseBool(v)

  /** dxfId attr → typed Dxf: absent → None-dxf; out-of-range or untypeable → degrade. */
  private def parseDxfRef(attrs: Map[String, String], dxfs: Vector[Elem]): Option[Option[Dxf]] =
    attrs.get("dxfId") match
      case None => Some(None)
      case Some(raw) =>
        raw.toIntOption.flatMap(dxfs.lift).flatMap(DxfCodec.parse).map(Some.apply)

  private def formulaTexts(children: Vector[Elem]): Option[Vector[String]] =
    Option.when(children.forall(_.label == "formula"))(
      children.map(XmlUtil.getTextPreservingWhitespace)
    )

  private def parseFamily(
    rule: Elem,
    typeToken: String,
    attrs: Map[String, String],
    children: Vector[Elem],
    priority: Int,
    ranges: Vector[CellRange],
    dxfs: Vector[Elem]
  ): Option[CfRule] =
    val keys = attrs.keySet
    typeToken match
      case "cellIs" =>
        for
          _ <- Option.when(
            keys.subsetOf(Set("type", "dxfId", "priority", "stopIfTrue", "operator"))
          )(())
          op <- attrs.get("operator").flatMap(operatorByToken.get)
          stop <- parseStopIfTrue(attrs)
          dxf <- parseDxfRef(attrs, dxfs)
          formulas <- formulaTexts(children)
          rule <- formulas match
            case Vector(f1) => Some(CfRule.CellIs(op, f1, None, dxf, priority, stop))
            case Vector(f1, f2) if op == CfOperator.Between || op == CfOperator.NotBetween =>
              Some(CfRule.CellIs(op, f1, Some(f2), dxf, priority, stop))
            case _ => None
        yield rule

      case "expression" =>
        for
          _ <- Option.when(keys.subsetOf(Set("type", "dxfId", "priority", "stopIfTrue")))(())
          stop <- parseStopIfTrue(attrs)
          dxf <- parseDxfRef(attrs, dxfs)
          formulas <- formulaTexts(children)
          f <- formulas match
            case Vector(f1) => Some(f1)
            case _ => None
        yield CfRule.Expression(f, dxf, priority, stop)

      case "colorScale" =>
        for
          _ <- Option.when(keys.subsetOf(Set("type", "priority")))(())
          scale <- children match
            case Vector(cs) if cs.label == "colorScale" => parseColorScale(cs, priority)
            case _ => None
        yield scale

      case "dataBar" =>
        for
          _ <- Option.when(keys.subsetOf(Set("type", "priority")))(())
          bar <- children match
            case Vector(db) if db.label == "dataBar" => parseDataBar(db, priority)
            case _ => None
        yield bar

      case "top10" =>
        for
          _ <- Option.when(
            keys.subsetOf(
              Set("type", "dxfId", "priority", "stopIfTrue", "rank", "percent", "bottom")
            )
          )(())
          _ <- Option.when(children.isEmpty)(())
          rank <- attrs.get("rank").flatMap(_.toIntOption)
          percent <- attrs.get("percent").fold(Option(false))(parseBool)
          bottom <- attrs.get("bottom").fold(Option(false))(parseBool)
          stop <- parseStopIfTrue(attrs)
          dxf <- parseDxfRef(attrs, dxfs)
        yield CfRule.Top10(rank, percent, bottom, dxf, priority, stop)

      case t if textOpByType.contains(t) =>
        for
          op <- textOpByType.get(t)
          _ <- Option.when(
            keys.subsetOf(Set("type", "dxfId", "priority", "stopIfTrue", "operator", "text"))
          )(())
          _ <- Option.when(attrs.get("operator").contains(textTokens(op)._2))(())
          text <- attrs.get("text")
          stop <- parseStopIfTrue(attrs)
          dxf <- parseDxfRef(attrs, dxfs)
          formulas <- formulaTexts(children)
          // typed = fully understood: the stored formula must BE the canonical derivation
          tl <- ranges.headOption.map(_.start.toA1)
          rule <- formulas match
            case Vector(f) if f == textFormula(op, text, tl) =>
              Some(CfRule.Text(op, text, dxf, priority, stop))
            case _ => None
        yield rule

      case _ =>
        None // iconSet, timePeriod, aboveAverage, duplicate/uniqueValues, blanks/errors, ...

  private def parseCfvo(e: Elem): Option[Cfvo] =
    childElems(e).toOption.filter(_.isEmpty).flatMap { _ =>
      val attrs = e.attributes.asAttrMap
      if !attrs.keySet.subsetOf(Set("type", "val")) then None
      else
        val valAttr = attrs.get("val")
        attrs.get("type") match
          case Some("min") if valAttr.isEmpty => Some(Cfvo.Min)
          case Some("max") if valAttr.isEmpty => Some(Cfvo.Max)
          case Some("num") => valAttr.flatMap(parseDecimal).map(Cfvo.Num.apply)
          case Some("percent") => valAttr.flatMap(parseDecimal).map(Cfvo.Percent.apply)
          case Some("percentile") => valAttr.flatMap(parseDecimal).map(Cfvo.Percentile.apply)
          case Some("formula") => valAttr.map(f => Cfvo.Formula(f.stripPrefix("=")))
          case _ => None
    }

  private def parseDecimal(s: String): Option[BigDecimal] =
    scala.util.Try(BigDecimal(s.trim)).toOption

  private def parseCfColor(e: Elem): Option[com.tjclp.xl.styles.color.Color] =
    DxfCodec.parseColor(e)

  private def parseColorScale(cs: Elem, priority: Int): Option[CfRule] =
    childElems(cs).toOption.flatMap { children =>
      if attrKeys(cs).nonEmpty then None
      else
        // probe-verified child order: ALL cfvos then ALL colors, counts matching (2 or 3)
        val (cfvos, rest) = children.span(_.label == "cfvo")
        val colorsOk = rest.forall(_.label == "color")
        if !colorsOk || cfvos.size != rest.size then None
        else
          val points = cfvos.zip(rest).map { (v, c) =>
            for
              cfvo <- parseCfvo(v)
              color <- parseCfColor(c)
            yield CfPoint(cfvo, color)
          }
          points.sequenceOpt.flatMap {
            case Vector(min, max) => Some(CfRule.ColorScale(min, None, max, priority))
            case Vector(min, mid, max) => Some(CfRule.ColorScale(min, Some(mid), max, priority))
            case _ => None
          }
    }

  extension [A](v: Vector[Option[A]])
    private def sequenceOpt: Option[Vector[A]] =
      v.foldLeft(Option(Vector.empty[A])) { (acc, o) => acc.flatMap(a => o.map(a :+ _)) }

  private def parseDataBar(db: Elem, priority: Int): Option[CfRule] =
    childElems(db).toOption.flatMap { children =>
      if !attrKeys(db).subsetOf(Set("showValue")) then None
      else
        for
          showValue <- db.attribute("showValue").map(_.text).fold(Option(true))(parseBool)
          bar <- children match
            case Vector(minE, maxE, colorE)
                if minE.label == "cfvo" && maxE.label == "cfvo" && colorE.label == "color" =>
              for
                min <- parseCfvo(minE)
                max <- parseCfvo(maxE)
                color <- parseCfColor(colorE)
              yield CfRule.DataBar(min, max, color, showValue, priority)
            case _ => None
        yield bar
    }

  // ===== emission =====

  /** Typed-rule Dxf values in (block, rule) encounter order (with duplicates; callers dedup). */
  def collectDxfs(cfs: Vector[ConditionalFormat]): Vector[Dxf] =
    cfs.flatMap {
      case ConditionalFormat.Rules(_, rules, _) =>
        rules.flatMap {
          case r: CfRule.CellIs => r.dxf
          case r: CfRule.Expression => r.dxf
          case r: CfRule.Top10 => r.dxf
          case r: CfRule.Text => r.dxf
          case _: CfRule.ColorScale | _: CfRule.DataBar | _: CfRule.Preserved => None
        }
      case _: ConditionalFormat.Preserved => Vector.empty
    }

  /**
   * Canonical emission: blocks in vector order, rules in vector order, attributes in Excel's order.
   * Empty-envelope blocks are dropped (unexpressible in OOXML). Safety net: any residual
   * `priority <= 0` (direct construction bypassing `Sheet.conditionalFormat`) is assigned
   * `max+1, +2, ...` in document order, saturating at `Int.MaxValue` (never a schema-invalid
   * negative); explicit and Preserved priorities are NEVER renumbered.
   */
  def toElems(cfs: Vector[ConditionalFormat], dxfIds: Map[Dxf, Int]): Seq[Elem] =
    val maxExisting = Sheet.maxCfPriority(cfs)
    val (stamped, _) = cfs.foldLeft((Vector.empty[ConditionalFormat], maxExisting)) {
      case ((acc, cur), ConditionalFormat.Rules(ranges, rules, pivot)) =>
        val (rs, next) = rules.foldLeft((Vector.empty[CfRule], cur)) { case ((a, c), r) =>
          CfRule.priorityOf(r) match
            case Some(p) if p <= 0 =>
              val n = Sheet.nextCfPriority(c)
              (a :+ CfRule.withPriority(r, n), n)
            case _ => (a :+ r, c)
        }
        (acc :+ ConditionalFormat.Rules(ranges, rs, pivot), next)
      case ((acc, cur), p: ConditionalFormat.Preserved) => (acc :+ p, cur)
    }
    stamped.flatMap {
      case ConditionalFormat.Rules(ranges, rules, pivot) if ranges.nonEmpty && rules.nonEmpty =>
        val ruleElems = rules.flatMap(ruleToElem(_, ranges, dxfIds))
        Option.when(ruleElems.nonEmpty) {
          val attrs =
            (if pivot then Seq("pivot" -> "1") else Seq.empty) :+
              ("sqref" -> ranges.map(_.toA1).mkString(" "))
          elemOrdered("conditionalFormatting", attrs*)(ruleElems*)
        }
      case _: ConditionalFormat.Rules => None
      case ConditionalFormat.Preserved(xml) => parsePreserved(xml)
    }

  /** Preserved payloads re-parse to Elems; a non-canonical hand-built payload drops silently. */
  private def parsePreserved(xml: String): Option[Elem] =
    XmlSecurity.parseSafe(xml, "preserved conditional formatting").toOption

  private def formulaElem(text: String): Elem = elem("formula")(Text(text))

  private def boolAttr(name: String, value: Boolean): Seq[(String, String)] =
    if value then Seq(name -> "1") else Seq.empty

  private def ruleToElem(
    rule: CfRule,
    ranges: Vector[CellRange],
    dxfIds: Map[Dxf, Int]
  ): Option[Elem] =
    def dxfAttr(dxf: Option[Dxf]): Seq[(String, String)] =
      dxf.flatMap(dxfIds.get).map(id => "dxfId" -> id.toString).toList
    rule match
      case CfRule.CellIs(op, f1, f2, dxf, priority, stop) =>
        val attrs = Seq("type" -> "cellIs") ++ dxfAttr(dxf) ++
          Seq("priority" -> priority.toString) ++ boolAttr("stopIfTrue", stop) ++
          Seq("operator" -> operatorToken(op))
        Some(elemOrdered("cfRule", attrs*)((formulaElem(f1) +: f2.map(formulaElem).toList)*))

      case CfRule.Expression(f, dxf, priority, stop) =>
        val attrs = Seq("type" -> "expression") ++ dxfAttr(dxf) ++
          Seq("priority" -> priority.toString) ++ boolAttr("stopIfTrue", stop)
        Some(elemOrdered("cfRule", attrs*)(formulaElem(f)))

      case CfRule.ColorScale(min, mid, max, priority) =>
        val points = Vector(Some(min), mid, Some(max)).flatten
        val scale = elem("colorScale")(
          (points.map(p => cfvoToXml(p.cfvo)) ++ points.map(p => cfColorToXml(p.color)))*
        )
        Some(elemOrdered("cfRule", "type" -> "colorScale", "priority" -> priority.toString)(scale))

      case CfRule.DataBar(min, max, color, showValue, priority) =>
        val barAttrs = if showValue then Seq.empty else Seq("showValue" -> "0")
        val bar = elemOrdered("dataBar", barAttrs*)(
          cfvoToXml(min),
          cfvoToXml(max),
          cfColorToXml(color)
        )
        Some(elemOrdered("cfRule", "type" -> "dataBar", "priority" -> priority.toString)(bar))

      case CfRule.Top10(rank, percent, bottom, dxf, priority, stop) =>
        val attrs = Seq("type" -> "top10") ++ dxfAttr(dxf) ++
          Seq("priority" -> priority.toString) ++ boolAttr("stopIfTrue", stop) ++
          Seq("rank" -> rank.toString) ++ boolAttr("percent", percent) ++
          boolAttr("bottom", bottom)
        Some(elemOrdered("cfRule", attrs*)())

      case CfRule.Text(op, text, dxf, priority, stop) =>
        val (typeToken, opToken) = textTokens(op)
        val attrs = Seq("type" -> typeToken) ++ dxfAttr(dxf) ++
          Seq("priority" -> priority.toString) ++ boolAttr("stopIfTrue", stop) ++
          Seq("operator" -> opToken, "text" -> text)
        // derived against the FIRST range's top-left; total via headOption (empty ranges blocks
        // are dropped before this point)
        ranges.headOption.map { first =>
          val formula = textFormula(op, text, first.start.toA1)
          elemOrdered("cfRule", attrs*)(formulaElem(formula))
        }

      case CfRule.Preserved(xml, _) => parsePreserved(xml)

  private def cfvoToXml(cfvo: Cfvo): Elem = cfvo match
    case Cfvo.Min => elemOrdered("cfvo", "type" -> "min")()
    case Cfvo.Max => elemOrdered("cfvo", "type" -> "max")()
    case Cfvo.Num(v) => elemOrdered("cfvo", "type" -> "num", "val" -> v.toString)()
    case Cfvo.Percent(v) => elemOrdered("cfvo", "type" -> "percent", "val" -> v.toString)()
    case Cfvo.Percentile(v) => elemOrdered("cfvo", "type" -> "percentile", "val" -> v.toString)()
    case Cfvo.Formula(f) => elemOrdered("cfvo", "type" -> "formula", "val" -> f)()

  private def cfColorToXml(color: com.tjclp.xl.styles.color.Color): Elem =
    DxfCodec.colorToXml(color)
