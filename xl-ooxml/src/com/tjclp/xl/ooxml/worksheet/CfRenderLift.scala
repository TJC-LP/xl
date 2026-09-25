package com.tjclp.xl.ooxml.worksheet

import scala.xml.Elem

import com.tjclp.xl.cf.{CfRule, ConditionalFormat}
import com.tjclp.xl.ooxml.{FormulaStorage, XmlSecurity, XmlUtil}
import com.tjclp.xl.ooxml.style.DxfCodec
import com.tjclp.xl.styles.Dxf

/**
 * Lifts the formula-backed conditional-format rules xl does not type — `containsBlanks`,
 * `notContainsBlanks`, `containsErrors`, `notContainsErrors` and `timePeriod` — into
 * [[CfRule.Expression]] over the `<formula>` Excel stores for each (`LEN(TRIM(A1))=0`,
 * `ISERROR(A1)`, `FLOOR(A1,1)=TODAY()-1`, ...), with the rule's dxf typed through [[DxfCodec]], so
 * the render path's evaluator can paint them (GH-497). This is what makes Excel's common "Blanks,
 * no format, Stop If True" guard stop the rules below it.
 *
 * RENDER-ONLY: never write the result back. The lifted rule would emit as `type="expression"`,
 * changing the file; the sheet's own `conditionalFormats` keep the verbatim Preserved payload.
 *
 * A rule lifts only when it is fully understood: a `<cfRule>` of one of those kinds whose
 * attributes lie within {type, dxfId, priority, stopIfTrue, timePeriod}, a priority of at least 1,
 * exactly one `<formula>` child, and either no dxfId or a dxf the typed model holds. Everything
 * else — other kinds, typed rules, unreadable blocks — passes through untouched.
 */
object CfRenderLift:

  private val liftable: Set[String] =
    Set("containsBlanks", "notContainsBlanks", "containsErrors", "notContainsErrors", "timePeriod")

  private val allowedAttrs: Set[String] =
    Set("type", "dxfId", "priority", "stopIfTrue", "timePeriod")

  /** `formats` with every liftable Preserved rule lifted (idempotent). */
  def lift(formats: Vector[ConditionalFormat]): Vector[ConditionalFormat] =
    formats.map {
      case ConditionalFormat.Rules(ranges, rules, pivot) =>
        ConditionalFormat.Rules(ranges, rules.map(liftRule), pivot)
      case unreadable: ConditionalFormat.Preserved => unreadable
    }

  private def liftRule(rule: CfRule): CfRule = rule match
    case preserved: CfRule.Preserved => lifted(preserved).getOrElse(preserved)
    case typed => typed

  private def lifted(rule: CfRule.Preserved): Option[CfRule] =
    for
      e <- XmlSecurity.parseSafe(rule.xml, "conditional format render lift").toOption
      if e.label == "cfRule"
      attrs = e.attributes.asAttrMap
      if attrs.get("type").exists(liftable.contains) && attrs.keySet.subsetOf(allowedAttrs)
      priority <- attrs.get("priority").flatMap(_.toIntOption).filter(_ >= 1)
      stop <- attrs.get("stopIfTrue").fold(Option(false))(CfCodec.parseBool)
      formula <- CfCodec.childElems(e).toOption.flatMap(storedFormula)
      dxf <- typedDxf(attrs.contains("dxfId"), rule.dxf)
    yield CfRule.Expression(formula, dxf, priority, stop)

  /** The one `<formula>` child, in the model form (storage prefixes stripped, GH-577). */
  private def storedFormula(children: Vector[Elem]): Option[String] = children match
    case Vector(f) if f.label == "formula" =>
      Some(FormulaStorage.fromStored(XmlUtil.getTextPreservingWhitespace(f)))
    case _ => None

  /** No dxfId: no format. A dxfId: its carried payload, only when the typed model holds it. */
  private def typedDxf(referenced: Boolean, payload: Option[String]): Option[Option[Dxf]] =
    if !referenced then Some(None)
    else
      payload
        .flatMap(xml => XmlSecurity.parseSafe(xml, "conditional format dxf").toOption)
        .flatMap(DxfCodec.parse)
        .map(Some.apply)
