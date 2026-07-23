package com.tjclp.xl.ooxml.worksheet

import scala.xml.*

import com.tjclp.xl.ooxml.{XmlSecurity, XmlUtil}
import com.tjclp.xl.ooxml.XmlUtil.{elem, elemOrdered}
import com.tjclp.xl.sheets.{
  DataValidation,
  DvBoundedType,
  DvErrorStyle,
  DvKind,
  DvMessages,
  DvOperator
}

/**
 * Total data-validation codec (GH-375, widened by GH-429): the `<dataValidations>` container ↔ the
 * typed [[DataValidation]] model — the CfCodec contract applied to validations.
 *
 * Parse is TOTAL: an entry outside the typed subset (unknown type token, foreign/prefixed attrs
 * such as `xr:uid`/`imeMode`, corrupt sqref, unexpected children) falls back to
 * [[DataValidation.Preserved]] whose payload is scope-self-contained canonical XML re-emitted
 * verbatim. The typed subset covers what Excel actually stamps — list/custom/anyValue/bounded
 * kinds, the operator, and the full prompt/error message set (`showInputMessage="1"
 * showErrorMessage="1"` appear on virtually every Excel-authored validation; before GH-429 they
 * forced every real-world entry into Preserved, which is structurally inert). THE INVERSION: the
 * OOXML attribute `showDropDown="1"` SUPPRESSES the in-cell dropdown; the friendly model field
 * `showDropdown` is its negation, so the attribute is emitted only when the model says false.
 *
 * Unlike conditional formatting (one element per block), OOXML allows a single
 * `<dataValidations count="N">` container: emission always regenerates ONE container holding typed
 * emissions and Preserved payloads together, overlaying `count` while any unmodeled container
 * attributes of the source (disablePrompts, xWindow, ...) ride through via `base`.
 */
object DataValidationCodec:

  private val typedAttrs = Set(
    "type",
    "sqref",
    "allowBlank",
    "showDropDown",
    "showInputMessage",
    "showErrorMessage",
    "promptTitle",
    "prompt",
    "errorTitle",
    "error",
    "errorStyle",
    "operator"
  )

  private val boundedTypes: Map[String, DvBoundedType] = Map(
    "whole" -> DvBoundedType.Whole,
    "decimal" -> DvBoundedType.Decimal,
    "date" -> DvBoundedType.Date,
    "time" -> DvBoundedType.Time,
    "textLength" -> DvBoundedType.TextLength
  )

  private val operators: Map[String, DvOperator] = Map(
    "between" -> DvOperator.Between,
    "notBetween" -> DvOperator.NotBetween,
    "equal" -> DvOperator.Equal,
    "notEqual" -> DvOperator.NotEqual,
    "lessThan" -> DvOperator.LessThan,
    "lessThanOrEqual" -> DvOperator.LessThanOrEqual,
    "greaterThan" -> DvOperator.GreaterThan,
    "greaterThanOrEqual" -> DvOperator.GreaterThanOrEqual
  )

  private val errorStyles: Map[String, DvErrorStyle] = Map(
    "stop" -> DvErrorStyle.Stop,
    "warning" -> DvErrorStyle.Warning,
    "information" -> DvErrorStyle.Information
  )

  // token printers (emission omits schema-default values)
  private def boundedToken(t: DvBoundedType): String = t match
    case DvBoundedType.Whole => "whole"
    case DvBoundedType.Decimal => "decimal"
    case DvBoundedType.Date => "date"
    case DvBoundedType.Time => "time"
    case DvBoundedType.TextLength => "textLength"

  private def operatorToken(op: DvOperator): String = op match
    case DvOperator.Between => "between"
    case DvOperator.NotBetween => "notBetween"
    case DvOperator.Equal => "equal"
    case DvOperator.NotEqual => "notEqual"
    case DvOperator.LessThan => "lessThan"
    case DvOperator.LessThanOrEqual => "lessThanOrEqual"
    case DvOperator.GreaterThan => "greaterThan"
    case DvOperator.GreaterThanOrEqual => "greaterThanOrEqual"

  private def errorStyleToken(s: DvErrorStyle): String = s match
    case DvErrorStyle.Stop => "stop"
    case DvErrorStyle.Warning => "warning"
    case DvErrorStyle.Information => "information"

  // ===== parse =====

  /** Parse the (single) `<dataValidations>` container (total — never fails the read). */
  def parseAll(container: Option[Elem]): Vector[DataValidation] =
    container.toList.toVector.flatMap { dvs =>
      dvs.child.collect { case e: Elem => e }.map(parseEntry(_, dvs.scope))
    }

  /** Per-entry typed parse with the widened whitelist; falls back to Preserved. */
  private def parseEntry(entry: Elem, containerScope: NamespaceBinding): DataValidation =
    val typed: Option[DataValidation.Rules] =
      for
        _ <- Option.when(entry.label == "dataValidation")(())
        _ <- Option.when(CfCodec.attrKeys(entry).subsetOf(typedAttrs))(())
        sqref <- entry.attribute("sqref").map(_.text)
        ranges <- CfCodec.parseSqref(sqref)
        allowBlank <- boolAttr(entry, "allowBlank")
        suppressDropdown <- boolAttr(entry, "showDropDown")
        messages <- parseMessages(entry)
        children <- CfCodec.childElems(entry).toOption
        kind <- parseKind(
          entry.attribute("type").map(_.text),
          entry.attribute("operator").map(_.text),
          children
        )
      yield DataValidation.Rules(
        ranges,
        kind,
        allowBlank = allowBlank,
        showDropdown = !suppressDropdown,
        messages = messages
      )
    typed.getOrElse(DataValidation.Preserved(CfCodec.preservedXml(entry, containerScope)))

  /** Absent boolean attr = the OOXML schema default (false); a corrupt token fails the parse. */
  private def boolAttr(entry: Elem, name: String): Option[Boolean] =
    entry.attribute(name).map(_.text) match
      case None => Some(false)
      case Some(v) => CfCodec.parseBool(v)

  /** Text attrs decode `_xHHHH_` escapes (attribute-value normalization eats raw LF/TAB/CR). */
  private def textAttr(entry: Elem, name: String): Option[String] =
    entry.attribute(name).map(n => XmlUtil.decodeXstring(n.text))

  private def parseMessages(entry: Elem): Option[DvMessages] =
    for
      showInput <- boolAttr(entry, "showInputMessage")
      showError <- boolAttr(entry, "showErrorMessage")
      errorStyle <- entry.attribute("errorStyle").map(_.text) match
        case None => Some(DvErrorStyle.Stop)
        case Some(v) => errorStyles.get(v)
    yield DvMessages(
      showInputMessage = showInput,
      showErrorMessage = showError,
      promptTitle = textAttr(entry, "promptTitle"),
      prompt = textAttr(entry, "prompt"),
      errorTitle = textAttr(entry, "errorTitle"),
      error = textAttr(entry, "error"),
      errorStyle = errorStyle
    )

  /**
   * Per-kind parse. `operator` is accepted only for the bounded families (parse-and-drop on
   * list/custom/none would lose the attribute on a dirty rewrite); formula children must match the
   * kind exactly — anything else falls back to Preserved.
   */
  private def parseKind(
    typeAttr: Option[String],
    opAttr: Option[String],
    children: Vector[Elem]
  ): Option[DvKind] =
    def formulas: Option[(String, Option[String])] = children match
      case Vector(f1) if f1.label == "formula1" =>
        Some((XmlUtil.getTextPreservingWhitespace(f1), None))
      case Vector(f1, f2) if f1.label == "formula1" && f2.label == "formula2" =>
        Some(
          (
            XmlUtil.getTextPreservingWhitespace(f1),
            Some(XmlUtil.getTextPreservingWhitespace(f2))
          )
        )
      case _ => None
    typeAttr match
      case None | Some("none") =>
        Option.when(children.isEmpty && opAttr.isEmpty)(DvKind.AnyValue)
      case Some("list") =>
        formulas.collect { case (f1, None) if opAttr.isEmpty => DvKind.List(f1) }
      case Some("custom") =>
        formulas.collect { case (f1, None) if opAttr.isEmpty => DvKind.Custom(f1) }
      case Some(t) =>
        for
          dataType <- boundedTypes.get(t)
          operator <- opAttr.fold(Option(DvOperator.Between))(operators.get)
          (f1, f2) <- formulas
        yield DvKind.Bounded(dataType, operator, f1, f2)

  // ===== emission =====

  /**
   * Canonical emission: ONE `<dataValidations count="N">` container with entries in vector order
   * (typed emissions and Preserved payloads together); None when nothing emits. Empty-ranges typed
   * entries are dropped (unexpressible in OOXML); a Preserved payload that fails to re-parse drops
   * silently (the CfCodec contract). `base` carries the preserved source container so unmodeled
   * container attributes ride through a dirty write; `count` is always restamped.
   *
   * Attributes emit in Excel's stamp order with schema-default values omitted (incl. `type` for
   * AnyValue); message text goes through [[XmlUtil.escapeXstringAttr]] so multiline prompts survive
   * attribute-value normalization.
   */
  def toElem(dvs: Vector[DataValidation], base: Option[Elem]): Option[Elem] =
    val children: Vector[Elem] = dvs.flatMap {
      case rules: DataValidation.Rules if rules.ranges.nonEmpty =>
        val (typeToken, operator, formula1, formula2) = rules.kind match
          case DvKind.List(f) => (Some("list"), None, Some(f), None)
          case DvKind.Custom(f) => (Some("custom"), None, Some(f), None)
          case DvKind.AnyValue => (None, None, None, None)
          case DvKind.Bounded(t, op, f1, f2) =>
            (Some(boundedToken(t)), Some(op).filter(_ != DvOperator.Between), Some(f1), f2)
        val m = rules.messages
        val attrs: Seq[(String, String)] = Seq(
          typeToken.map("type" -> _),
          Option.when(m.errorStyle != DvErrorStyle.Stop)(
            "errorStyle" -> errorStyleToken(m.errorStyle)
          ),
          operator.map(op => "operator" -> operatorToken(op)),
          Option.when(rules.allowBlank)("allowBlank" -> "1"),
          // THE INVERSION: showDropDown="1" suppresses the dropdown
          Option.when(!rules.showDropdown)("showDropDown" -> "1"),
          Option.when(m.showInputMessage)("showInputMessage" -> "1"),
          Option.when(m.showErrorMessage)("showErrorMessage" -> "1"),
          m.errorTitle.map(t => "errorTitle" -> XmlUtil.escapeXstringAttr(t)),
          m.error.map(t => "error" -> XmlUtil.escapeXstringAttr(t)),
          m.promptTitle.map(t => "promptTitle" -> XmlUtil.escapeXstringAttr(t)),
          m.prompt.map(t => "prompt" -> XmlUtil.escapeXstringAttr(t)),
          Some("sqref" -> rules.ranges.map(_.toA1).mkString(" "))
        ).flatten
        val formulaElems =
          formula1.map(f => elem("formula1")(Text(f))).toList ++
            formula2.map(f => elem("formula2")(Text(f))).toList
        Some(elemOrdered("dataValidation", attrs*)(formulaElems*))
      case _: DataValidation.Rules => None
      case DataValidation.Preserved(xml) => parsePreserved(xml)
    }
    Option.when(children.nonEmpty) {
      val baseAttrs = base.map(_.attributes.remove("count")).getOrElse(Null)
      val attrs = new UnprefixedAttribute("count", children.size.toString, baseAttrs)
      base match
        case Some(b) => b.copy(attributes = attrs, child = children)
        case None =>
          Elem(null, "dataValidations", attrs, TopScope, minimizeEmpty = false, children*)
    }

  /** Preserved payloads re-parse to Elems; a non-canonical hand-built payload drops silently. */
  private def parsePreserved(xml: String): Option[Elem] =
    XmlSecurity.parseSafe(xml, "preserved data validation").toOption
