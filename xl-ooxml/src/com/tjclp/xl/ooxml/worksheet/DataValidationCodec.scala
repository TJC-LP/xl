package com.tjclp.xl.ooxml.worksheet

import scala.xml.*

import com.tjclp.xl.ooxml.{XmlSecurity, XmlUtil}
import com.tjclp.xl.ooxml.XmlUtil.{elem, elemOrdered}
import com.tjclp.xl.sheets.{DataValidation, DvKind}

/**
 * Total data-validation codec (GH-375): the `<dataValidations>` container ↔ the typed
 * [[DataValidation]] model — the CfCodec contract applied to validations.
 *
 * Parse is TOTAL: an entry outside the typed subset (any type other than `list`, operators,
 * prompt/error messages, formula2, unknown attrs/children) falls back to
 * [[DataValidation.Preserved]] whose payload is scope-self-contained canonical XML re-emitted
 * verbatim. THE INVERSION: the OOXML attribute `showDropDown="1"` SUPPRESSES the in-cell dropdown;
 * the friendly model field `showDropdown` is its negation, so the attribute is emitted only when
 * the model says false.
 *
 * Unlike conditional formatting (one element per block), OOXML allows a single
 * `<dataValidations count="N">` container: emission always regenerates ONE container holding typed
 * emissions and Preserved payloads together, overlaying `count` while any unmodeled container
 * attributes of the source (disablePrompts, xWindow, ...) ride through via `base`.
 */
object DataValidationCodec:

  private val typedAttrs = Set("type", "sqref", "allowBlank", "showDropDown")

  // ===== parse =====

  /** Parse the (single) `<dataValidations>` container (total — never fails the read). */
  def parseAll(container: Option[Elem]): Vector[DataValidation] =
    container.toList.toVector.flatMap { dvs =>
      dvs.child.collect { case e: Elem => e }.map(parseEntry(_, dvs.scope))
    }

  /** Per-entry typed parse with the list-family whitelist; falls back to Preserved. */
  private def parseEntry(entry: Elem, containerScope: NamespaceBinding): DataValidation =
    val typed: Option[DataValidation.Rules] =
      for
        _ <- Option.when(entry.label == "dataValidation")(())
        _ <- Option.when(CfCodec.attrKeys(entry).subsetOf(typedAttrs))(())
        _ <- Option.when(entry.attribute("type").map(_.text).contains("list"))(())
        sqref <- entry.attribute("sqref").map(_.text)
        ranges <- CfCodec.parseSqref(sqref)
        allowBlank <- entry.attribute("allowBlank").map(_.text) match
          case None => Some(false) // OOXML schema default
          case Some(v) => CfCodec.parseBool(v)
        suppressDropdown <- entry.attribute("showDropDown").map(_.text) match
          case None => Some(false) // absent attr = dropdown SHOWN
          case Some(v) => CfCodec.parseBool(v)
        children <- CfCodec.childElems(entry).toOption
        formula <- children match
          case Vector(f1) if f1.label == "formula1" =>
            Some(XmlUtil.getTextPreservingWhitespace(f1))
          case _ => None
      yield DataValidation.Rules(
        ranges,
        DvKind.List(formula),
        allowBlank = allowBlank,
        showDropdown = !suppressDropdown
      )
    typed.getOrElse(DataValidation.Preserved(CfCodec.preservedXml(entry, containerScope)))

  // ===== emission =====

  /**
   * Canonical emission: ONE `<dataValidations count="N">` container with entries in vector order
   * (typed emissions and Preserved payloads together); None when nothing emits. Empty-ranges typed
   * entries are dropped (unexpressible in OOXML); a Preserved payload that fails to re-parse drops
   * silently (the CfCodec contract). `base` carries the preserved source container so unmodeled
   * container attributes ride through a dirty write; `count` is always restamped.
   */
  def toElem(dvs: Vector[DataValidation], base: Option[Elem]): Option[Elem] =
    val children: Vector[Elem] = dvs.flatMap {
      case DataValidation.Rules(ranges, kind, allowBlank, showDropdown) if ranges.nonEmpty =>
        kind match
          case DvKind.List(formula) =>
            val attrs = Seq("type" -> "list") ++
              (if allowBlank then Seq("allowBlank" -> "1") else Seq.empty) ++
              // THE INVERSION: showDropDown="1" suppresses the dropdown
              (if showDropdown then Seq.empty else Seq("showDropDown" -> "1")) ++
              Seq("sqref" -> ranges.map(_.toA1).mkString(" "))
            Some(elemOrdered("dataValidation", attrs*)(elem("formula1")(Text(formula))))
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
