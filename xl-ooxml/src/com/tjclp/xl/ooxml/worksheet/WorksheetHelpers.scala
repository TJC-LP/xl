package com.tjclp.xl.ooxml.worksheet

import scala.xml.*

import com.tjclp.xl.addressing.{ARef, Column}
import com.tjclp.xl.ooxml.XmlUtil
import com.tjclp.xl.ooxml.XmlUtil.nsSpreadsheetML
import com.tjclp.xl.sheets.{ColumnProperties, PageSetup, RowProperties, Sheet}

// Default namespaces for generated worksheets. Real files capture the original scope/attributes to
// avoid redundant declarations and preserve mc/x14/xr bindings from the source sheet.
private[ooxml] val defaultWorksheetScope: NamespaceBinding =
  NamespaceBinding(null, nsSpreadsheetML, NamespaceBinding("r", XmlUtil.nsRelationships, TopScope))

/** Check if a NamespaceBinding chain contains a given prefix. */
private[ooxml] def scopeHasPrefix(scope: NamespaceBinding, prefix: String): Boolean =
  @annotation.tailrec
  def loop(s: NamespaceBinding): Boolean =
    if s == null || s == TopScope then false
    else if s.prefix == prefix then true
    else loop(s.parent)
  loop(scope)

private[ooxml] def cleanNamespaces(elem: Elem): Elem =
  val cleanedChildren = elem.child.map {
    case e: Elem => cleanNamespaces(e)
    case other => other
  }
  elem.copy(scope = TopScope, child = cleanedChildren)

// ===== Preserved inline worksheet elements (GH-232) =====
// Inline CT_Worksheet children that are listed in WorksheetReader.knownElements (so they are
// excluded from the `otherElements` catch-all) but have NO dedicated field on OoxmlWorksheet.
// Without capturing these, any modification of a sheet silently DROPS data validations, sheet
// protection, autofilters, hyperlinks, etc. They are captured into OoxmlWorksheet.preservedKnown
// and re-emitted at their correct OOXML schema positions (ECMA-376 18.3.1.99) by the groups below.
private[worksheet] val preservedAfterSheetData: Seq[String] =
  Seq(
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews"
  )
private[worksheet] val preservedAfterMergeCells: Seq[String] = Seq("phoneticPr")
private[worksheet] val preservedAfterCondFmt: Seq[String] = Seq("dataValidations", "hyperlinks")
private[worksheet] val preservedAfterCustomProps: Seq[String] =
  Seq("cellWatches", "ignoredErrors", "smartTags")
private[worksheet] val preservedAfterLegacyDrawing: Seq[String] = Seq("legacyDrawingHF")
private[worksheet] val preservedAfterControls: Seq[String] = Seq("webPublishItems")

/** Union of all inline labels preserved via OoxmlWorksheet.preservedKnown (used by the reader). */
private[ooxml] val preservedKnownLabels: Set[String] =
  (preservedAfterSheetData ++ preservedAfterMergeCells ++ preservedAfterCondFmt ++
    preservedAfterCustomProps ++ preservedAfterLegacyDrawing ++ preservedAfterControls).toSet

/**
 * All inline CT_Worksheet child labels the reader recognizes (excluded from the `otherElements`
 * verbatim catch-all). INVARIANT: every label here MUST be either modeled by a dedicated
 * OoxmlWorksheet field, regenerated (sheetData, mergeCells), OR captured in `preservedKnownLabels`.
 * A label that is none of these is silently dropped on modification — the GH-232 class of bug. The
 * invariant is enforced by WorksheetPreservationInvariantSpec.
 */
private[ooxml] val worksheetKnownElements: Set[String] = Set(
  // Modeled by a dedicated field or regenerated
  "sheetPr",
  "dimension",
  "sheetViews",
  "sheetFormatPr",
  "cols",
  "sheetData",
  "mergeCells",
  "conditionalFormatting",
  "printOptions",
  "rowBreaks",
  "colBreaks",
  "customProperties",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "drawing",
  "legacyDrawing",
  "picture",
  "oleObjects",
  "controls",
  "tableParts",
  "extLst",
  // Preserved via OoxmlWorksheet.preservedKnown (GH-232)
  "sheetCalcPr",
  "sheetProtection",
  "protectedRanges",
  "scenarios",
  "autoFilter",
  "sortState",
  "dataConsolidate",
  "customSheetViews",
  "phoneticPr",
  "dataValidations",
  "hyperlinks",
  "cellWatches",
  "ignoredErrors",
  "smartTags",
  "legacyDrawingHF",
  "webPublishItems"
)

// ===== Hyperlink authoring (GH-235) =====
// Cell.hyperlink is the typed model. On write we emit a <hyperlinks> element (and, for external
// targets, sheet .rels relationships) so the model is no longer a silent no-op.

/** A hyperlink to emit: cell ref, target, whether external (needs a rel), and its rel id. */
private[ooxml] case class HyperlinkEntry(
  ref: ARef,
  target: String,
  external: Boolean,
  relId: String
)

/** External hyperlinks (URLs/mailto/file) need a relationship; others are internal `location`s. */
private[ooxml] def isExternalHyperlink(target: String): Boolean =
  target.contains("://") || target.startsWith("mailto:") || target.startsWith("file:")

/**
 * Collect a sheet's hyperlinks in deterministic order (by cell ref). External targets get a stable
 * relationship id ("rIdHL{n}") used by BOTH the `<hyperlink r:id>` attribute and the sheet .rels,
 * so the two agree without explicit plumbing (GH-235).
 */
private[ooxml] def collectHyperlinks(sheet: Sheet): Seq[HyperlinkEntry] =
  val sorted = sheet.cells.values.toSeq
    .filter(_.hyperlink.isDefined)
    .sortBy(c => (c.ref.row.index0, c.ref.col.index0))
  val relIdByRef = sorted
    .filter(c => isExternalHyperlink(c.hyperlink.getOrElse("")))
    .zipWithIndex
    .map { case (c, i) => c.ref -> s"rIdHL${i + 1}" }
    .toMap
  sorted.map { c =>
    val target = c.hyperlink.getOrElse("")
    HyperlinkEntry(c.ref, target, isExternalHyperlink(target), relIdByRef.getOrElse(c.ref, ""))
  }

/** Build a `<hyperlinks>` element from collected entries (GH-235), or None if empty. */
private[ooxml] def buildHyperlinksElem(entries: Seq[HyperlinkEntry]): Option[Elem] =
  if entries.isEmpty then None
  else
    val children = entries.map { e =>
      val tail: MetaData =
        if e.external then new PrefixedAttribute("r", "id", e.relId, Null)
        else new UnprefixedAttribute("location", e.target, Null)
      val attrs = new UnprefixedAttribute("ref", e.ref.toA1, tail)
      Elem(null, "hyperlink", attrs, TopScope, minimizeEmpty = true)
    }
    Some(Elem(null, "hyperlinks", Null, TopScope, minimizeEmpty = false, children*))

/**
 * Group consecutive columns with identical properties into spans.
 *
 * Excel's `<col>` element supports min/max attributes to apply the same properties to a range of
 * columns. This reduces file size by avoiding repeated `<col>` elements.
 *
 * @return
 *   Sequence of (minCol, maxCol, properties) tuples for span generation
 */
private[ooxml] def groupConsecutiveColumns(
  props: Map[Column, ColumnProperties]
): Seq[(Column, Column, ColumnProperties)] =
  if props.isEmpty then Seq.empty
  else
    props.toSeq
      .sortBy(_._1.index0)
      .foldLeft(Vector.empty[(Column, Column, ColumnProperties)]) { case (acc, (col, p)) =>
        acc.lastOption match
          case Some((minCol, maxCol, lastProps))
              if maxCol.index0 + 1 == col.index0 && lastProps == p =>
            // Extend current span
            acc.dropRight(1) :+ (minCol, col, p)
          case _ =>
            // Start new span
            acc :+ (col, col, p)
      }

/**
 * Build `<cols>` XML element from domain column properties.
 *
 * Generates OOXML-compliant column definitions: {{{<cols> <col min="1" max="3" width="15.5"
 * customWidth="1" hidden="1"/> </cols>}}}
 *
 * @return
 *   Some(cols element) if there are column properties, None otherwise
 */
private[ooxml] def buildColsElement(sheet: Sheet): Option[Elem] =
  val props = sheet.columnProperties
  if props.isEmpty then None
  else
    val spans = groupConsecutiveColumns(props)
    val colElems = spans.map { case (minCol, maxCol, p) =>
      // Build attribute sequence (only include non-default values)
      val attrs = Seq.newBuilder[(String, String)]
      attrs += ("min" -> (minCol.index0 + 1).toString) // OOXML is 1-based
      attrs += ("max" -> (maxCol.index0 + 1).toString)
      p.width.foreach { w =>
        attrs += ("width" -> w.toString)
        attrs += ("customWidth" -> "1")
      }
      if p.hidden then attrs += ("hidden" -> "1")
      p.outlineLevel.foreach(l => attrs += ("outlineLevel" -> l.toString))
      if p.collapsed then attrs += ("collapsed" -> "1")
      // Note: styleId would need remapping to workbook-level index (deferred)

      XmlUtil.elemOrdered("col", attrs.result()*)( /* no children */ )
    }
    Some(XmlUtil.elem("cols")(colElems*))

// ===== Print setup authoring (GH-259, GH-266) =====
// PageSetup is the typed model for <pageMargins>, <pageSetup>, <headerFooter>, and the
// sheetPr/pageSetUpPr fitToPage flag. The merge helpers below overlay the modeled fields onto any
// preserved source XML so unmodeled attributes (paperSize, printer r:id, scaleWithDoc, ...)
// survive a rewrite.

/** Replace or remove an unprefixed attribute on an element (deterministic for a given input). */
private def setOrRemoveAttr(e: Elem, key: String, value: Option[String]): Elem =
  value match
    case Some(v) => e % new UnprefixedAttribute(key, v, Null)
    case None => e.copy(attributes = e.attributes.remove(key))

/**
 * Build a `<pageMargins>` element from the domain margins, or None when not set.
 *
 * All six attributes are required by the schema (CT_PageMargins), so the element is fully modeled
 * and regenerated rather than merged.
 */
private[ooxml] def buildPageMarginsElem(pageSetup: Option[PageSetup]): Option[Elem] =
  pageSetup.flatMap(_.margins).map { m =>
    XmlUtil.elem(
      "pageMargins",
      "left" -> m.left.toString,
      "right" -> m.right.toString,
      "top" -> m.top.toString,
      "bottom" -> m.bottom.toString,
      "header" -> m.header.toString,
      "footer" -> m.footer.toString
    )()
  }

/**
 * Merge the modeled `<pageSetup>` attributes (scale, orientation, fitToWidth, fitToHeight) into the
 * preserved element, keeping unmodeled attributes (paperSize, r:id, dpi, ...) intact.
 *
 * Default-valued fields remove their attribute (the OOXML defaults match), so parse-then-merge of
 * an untouched sheet is an identity. When nothing is preserved and every modeled field is at its
 * default, no element is emitted.
 */
private[ooxml] def mergePageSetupElem(
  existing: Option[Elem],
  pageSetup: Option[PageSetup]
): Option[Elem] =
  pageSetup match
    case None => existing
    case Some(ps) =>
      val base = existing.getOrElse(XmlUtil.elem("pageSetup")())
      val overrides: Seq[(String, Option[String])] = Seq(
        "scale" -> Some(ps.scale.toString).filter(_ => ps.scale != 100),
        "orientation" -> ps.orientation,
        "fitToWidth" -> ps.fitToWidth.map(_.toString),
        "fitToHeight" -> ps.fitToHeight.map(_.toString)
      )
      val updated = overrides.foldLeft(base) { case (e, (key, value)) =>
        setOrRemoveAttr(e, key, value)
      }
      if existing.isEmpty && updated.attributes == Null then None else Some(updated)

/** CT_HeaderFooter child labels in schema sequence order (ECMA-376 18.3.1.46). */
private val headerFooterPartLabels: Seq[String] =
  Seq("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter")

/**
 * Merge the modeled header/footer into the preserved `<headerFooter>` element. All six page parts
 * (odd/even/first × header/footer) plus the differentOddEven/differentFirst flags are modeled
 * (GH-266); remaining unmodeled bits (scaleWithDoc/alignWithMargins attributes, foreign children)
 * stay intact. `headerFooter = None` preserves the source element verbatim; `Some(HeaderFooter())`
 * actively clears all modeled parts.
 */
private[ooxml] def mergeHeaderFooterElem(
  existing: Option[Elem],
  pageSetup: Option[PageSetup]
): Option[Elem] =
  pageSetup.flatMap(_.headerFooter) match
    case None => existing
    case Some(hf) =>
      val base = existing.getOrElse(XmlUtil.elem("headerFooter")())
      val unmodeled = base.child.filterNot {
        case e: Elem => headerFooterPartLabels.contains(e.label)
        case _ => false
      }
      val parts: Seq[Option[String]] = Seq(
        hf.oddHeader,
        hf.oddFooter,
        hf.evenHeader,
        hf.evenFooter,
        hf.firstHeader,
        hf.firstFooter
      )
      // CT_HeaderFooter children precede any foreign content — prepend in schema order
      val partElems: Seq[Node] = headerFooterPartLabels.zip(parts).flatMap { (label, text) =>
        text.map(t => XmlUtil.elem(label)(Text(t)))
      }
      val flagged = Seq(
        "differentOddEven" -> Option.when(hf.differentOddEven)("1"),
        "differentFirst" -> Option.when(hf.differentFirst)("1")
      ).foldLeft(base.copy(child = partElems ++ unmodeled)) { case (e, (key, value)) =>
        setOrRemoveAttr(e, key, value)
      }
      if flagged.child.isEmpty && flagged.attributes == Null then None else Some(flagged)

/**
 * Ensure `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>` when the model requests page-fit
 * scaling (GH-266): Excel ignores pageSetup's fitToWidth/fitToHeight without this flag. When no
 * fitTo* is modeled, the preserved element rides through verbatim — absence of the model fields is
 * not evidence the source flag should be cleared (fitToWidth/fitToHeight default to 1 in OOXML, so
 * a bare flag without pageSetup attributes is still meaningful).
 */
private[ooxml] def mergeSheetPrElem(
  existing: Option[Elem],
  pageSetup: Option[PageSetup]
): Option[Elem] =
  val wantsFit = pageSetup.exists(ps => ps.fitToWidth.isDefined || ps.fitToHeight.isDefined)
  if !wantsFit then existing
  else
    val base = existing.getOrElse(XmlUtil.elem("sheetPr")())
    val hasPageSetUpPr = base.child.exists {
      case e: Elem => e.label == "pageSetUpPr"
      case _ => false
    }
    val children =
      if hasPageSetUpPr then
        base.child.map {
          case e: Elem if e.label == "pageSetUpPr" =>
            e % new UnprefixedAttribute("fitToPage", "1", Null)
          case other => other
        }
      // CT_SheetPr's child sequence ends with pageSetUpPr — appending keeps schema order
      else base.child :+ XmlUtil.elem("pageSetUpPr", "fitToPage" -> "1")()
    Some(base.copy(child = children))

/**
 * Apply domain RowProperties to an OoxmlRow.
 *
 * Domain properties override existing row attributes (if any). This allows setting row height,
 * hidden state, and outline level from the domain model.
 */
private[ooxml] def applyDomainRowProps(row: OoxmlRow, props: RowProperties): OoxmlRow =
  row.copy(
    height = props.height.orElse(row.height),
    customHeight = props.height.isDefined || row.customHeight,
    hidden = props.hidden, // Domain always wins (allows unhide)
    outlineLevel = props.outlineLevel.orElse(row.outlineLevel),
    collapsed = props.collapsed // Domain always wins (allows uncollapse)
    // Note: styleId would need remapping to workbook-level index (deferred)
  )
