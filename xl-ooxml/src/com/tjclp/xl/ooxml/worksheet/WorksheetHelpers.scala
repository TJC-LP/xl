package com.tjclp.xl.ooxml.worksheet

import scala.xml.*

import com.tjclp.xl.addressing.{ARef, Column}
import com.tjclp.xl.ooxml.XmlUtil
import com.tjclp.xl.ooxml.XmlUtil.nsSpreadsheetML
import com.tjclp.xl.sheets.{
  ColumnProperties,
  FreezePane,
  PageSetup,
  RowProperties,
  Sheet,
  SheetView
}

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

// ===== Namespace preservation on regeneration (GH-291) =====
// Some producers (openpyxl) bind a prefix locally on a child element — e.g.
// `<drawing xmlns:r="…" r:id="rId1"/>` — instead of on the worksheet root. Regeneration strips
// child scopes (cleanNamespaces) to avoid redundant re-declarations, so a binding that exists
// ONLY on a child must be re-bound or the emitted prefix is unbound and Excel reports the
// workbook as corrupt. The reader captures each preserved subtree with the bindings it actually
// uses (rebindUsedNamespaces); emission hoists any binding the root lacks onto the regenerated
// worksheet root (hoistUsedBindings) — the standard Excel layout.

/**
 * Well-known OOXML prefixes, used as a last resort when a used prefix's binding was lost before
 * capture (mirrors StaxSaxWriter.knownNamespaces, which heals the same way on the StAX path).
 */
private val wellKnownNamespaces: Map[String, String] = Map(
  "r" -> XmlUtil.nsRelationships,
  "mc" -> "http://schemas.openxmlformats.org/markup-compatibility/2006",
  "xr" -> "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
  "x14ac" -> "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
)

/**
 * Prefix -> URI for every namespace prefix USED by the tree (on element tags and attributes),
 * resolved against the scope in effect at the point of use, falling back to wellKnownNamespaces.
 * First binding wins (document order); unresolvable prefixes are omitted (the input was already
 * namespace-broken). `xml` and `xmlns` are never collected. Attributes built via XmlUtil.elem carry
 * the prefix inside the key ("x14ac:dyDescent"), so colon-keys are treated as prefixed. Parsed
 * trees carry full scope chains on every element; for programmatically built trees (child scope
 * TopScope) the nearest ancestor scope is consulted instead.
 */
private[ooxml] def usedPrefixBindings(root: Elem): Seq[(String, String)] =
  @annotation.tailrec
  def attrPrefixes(md: MetaData, acc: List[String]): List[String] = md match
    case Null => acc.reverse
    case PrefixedAttribute(pre, _, _, next) => attrPrefixes(next, pre :: acc)
    case UnprefixedAttribute(key, _, next) =>
      key.split(":", 2) match
        case Array(pre, _) => attrPrefixes(next, pre :: acc)
        case _ => attrPrefixes(next, acc)
  def collect(e: Elem, inherited: NamespaceBinding): Vector[(String, String)] =
    val own = Option(e.scope).filterNot(_ == TopScope)
    val effectiveScope = own.getOrElse(inherited)
    val prefixes = (Option(e.prefix).toList ++ attrPrefixes(e.attributes, Nil))
      .filter(p => p.nonEmpty && p != "xml" && p != "xmlns")
    val here = prefixes.flatMap { p =>
      Option(effectiveScope.getURI(p))
        .orElse(wellKnownNamespaces.get(p))
        .map(p -> _)
    }
    here.toVector ++
      e.child.collect { case c: Elem => c }.toVector.flatMap(collect(_, effectiveScope))
  collect(root, TopScope).distinctBy(_._1)

/**
 * cleanNamespaces, then re-bind the prefixes the subtree actually uses on the subtree root, so a
 * preserved element stays namespace-well-formed in isolation (GH-291). Bindings are attached in
 * sorted-prefix order for deterministic output.
 */
private[ooxml] def rebindUsedNamespaces(elem: Elem): Elem =
  rebindUsedNamespaces(elem, includeDefault = false)

/**
 * The default-namespace URI the subtree's unprefixed elements resolve to, taken from the scope in
 * effect at the first unprefixed element in document order (GH-221: preserved drawing anchors are
 * unprefixed in the openpyxl dialect, so their default binding must travel with the fragment for it
 * to be scope-self-contained). None when every element is prefixed or the binding is absent.
 */
private[ooxml] def usedDefaultNamespace(root: Elem): Option[String] =
  def find(e: Elem, inherited: NamespaceBinding): Option[String] =
    val effective = Option(e.scope).filterNot(_ == TopScope).getOrElse(inherited)
    val here =
      if e.prefix == null || e.prefix.isEmpty then Option(effective.getURI(null))
      else None
    here.orElse {
      e.child.collect { case c: Elem => c }.foldLeft(Option.empty[String]) { (acc, c) =>
        acc.orElse(find(c, effective))
      }
    }
  find(root, TopScope)

/**
 * rebindUsedNamespaces extended to also re-bind the DEFAULT namespace used by unprefixed elements
 * (GH-221). Used for preserved drawing fragments, which must remain namespace-self-contained as
 * standalone strings; the worksheet path keeps the prefix-only variant (the worksheet's default
 * namespace is structural and never travels with fragments).
 */
private[ooxml] def rebindUsedNamespaces(elem: Elem, includeDefault: Boolean): Elem =
  val used = usedPrefixBindings(elem)
  val default = if includeDefault then usedDefaultNamespace(elem) else None
  val cleaned = cleanNamespaces(elem)
  if used.isEmpty && default.isEmpty then cleaned
  else
    val base = default.fold(TopScope: NamespaceBinding)(NamespaceBinding(null, _, TopScope))
    val scope = used.sortBy(_._1).foldRight(base) { case ((prefix, uri), parent) =>
      NamespaceBinding(prefix, uri, parent)
    }
    cleaned.copy(scope = scope)

/**
 * Extend `base` with bindings for prefixes used by `children` but not bound on `base` — hoisting
 * locally-declared child bindings onto the regenerated worksheet root, the standard Excel layout
 * (GH-291). Missing prefixes are prepended in sorted order for deterministic output. A prefix
 * already bound on `base` (even to a different URI) is left untouched.
 */
private[ooxml] def hoistUsedBindings(
  base: NamespaceBinding,
  children: Seq[Elem]
): NamespaceBinding =
  children
    .flatMap(usedPrefixBindings)
    .distinctBy(_._1)
    .filterNot { case (prefix, _) => scopeHasPrefix(base, prefix) }
    .sortBy(_._1)
    .foldRight(base) { case ((prefix, uri), parent) => NamespaceBinding(prefix, uri, parent) }

/**
 * Ensure a single well-known prefix is bound on `scope` (no-op for unknown prefixes). Rows re-emit
 * `x14ac:dyDescent` as a raw qualified attribute outside the Elem-based collection, so
 * OoxmlWorksheet binds that prefix explicitly when any row carries the attribute (GH-291).
 */
private[worksheet] def ensureWellKnownPrefix(
  scope: NamespaceBinding,
  prefix: String
): NamespaceBinding =
  if scopeHasPrefix(scope, prefix) then scope
  else wellKnownNamespaces.get(prefix).fold(scope)(uri => NamespaceBinding(prefix, uri, scope))

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
    // External entries use r:id — carry the binding so emission can hoist it (GH-291)
    val scope =
      if entries.exists(_.external) then NamespaceBinding("r", XmlUtil.nsRelationships, TopScope)
      else TopScope
    Some(Elem(null, "hyperlinks", Null, scope, minimizeEmpty = false, children*))

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

// ===== Sheet view authoring (GH-258) =====
// Shared by the DOM writer (OoxmlWorksheet) and the streaming writer (DirectSaxEmitter, GH-265)
// so freeze panes and view settings serialize identically on both paths.

/**
 * Build sheetViews XML from the modeled freeze pane and view settings (GH-258), overlaying any
 * preserved `<sheetViews>` element.
 *
 * The freeze pane override runs first (creating a minimal `<sheetViews><sheetView/></sheetViews>`
 * when needed), then view settings are merged onto the same `<sheetView>` element, so freeze panes
 * and view settings always share ONE sheetView.
 */
private[ooxml] def buildSheetViewsElem(
  existing: Option[Elem],
  freezeOverride: Option[FreezePane],
  viewSettings: Option[SheetView]
): Option[Elem] =
  applyViewSettingsOverride(applyFreezePaneOverride(existing, freezeOverride), viewSettings)

/**
 * Apply sheet view settings to sheetViews XML.
 *
 * When `None`, preserves existing sheetViews unchanged (passive default, mirroring freezePane).
 * When `Some(view)`, sets `showGridLines` ("1"/"0", always written so the setting round-trips) and
 * `zoomScale` (written when defined, removed when None) on every `<sheetView>` element, creating a
 * minimal `<sheetViews><sheetView workbookViewId="0"/></sheetViews>` when absent. Unmodeled
 * attributes (tabSelected, topLeftCell, view, ...) are preserved.
 */
private def applyViewSettingsOverride(
  existing: Option[Elem],
  viewSettings: Option[SheetView]
): Option[Elem] =
  viewSettings match
    case None => existing
    case Some(view) =>
      val base = existing.getOrElse(
        XmlUtil.elem("sheetViews")(XmlUtil.elem("sheetView", "workbookViewId" -> "0")())
      )
      // Degenerate guard: a <sheetViews> without any <sheetView> child gets one appended
      val hasSheetView = base.child.exists {
        case e: Elem => e.label == "sheetView"
        case _ => false
      }
      val withView =
        if hasSheetView then base
        else base.copy(child = base.child :+ XmlUtil.elem("sheetView", "workbookViewId" -> "0")())
      val newChildren = withView.child.map {
        case e: Elem if e.label == "sheetView" => applyViewAttrs(e, view)
        case other => other
      }
      Some(withView.copy(child = newChildren))

/** Set the modeled view attributes on a single `<sheetView>` element. */
private def applyViewAttrs(sheetView: Elem, view: SheetView): Elem =
  val withGrid = sheetView % new UnprefixedAttribute(
    "showGridLines",
    if view.showGridLines then "1" else "0",
    Null
  )
  view.zoomScale match
    case Some(zoom) => withGrid % new UnprefixedAttribute("zoomScale", zoom.toString, Null)
    case None => withGrid.copy(attributes = withGrid.attributes.remove("zoomScale"))

/**
 * Apply freeze pane override to sheetViews XML.
 *
 * When `freezeOverride` is `Some(FreezePane.At(ref))`, injects a `<pane>` element. When
 * `Some(FreezePane.Remove)`, strips `<pane>` from existing sheetViews. When `None`, preserves
 * existing sheetViews unchanged.
 */
private def applyFreezePaneOverride(
  existing: Option[Elem],
  freezeOverride: Option[FreezePane]
): Option[Elem] =
  freezeOverride match
    case None => existing
    case Some(FreezePane.Remove) =>
      existing.map { sv =>
        val newChildren = sv.child.map {
          case e: Elem if e.label == "sheetView" =>
            e.copy(child = e.child.filterNot {
              case c: Elem => c.label == "pane"
              case _ => false
            })
          case other => other
        }
        sv.copy(child = newChildren)
      }
    case Some(FreezePane.At(ref)) =>
      val colSplit = ref.col.index0
      val rowSplit = ref.row.index0
      if colSplit == 0 && rowSplit == 0 then existing
      else
        val activePane = (colSplit > 0, rowSplit > 0) match
          case (true, true) => "bottomRight"
          case (false, true) => "bottomLeft"
          case (true, false) => "topRight"
          case _ => "bottomLeft"

        val paneAttrs: Vector[(String, String)] =
          (if colSplit > 0 then Vector("xSplit" -> colSplit.toString) else Vector.empty) ++
            (if rowSplit > 0 then Vector("ySplit" -> rowSplit.toString) else Vector.empty) ++
            Vector(
              "topLeftCell" -> ref.toA1,
              "activePane" -> activePane,
              "state" -> "frozen"
            )

        val paneElem = XmlUtil.elem("pane", paneAttrs*)()

        // Inject pane into existing sheetViews or create new
        existing match
          case Some(sv) =>
            val newChildren = sv.child.map {
              case e: Elem if e.label == "sheetView" =>
                // Remove old pane, add new one at the beginning (before selection elements)
                val withoutPane = e.child.filterNot {
                  case c: Elem => c.label == "pane"
                  case _ => false
                }
                e.copy(child = paneElem +: withoutPane)
              case other => other
            }
            Some(sv.copy(child = newChildren))
          case None =>
            // Create minimal sheetViews from scratch
            val sheetView = XmlUtil.elem(
              "sheetView",
              "workbookViewId" -> "0"
            )(paneElem)
            Some(XmlUtil.elem("sheetViews")(sheetView))

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
 * Ensure `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>` when the model requests page-fit scaling
 * (GH-266): Excel ignores pageSetup's fitToWidth/fitToHeight without this flag. When no fitTo* is
 * modeled, the preserved element rides through verbatim — absence of the model fields is not
 * evidence the source flag should be cleared (fitToWidth/fitToHeight default to 1 in OOXML, so a
 * bare flag without pageSetup attributes is still meaningful).
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
