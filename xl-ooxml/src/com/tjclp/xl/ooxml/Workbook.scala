package com.tjclp.xl.ooxml

import scala.xml.*
import XmlUtil.*
import SaxSupport.*

// Default namespace bindings for generated workbook parts. When reading from an existing workbook
// we capture the original scope (which may include mc/xr/x15, etc.) and reuse it to avoid Excel
// namespace warnings.
private val defaultWorkbookScope =
  NamespaceBinding(null, nsSpreadsheetML, NamespaceBinding("r", nsRelationships, TopScope))
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.api.Workbook
import com.tjclp.xl.workbooks.DefinedName

/**
 * Sheet reference in workbook.xml
 *
 * @param name
 *   Sheet name
 * @param sheetId
 *   Sheet ID (original, not renumbered - preserve 872, 871, 699, etc.)
 * @param relationshipId
 *   Relationship ID (r:id) linking to worksheet part
 * @param state
 *   Visibility: None (visible), Some("hidden"), Some("veryHidden")
 */
case class SheetRef(
  name: SheetName,
  sheetId: Int,
  relationshipId: String,
  state: Option[String] = None
)

/**
 * Workbook for xl/workbook.xml
 *
 * Contains sheet references and workbook-level properties. For surgical modification, preserves all
 * unparsed elements from the original workbook.xml (defined names, workbook properties, etc.) to
 * maintain Excel compatibility and prevent corruption warnings.
 *
 * @param sheets
 *   Sheet references (name, ID, visibility, relationship)
 * @param fileVersion
 *   File version metadata (appName, lastEdited, etc.)
 * @param workbookPr
 *   Workbook properties (codeName, defaultThemeVersion, etc.)
 * @param alternateContent
 *   mc:AlternateContent element (file path metadata)
 * @param revisionPtr
 *   xr:revisionPtr element (revision tracking)
 * @param bookViews
 *   Book views (window size, active tab, first visible sheet)
 * @param definedNames
 *   Defined names element with ALL named ranges (CRITICAL - thousands in real files)
 * @param calcPr
 *   Calculation properties (calcId, fullCalcOnLoad, etc.)
 * @param extLst
 *   Extensions list
 * @param otherElements
 *   Any other top-level elements not explicitly handled
 */
case class OoxmlWorkbook(
  sheets: Seq[SheetRef],
  fileVersion: Option[Elem] = None,
  workbookPr: Option[Elem] = None,
  alternateContent: Option[Elem] = None,
  revisionPtr: Option[Elem] = None,
  bookViews: Option[Elem] = None,
  definedNames: Option[Elem] = None,
  calcPr: Option[Elem] = None,
  extLst: Option[Elem] = None,
  otherElements: Seq[Elem] = Seq.empty,
  rootAttributes: MetaData = Null,
  rootScope: NamespaceBinding = defaultWorkbookScope
) extends XmlWritable,
      SaxSerializable:

  /**
   * True when this workbook uses the 1904 date system (`<workbookPr date1904="1"/>`, GH-243). Date
   * serials then count days since 1904-01-01 instead of the default 1900 system.
   */
  def date1904: Boolean = OoxmlWorkbook.parseDate1904(workbookPr)

  /**
   * Overlay the model's active tab onto the bookViews element (GH-294, model wins). See
   * [[OoxmlWorkbook.buildBookViews]] for the merge rules; the caller clamps.
   */
  def withActiveTab(activeTab: Int): OoxmlWorkbook =
    copy(bookViews = OoxmlWorkbook.buildBookViews(bookViews, activeTab))

  /**
   * Update sheets while preserving all workbook metadata.
   *
   * Maps domain sheets to SheetRefs, preserving original sheetIds from the source file. For new
   * sheets, generates new IDs. Applies state overrides from domain metadata.
   *
   * @param newSheets
   *   Domain sheets to update
   * @param stateOverrides
   *   Visibility overrides from domain metadata (SheetName -> state)
   */
  def updateSheets(
    newSheets: Vector[com.tjclp.xl.api.Sheet],
    stateOverrides: Map[SheetName, Option[String]] = Map.empty
  ): OoxmlWorkbook =
    updateSheets(newSheets, stateOverrides, OoxmlWorkbook.sequentialRelIds(newSheets.size))

  /**
   * Update sheets with EXPLICIT relationship ids (GH-320). The ids must come from the same
   * reconciliation pass that regenerates workbook.xml.rels — renumbering them independently
   * (rId1..N) corrupted foreign files whose preserved rels number sheets after the theme part
   * (LibreOffice maps rId1 to the theme, so the first sheet resolved to xl/theme/theme1.xml).
   */
  def updateSheets(
    newSheets: Vector[com.tjclp.xl.api.Sheet],
    stateOverrides: Map[SheetName, Option[String]],
    relationshipIds: Seq[String]
  ): OoxmlWorkbook =
    // Fresh sheetId allocation threads through the pass (seeded with the preserved max, which
    // already covers every matched sheet's id) so multiple unmatched sheets in one session —
    // multi-add, multi-rename — get DISTINCT ids; ECMA-376 requires sheetId uniqueness. Mirrors
    // the rId counter in Relationships.reconcileWorkbook.
    val seedMaxId = sheets.map(_.sheetId).maxOption.getOrElse(0)
    val (updatedRefs, _) = newSheets.zipWithIndex
      .foldLeft((Vector.empty[SheetRef], seedMaxId)) { case ((acc, maxId), (sheet, idx)) =>
        val relId = relationshipIds.lift(idx).getOrElse(s"rId${idx + 1}")
        // Check for explicit state override from domain metadata
        val overriddenState = stateOverrides.get(sheet.name)

        // Try to find original SheetRef to preserve sheetId
        sheets.find(_.name == sheet.name) match
          case Some(original) =>
            // Use override if present, otherwise preserve original state
            val finalState = overriddenState.getOrElse(original.state)
            (acc :+ original.copy(relationshipId = relId, state = finalState), maxId)
          case None =>
            // New sheet - allocate above every id seen so far, use override or default to visible
            val newId = maxId + 1
            val finalState = overriddenState.getOrElse(None)
            (acc :+ SheetRef(sheet.name, newId, relId, finalState), newId)
      }
    copy(sheets = updatedRefs)

  def toXml: Elem =
    val children = Seq.newBuilder[Node]

    // Emit elements in OOXML Part 1 schema order (critical for Excel compatibility)
    fileVersion.foreach(children += _)
    workbookPr.foreach(children += _)
    alternateContent.foreach(children += _)
    revisionPtr.foreach(children += _)
    bookViews.foreach(children += _)

    // Sheets element (regenerated with current names/order/visibility)
    // Use a scope with "r" namespace bound for r:id attributes
    val sheetScope = NamespaceBinding("r", nsRelationships, TopScope)
    val sheetElems = sheets.map { ref =>
      val baseAttrs = Seq(
        "name" -> ref.name.value,
        "sheetId" -> ref.sheetId.toString
      )
      // Add state attribute if sheet is hidden/veryHidden
      val stateAttr = ref.state.map(s => Seq("state" -> s)).getOrElse(Seq.empty)
      val allAttrs = baseAttrs ++ stateAttr

      // Build with r:id as namespaced attribute
      val rId = new PrefixedAttribute("r", "id", ref.relationshipId, Null)
      val attrs = allAttrs.foldRight(rId: MetaData) { case ((k, v), acc) =>
        new UnprefixedAttribute(k, v, acc)
      }
      Elem(null, "sheet", attrs, sheetScope, minimizeEmpty = true)
    }
    children += Elem(null, "sheets", Null, TopScope, minimizeEmpty = false, sheetElems*)

    // Defined names (CRITICAL - preserve all named ranges)
    definedNames.foreach(children += _)

    // Calculation properties
    calcPr.foreach(children += _)

    // Extensions
    extLst.foreach(children += _)

    // Any other elements
    otherElements.foreach(children += _)

    val scope = Option(rootScope).getOrElse(defaultWorkbookScope)
    val attrs = Option(rootAttributes).getOrElse(Null)

    Elem(null, "workbook", attrs, scope, minimizeEmpty = false, children.result()*)

  def writeSax(writer: SaxWriter): Unit =
    writer.startDocument()
    writer.startElement("workbook")

    val scope = Option(rootScope).getOrElse(defaultWorkbookScope)
    val attrs = Option(rootAttributes).getOrElse(Null)

    SaxWriter.withAttributes(writer, writer.combinedAttributes(scope, attrs)*) {
      fileVersion.foreach(writer.writeElem)
      workbookPr.foreach(writer.writeElem)
      alternateContent.foreach(writer.writeElem)
      revisionPtr.foreach(writer.writeElem)
      bookViews.foreach(writer.writeElem)

      writer.startElement("sheets")
      sheets.foreach { ref =>
        writer.startElement("sheet")
        val baseAttrs = Seq(
          "name" -> ref.name.value,
          "sheetId" -> ref.sheetId.toString
        )
        val stateAttrs = ref.state.map("state" -> _).toList
        val sheetAttrs = baseAttrs ++ stateAttrs
        SaxWriter.withAttributes(writer, sheetAttrs*) {
          writer.writeAttribute("r:id", ref.relationshipId)
        }
        writer.endElement()
      }
      writer.endElement() // sheets

      definedNames.foreach(writer.writeElem)
      calcPr.foreach(writer.writeElem)
      extLst.foreach(writer.writeElem)
      otherElements.foreach(writer.writeElem)
    }

    writer.endElement()
    writer.endDocument()
    writer.flush()

object OoxmlWorkbook extends XmlReadable[OoxmlWorkbook]:
  /** Create minimal workbook with one sheet */
  def minimal(sheetName: String = "Sheet1"): OoxmlWorkbook =
    OoxmlWorkbook(Seq(SheetRef(SheetName.unsafe(sheetName), 1, "rId1")))

  /** Sequential rId1..N — the fresh-workbook id scheme (matches Relationships.workbook). */
  private[ooxml] def sequentialRelIds(count: Int): Vector[String] =
    (1 to count).map(idx => s"rId$idx").toVector

  /** Create workbook from domain model */
  def fromDomain(wb: Workbook): OoxmlWorkbook =
    fromDomain(wb, None, sequentialRelIds(wb.sheets.size))

  /**
   * Create workbook.xml from the domain model, carrying preserved workbook-level elements through
   * when the source structure is available (GH-320 — the metadata-modified path).
   *
   * The model wins for everything it represents — sheets (names/order/visibility), activeTab,
   * definedNames, date1904 — while unmodeled preserved elements ride through in schema order:
   * fileVersion, workbookPr attributes, mc:AlternateContent, xr:revisionPtr, bookViews siblings,
   * calcPr, extLst, and otherElements (workbookProtection, pivotCaches, externalReferences, ...) —
   * the worksheet-level preserved-metadata machinery applied at workbook level. The relationship
   * ids come from the rels reconciliation pass (see [[Relationships.reconcileWorkbook]]) so every
   * emitted `r:id` resolves.
   */
  def fromDomain(
    wb: Workbook,
    preserved: Option[OoxmlWorkbook],
    relationshipIds: Seq[String]
  ): OoxmlWorkbook =
    preserved match
      case Some(p) =>
        p.updateSheets(wb.sheets, wb.metadata.sheetStates, relationshipIds)
          // GH-294: model activeTab overlaid on the preserved bookViews (siblings ride through)
          .withActiveTab(clampActiveTab(wb.activeSheetIndex, wb.sheets.size))
          .copy(
            workbookPr = reconcileDate1904(p.workbookPr, wb.metadata.date1904),
            definedNames = reconcileDefinedNames(p.definedNames, PrintNames.effective(wb))
          )
      case None =>
        val sheetRefs = wb.sheets.zipWithIndex.map { case (sheet, idx) =>
          val state = wb.metadata.sheetStates.get(sheet.name).flatten
          val relId = relationshipIds.lift(idx).getOrElse(s"rId${idx + 1}")
          SheetRef(sheet.name, idx + 1, relId, state)
        }
        // GH-243: preserve-the-system — declare the 1904 date system when the model says so, so the
        // raw serials riding through (and DateTime cells serialized with the 1904 epoch) stay
        // correct.
        val workbookPr =
          if wb.metadata.date1904 then Some(elem("workbookPr", "date1904" -> "1")()) else None
        // GH-236: serialize named ranges from the typed model (previously dropped on write).
        // GH-259: print area / repeat rows live on Sheet.pageSetup and are appended here as
        // sheet-scoped _xlnm.Print_Area / _xlnm.Print_Titles names.
        OoxmlWorkbook(
          sheetRefs,
          workbookPr = workbookPr,
          // GH-294: fresh workbooks always ship bookViews/activeTab (Excel always writes bookViews)
          bookViews = buildBookViews(None, clampActiveTab(wb.activeSheetIndex, wb.sheets.size)),
          definedNames = buildDefinedNames(PrintNames.effective(wb))
        )

  /**
   * Keep the preserved `<definedNames>` bytes when the model agrees with them; regenerate the
   * element from the model otherwise (the GH-259 reconcile, shared by both writer branches).
   */
  def reconcileDefinedNames(
    preservedElem: Option[Elem],
    expected: Vector[DefinedName]
  ): Option[Elem] =
    if parseDefinedNames(preservedElem).toSet == expected.toSet then preservedElem
    else buildDefinedNames(expected)

  /**
   * Reconcile the date1904 declaration with the model (model wins, GH-243) while every other
   * preserved workbookPr attribute (codeName, defaultThemeVersion, ...) rides through.
   * Byte-identical when the model agrees with the preserved spelling.
   */
  def reconcileDate1904(preservedPr: Option[Elem], date1904: Boolean): Option[Elem] =
    if parseDate1904(preservedPr) == date1904 then preservedPr
    else
      (preservedPr, date1904) match
        case (Some(pr), true) => Some(pr % new UnprefixedAttribute("date1904", "1", Null))
        case (Some(pr), false) => Some(pr.copy(attributes = pr.attributes.remove("date1904")))
        case (None, true) => Some(elem("workbookPr", "date1904" -> "1")())
        case (None, false) => None

  /** Clamp an active-tab index into [0, sheetCount-1] (write side; the reader clamps too). */
  def clampActiveTab(index: Int, sheetCount: Int): Int =
    index.max(0).min((sheetCount - 1).max(0))

  /**
   * Parse `<bookViews>/<workbookView activeTab>` (first view wins, like Excel). Returns None when
   * the element or attribute is absent — the schema default is 0.
   */
  def parseActiveTab(bookViews: Option[Elem]): Option[Int] =
    for
      bv <- bookViews
      view <- (bv \ "workbookView").collectFirst { case e: Elem => e }
      value <- getAttrOpt(view, "activeTab")
      tab <- value.toIntOption
    yield tab

  /**
   * Overlay the model's activeTab onto a (possibly preserved) `<bookViews>` element — the model
   * value wins (GH-294, docProps GH-242 precedent) while unmodeled attributes (window geometry,
   * tabRatio, ...) ride through:
   *   - no element / no `<workbookView>` child: a fresh `<workbookView activeTab="N"/>` is added
   *     (always, including N=0 — Excel always writes bookViews)
   *   - existing first `<workbookView>`: activeTab is set to N; when N=0 and the source omitted the
   *     attribute it stays omitted (0 is the schema default), so parse→merge is an identity for
   *     both the Excel (omitted) and openpyxl (explicit ="0") spellings
   */
  def buildBookViews(preserved: Option[Elem], activeTab: Int): Option[Elem] =
    val base = preserved.getOrElse(elem("bookViews")())
    val firstViewIdx = base.child.indexWhere {
      case e: Elem => e.label == "workbookView"
      case _ => false
    }
    val children =
      if firstViewIdx < 0 then
        base.child :+ elem("workbookView", "activeTab" -> activeTab.toString)()
      else
        base.child.zipWithIndex.map {
          case (e: Elem, idx) if idx == firstViewIdx =>
            if activeTab > 0 || getAttrOpt(e, "activeTab").isDefined then
              e % new UnprefixedAttribute("activeTab", activeTab.toString, Null)
            else e
          case (other, _) => other
        }
    Some(base.copy(child = children))

  /**
   * Build a `<definedNames>` element from the typed model (GH-236), or None when empty. Order
   * follows the model (preserving the source order on round-trips); attributes are emitted in a
   * fixed order for deterministic output.
   */
  def buildDefinedNames(names: Vector[DefinedName]): Option[Elem] =
    if names.isEmpty then None
    else
      val children = names.map { dn =>
        val attrs = Seq.newBuilder[(String, String)]
        attrs += ("name" -> dn.name)
        dn.comment.foreach(c => attrs += ("comment" -> c))
        dn.localSheetId.foreach(id => attrs += ("localSheetId" -> id.toString))
        if dn.hidden then attrs += ("hidden" -> "1")
        elemOrdered("definedName", attrs.result()*)(Text(dn.formula))
      }
      Some(elem("definedNames")(children*))

  /**
   * Parse the `date1904` attribute of a raw `<workbookPr>` element (GH-243). Single source of truth
   * shared by the full reader and the lightweight metadata reader. OOXML uses xsd:boolean, so both
   * "1" and "true" mean the 1904 date system; absent attribute/element means 1900.
   */
  def parseDate1904(workbookPr: Option[Elem]): Boolean =
    workbookPr.exists { e =>
      val value = e \@ "date1904"
      value == "1" || value == "true"
    }

  /**
   * Parse `<definedNames>` into the typed model. Single source of truth shared by the reader and
   * the surgical writer (so the writer can detect whether the model is unchanged and keep raw
   * bytes).
   */
  def parseDefinedNames(rawElem: Option[Elem]): Vector[DefinedName] =
    rawElem match
      case None => Vector.empty
      case Some(dnsElem) =>
        (dnsElem \ "definedName").collect { case e: Elem =>
          DefinedName(
            name = e \@ "name",
            formula = e.text.trim,
            localSheetId = Option(e \@ "localSheetId").filter(_.nonEmpty).flatMap(_.toIntOption),
            hidden = (e \@ "hidden") == "1",
            comment = Option(e \@ "comment").filter(_.nonEmpty)
          )
        }.toVector

  def fromXml(elem: Elem): Either[String, OoxmlWorkbook] =
    for
      // Parse sheets (required element)
      sheetsElem <- getChild(elem, "sheets")
      sheetElems = getChildren(sheetsElem, "sheet")
      sheets <- parseSheets(sheetElems)

      // Extract all preserved elements (optional)
      fileVersion = (elem \ "fileVersion").headOption.collect { case e: Elem => e }
      workbookPr = (elem \ "workbookPr").headOption.collect { case e: Elem => e }

      // Handle mc:AlternateContent (may have namespace prefix)
      alternateContent = elem.child.collectFirst {
        case e: Elem if e.label == "AlternateContent" => e
      }

      // Handle xr:revisionPtr (may have namespace prefix)
      revisionPtr = elem.child.collectFirst {
        case e: Elem if e.label == "revisionPtr" => e
      }

      bookViews = (elem \ "bookViews").headOption.collect { case e: Elem => e }
      definedNames = (elem \ "definedNames").headOption.collect { case e: Elem => e }
      calcPr = (elem \ "calcPr").headOption.collect { case e: Elem => e }
      extLst = (elem \ "extLst").headOption.collect { case e: Elem => e }

      // Collect any other top-level elements we don't explicitly handle
      knownElements = Set(
        "fileVersion",
        "workbookPr",
        "AlternateContent",
        "revisionPtr",
        "bookViews",
        "sheets",
        "definedNames",
        "calcPr",
        "extLst"
      )
      otherElements = elem.child.collect {
        case e: Elem if !knownElements.contains(e.label) => e
      }
    yield OoxmlWorkbook(
      sheets,
      fileVersion,
      workbookPr,
      alternateContent,
      revisionPtr,
      bookViews,
      definedNames,
      calcPr,
      extLst,
      otherElements.toSeq,
      rootAttributes = elem.attributes,
      rootScope = Option(elem.scope).getOrElse(defaultWorkbookScope)
    )

  private def parseSheets(elems: Seq[Elem]): Either[String, Seq[SheetRef]] =
    val parsed = elems.map { e =>
      for
        name <- getAttr(e, "name")
        sheetIdStr <- getAttr(e, "sheetId")
        sheetId <- sheetIdStr.toIntOption.toRight(s"Invalid sheetId: $sheetIdStr")
        rId <- e.attribute(nsRelationships, "id").map(_.text).toRight("Missing r:id")
        state = getAttrOpt(e, "state") // Extract visibility: "hidden" | "veryHidden"
      yield SheetRef(SheetName.unsafe(name), sheetId, rId, state)
    }

    val errors = parsed.collect { case Left(err) => err }
    if errors.nonEmpty then Left(s"Sheet parse errors: ${errors.mkString(", ")}")
    else Right(parsed.collect { case Right(ref) => ref })
