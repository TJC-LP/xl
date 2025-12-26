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
   * Update sheets while preserving all workbook metadata.
   *
   * Maps domain sheets to SheetRefs, preserving original sheetIds and visibility from the source
   * file. For new sheets, generates new IDs.
   */
  def updateSheets(newSheets: Vector[com.tjclp.xl.api.Sheet]): OoxmlWorkbook =
    val updatedRefs = newSheets.zipWithIndex.map { case (sheet, idx) =>
      // Try to find original SheetRef to preserve sheetId and visibility
      sheets.find(_.name == sheet.name) match
        case Some(original) =>
          // Preserve original ID and visibility, update relationship ID
          original.copy(relationshipId = s"rId${idx + 1}")
        case None =>
          // New sheet - generate new ID (use max existing ID + 1)
          val newId = sheets.map(_.sheetId).maxOption.getOrElse(0) + 1
          SheetRef(sheet.name, newId, s"rId${idx + 1}", None)
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

  /** Create workbook from domain model */
  def fromDomain(wb: Workbook): OoxmlWorkbook =
    val sheetRefs = wb.sheets.zipWithIndex.map { case (sheet, idx) =>
      SheetRef(sheet.name, idx + 1, s"rId${idx + 1}")
    }
    OoxmlWorkbook(sheetRefs)

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
