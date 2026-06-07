package com.tjclp.xl.ooxml.worksheet

import scala.xml.*

import com.tjclp.xl.addressing.Column
import com.tjclp.xl.ooxml.XmlUtil
import com.tjclp.xl.ooxml.XmlUtil.nsSpreadsheetML
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties, Sheet}

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
private[worksheet] val preservedKnownLabels: Set[String] =
  (preservedAfterSheetData ++ preservedAfterMergeCells ++ preservedAfterCondFmt ++
    preservedAfterCustomProps ++ preservedAfterLegacyDrawing ++ preservedAfterControls).toSet

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
