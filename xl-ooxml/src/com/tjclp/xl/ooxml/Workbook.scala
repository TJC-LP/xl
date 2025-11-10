package com.tjclp.xl.ooxml

import scala.xml.*
import XmlUtil.*
import com.tjclp.xl.SheetName

/** Sheet reference in workbook.xml */
case class SheetRef(
  name: SheetName,
  sheetId: Int,
  relationshipId: String  // r:id
)

/** Workbook for xl/workbook.xml
  *
  * Contains sheet references and workbook-level properties.
  * Each sheet has a name, sheetId (1-based), and r:id pointing to the worksheet part.
  */
case class OoxmlWorkbook(
  sheets: Seq[SheetRef]
) extends XmlWritable:

  def toXml: Elem =
    val sheetsElem = Elem(
      null, "sheets", Null, TopScope, minimizeEmpty = false,
      sheets.sortBy(_.sheetId).map { ref =>
        val rId = new PrefixedAttribute("r", "id", ref.relationshipId, Null)
        val attrs = new UnprefixedAttribute("name", ref.name.value,
          new UnprefixedAttribute("sheetId", ref.sheetId.toString, rId))
        Elem(null, "sheet", attrs, TopScope, minimizeEmpty = true)
      }*
    )

    val nsBindings = NamespaceBinding(null, nsSpreadsheetML,
      NamespaceBinding("r", nsRelationships, TopScope))

    Elem(null, "workbook", Null, nsBindings, minimizeEmpty = false, sheetsElem)

object OoxmlWorkbook extends XmlReadable[OoxmlWorkbook]:
  /** Create minimal workbook with one sheet */
  def minimal(sheetName: String = "Sheet1"): OoxmlWorkbook =
    OoxmlWorkbook(Seq(SheetRef(SheetName.unsafe(sheetName), 1, "rId1")))

  /** Create workbook from domain model */
  def fromDomain(wb: com.tjclp.xl.Workbook): OoxmlWorkbook =
    val sheetRefs = wb.sheets.zipWithIndex.map { case (sheet, idx) =>
      SheetRef(sheet.name, idx + 1, s"rId${idx + 1}")
    }
    OoxmlWorkbook(sheetRefs)

  def fromXml(elem: Elem): Either[String, OoxmlWorkbook] =
    for
      sheetsElem <- getChild(elem, "sheets")
      sheetElems = getChildren(sheetsElem, "sheet")
      sheets <- parseSheets(sheetElems)
    yield OoxmlWorkbook(sheets)

  private def parseSheets(elems: Seq[Elem]): Either[String, Seq[SheetRef]] =
    val parsed = elems.map { e =>
      for
        name <- getAttr(e, "name")
        sheetIdStr <- getAttr(e, "sheetId")
        sheetId <- sheetIdStr.toIntOption.toRight(s"Invalid sheetId: $sheetIdStr")
        rId <- e.attribute(nsRelationships, "id").map(_.text).toRight("Missing r:id")
      yield SheetRef(SheetName.unsafe(name), sheetId, rId)
    }

    val errors = parsed.collect { case Left(err) => err }
    if errors.nonEmpty then
      Left(s"Sheet parse errors: ${errors.mkString(", ")}")
    else
      Right(parsed.collect { case Right(ref) => ref })
