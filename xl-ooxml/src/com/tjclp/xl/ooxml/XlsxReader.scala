package com.tjclp.xl.ooxml

import scala.xml.*
import java.io.{ByteArrayInputStream, FileInputStream, InputStream}
import java.util.zip.ZipInputStream
import java.nio.file.Path
import scala.collection.mutable
import com.tjclp.xl.{Workbook, Sheet, Cell, CellValue, XLError, XLResult, SheetName}

/** Reader for XLSX files (ZIP parsing)
  *
  * Parses XLSX ZIP structure and converts to domain Workbook.
  */
object XlsxReader:

  /** Read workbook from XLSX file */
  def read(inputPath: Path): XLResult[Workbook] =
    try
      val is = new FileInputStream(inputPath.toFile)
      try
        readFromStream(is)
      finally
        is.close()
    catch
      case e: Exception => Left(XLError.IOError(s"Failed to read XLSX: ${e.getMessage}"))

  /** Read workbook from byte array (for testing) */
  def readFromBytes(bytes: Array[Byte]): XLResult[Workbook] =
    try
      readFromStream(new ByteArrayInputStream(bytes))
    catch
      case e: Exception => Left(XLError.IOError(s"Failed to read bytes: ${e.getMessage}"))

  /** Read workbook from input stream */
  private def readFromStream(is: InputStream): XLResult[Workbook] =
    val zip = new ZipInputStream(is)
    val parts = mutable.Map[String, String]()

    try
      // Read all ZIP entries
      var entry = zip.getNextEntry
      while entry != null do
        if !entry.isDirectory then
          val content = new String(zip.readAllBytes(), "UTF-8")
          parts(entry.getName) = content
        zip.closeEntry()
        entry = zip.getNextEntry

      // Parse workbook structure
      parseWorkbook(parts.toMap)

    finally
      zip.close()

  /** Parse workbook from collected parts */
  private def parseWorkbook(parts: Map[String, String]): XLResult[Workbook] =
    for
      // Parse workbook.xml
      workbookXml <- parts.get("xl/workbook.xml")
        .toRight(XLError.ParseError("xl/workbook.xml", "Missing workbook.xml"))
      workbookElem <- parseXml(workbookXml, "xl/workbook.xml")
      ooxmlWb <- OoxmlWorkbook.fromXml(workbookElem)
        .left.map(err => XLError.ParseError("xl/workbook.xml", err): XLError)

      // Parse optional shared strings
      sst <- parseOptionalSST(parts)

      // Parse sheets
      sheets <- parseSheets(parts, ooxmlWb.sheets, sst)

      // Assemble workbook
      workbook <- assembleWorkbook(sheets)

    yield workbook

  /** Parse optional shared strings table */
  private def parseOptionalSST(parts: Map[String, String]): XLResult[Option[SharedStrings]] =
    parts.get("xl/sharedStrings.xml") match
      case None => Right(None)
      case Some(xml) =>
        for
          elem <- parseXml(xml, "xl/sharedStrings.xml")
          sst <- SharedStrings.fromXml(elem)
            .left.map(err => XLError.ParseError("xl/sharedStrings.xml", err): XLError)
        yield Some(sst)

  /** Parse all worksheets */
  private def parseSheets(
    parts: Map[String, String],
    sheetRefs: Seq[SheetRef],
    sst: Option[SharedStrings]
  ): XLResult[Vector[Sheet]] =
    sheetRefs.toVector.traverse { ref =>
      val sheetPath = s"xl/worksheets/sheet${ref.sheetId}.xml"
      for
        xml <- parts.get(sheetPath)
          .toRight(XLError.ParseError(sheetPath, s"Missing worksheet: $sheetPath"))
        elem <- parseXml(xml, sheetPath)
        ooxmlSheet <- OoxmlWorksheet.fromXml(elem)
          .left.map(err => XLError.ParseError(sheetPath, err): XLError)
        domainSheet <- convertToDomainSheet(ref.name, ooxmlSheet, sst)
      yield domainSheet
    }

  /** Convert OoxmlWorksheet to domain Sheet */
  private def convertToDomainSheet(
    name: SheetName,
    ooxmlSheet: OoxmlWorksheet,
    sst: Option[SharedStrings]
  ): XLResult[Sheet] =
    val cells = ooxmlSheet.rows.flatMap { row =>
      row.cells.map { ooxmlCell =>
        // Resolve SST index if needed
        val value = (ooxmlCell.cellType, ooxmlCell.value, sst) match
          case ("s", CellValue.Text(idxStr), Some(sharedStrings)) =>
            // Resolve SST index
            idxStr.toIntOption match
              case Some(idx) => sharedStrings(idx) match
                case Some(text) => CellValue.Text(text)
                case None => CellValue.Error(com.tjclp.xl.CellError.Ref)
              case None => CellValue.Error(com.tjclp.xl.CellError.Value)
          case _ => ooxmlCell.value

        Cell(ooxmlCell.ref, value, ooxmlCell.styleIndex)
      }
    }

    Right(Sheet(
      name = name,
      cells = cells.map(c => c.ref -> c).toMap
    ))

  /** Assemble final workbook */
  private def assembleWorkbook(sheets: Vector[Sheet]): XLResult[Workbook] =
    if sheets.isEmpty then
      Left(XLError.InvalidWorkbook("Workbook must have at least one sheet"))
    else
      Right(Workbook(sheets = sheets))

  /** Parse XML string to Elem */
  private def parseXml(xmlString: String, location: String): XLResult[Elem] =
    try
      Right(XML.loadString(xmlString))
    catch
      case e: Exception =>
        Left(XLError.ParseError(location, s"XML parse error: ${e.getMessage}"))

  /** Extension for traverse on Vector */
  extension [A](vec: Vector[A])
    private def traverse[B](f: A => XLResult[B]): XLResult[Vector[B]] =
      vec.foldLeft[XLResult[Vector[B]]](Right(Vector.empty)) { (acc, a) =>
        for
          bs <- acc
          b <- f(a)
        yield bs :+ b
      }
