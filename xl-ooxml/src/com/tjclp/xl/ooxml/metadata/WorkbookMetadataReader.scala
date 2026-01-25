package com.tjclp.xl.ooxml.metadata

import java.io.InputStream
import java.nio.file.Path
import java.util.zip.ZipFile

import scala.xml.*

import com.tjclp.xl.addressing.{CellRange, SheetName}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.ooxml.XmlSecurity
import com.tjclp.xl.workbooks.DefinedName

/**
 * Lightweight metadata for instant workbook info without loading cell data.
 *
 * @param sheets
 *   Sheet info including name, ID, visibility, and dimension
 * @param definedNames
 *   Named ranges/formulas defined in the workbook
 */
final case class LightMetadata(
  sheets: Vector[SheetInfo],
  definedNames: Vector[DefinedName]
)

/**
 * Sheet information extracted from workbook metadata.
 *
 * @param name
 *   Sheet display name
 * @param sheetId
 *   Internal sheet ID
 * @param state
 *   Visibility: None (visible), Some("hidden"), Some("veryHidden")
 * @param dimension
 *   Used range from <dimension ref="..."/> element (may be inaccurate or missing)
 */
final case class SheetInfo(
  name: SheetName,
  sheetId: Int,
  state: Option[String],
  dimension: Option[CellRange]
)

/**
 * Lightweight workbook metadata reader.
 *
 * Parses only workbook.xml and worksheet dimension elements. No cell data is loaded. Suitable for
 * instant (<100ms) metadata queries on files of any size.
 */
object WorkbookMetadataReader:

  private val nsRelationships =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  /** Safely open a ZIP file, converting exceptions to XLResult errors. */
  private def openZipFile(path: Path): XLResult[ZipFile] =
    try Right(new ZipFile(path.toFile))
    catch
      case e: java.util.zip.ZipException =>
        Left(XLError.ParseError(path.toString, s"Invalid ZIP file: ${e.getMessage}"))
      case e: java.io.IOException =>
        Left(XLError.IOError(s"Failed to open file: ${e.getMessage}"))

  /**
   * Read lightweight metadata from XLSX file.
   *
   * Parses workbook.xml for sheets/names, and each worksheet for dimension element. Stops parsing
   * worksheets at <sheetData> for O(1) memory per sheet.
   *
   * @param path
   *   Path to XLSX file
   * @return
   *   LightMetadata with sheet info and defined names
   */
  def read(path: Path): XLResult[LightMetadata] =
    openZipFile(path).flatMap { zipFile =>
      try
        for
          // Parse workbook.xml
          wbXml <- readPart(zipFile, "xl/workbook.xml")
          wbElem <- XmlSecurity.parseSafe(wbXml, "xl/workbook.xml")

          // Parse workbook.xml.rels to map rId -> sheet path
          wbRels <- readPartOpt(zipFile, "xl/_rels/workbook.xml.rels")
          rIdMap = parseRelationships(wbRels)

          // Parse sheets from workbook.xml
          sheetRefs <- parseSheetRefs(wbElem)

          // For each sheet, read dimension from worksheet (stop at sheetData)
          sheetsWithDimensions <- sheetRefs.traverse { ref =>
            val sheetPath = rIdMap.getOrElse(ref.rId, s"xl/worksheets/sheet${ref.sheetId}.xml")
            val fullPath = if sheetPath.startsWith("xl/") then sheetPath else s"xl/$sheetPath"
            val dimension = readDimensionFromWorksheet(zipFile, fullPath)
            Right(
              SheetInfo(
                name = ref.name,
                sheetId = ref.sheetId,
                state = ref.state,
                dimension = dimension
              )
            )
          }
        yield
          val definedNames = parseDefinedNames(wbElem)
          LightMetadata(sheetsWithDimensions, definedNames)
      finally zipFile.close()
    }

  /**
   * Read only the sheet list (no dimensions, no defined names).
   *
   * Even faster than read() when only sheet names are needed.
   */
  def readSheetList(path: Path): XLResult[Vector[SheetInfo]] =
    openZipFile(path).flatMap { zipFile =>
      try
        for
          wbXml <- readPart(zipFile, "xl/workbook.xml")
          wbElem <- XmlSecurity.parseSafe(wbXml, "xl/workbook.xml")
          sheetRefs <- parseSheetRefs(wbElem)
        yield sheetRefs.map { ref =>
          SheetInfo(
            name = ref.name,
            sheetId = ref.sheetId,
            state = ref.state,
            dimension = None
          )
        }
      finally zipFile.close()
    }

  /**
   * Read only the defined names (no sheet info).
   */
  def readDefinedNames(path: Path): XLResult[Vector[DefinedName]] =
    openZipFile(path).flatMap { zipFile =>
      try
        for
          wbXml <- readPart(zipFile, "xl/workbook.xml")
          wbElem <- XmlSecurity.parseSafe(wbXml, "xl/workbook.xml")
        yield parseDefinedNames(wbElem)
      finally zipFile.close()
    }

  /**
   * Read dimension for a specific sheet by index (1-based).
   */
  def readDimension(path: Path, sheetIndex: Int): XLResult[Option[CellRange]] =
    openZipFile(path).flatMap { zipFile =>
      try
        val sheetPath = s"xl/worksheets/sheet$sheetIndex.xml"
        Right(readDimensionFromWorksheet(zipFile, sheetPath))
      finally zipFile.close()
    }

  // Internal sheet reference before dimension is added
  private case class SheetRefInternal(
    name: SheetName,
    sheetId: Int,
    rId: String,
    state: Option[String]
  )

  private def readPart(zipFile: ZipFile, entryName: String): XLResult[String] =
    Option(zipFile.getEntry(entryName)) match
      case None => Left(XLError.ParseError(entryName, s"Missing required part: $entryName"))
      case Some(entry) =>
        val is = zipFile.getInputStream(entry)
        try Right(new String(is.readAllBytes(), "UTF-8"))
        finally is.close()

  private def readPartOpt(zipFile: ZipFile, entryName: String): XLResult[Option[String]] =
    Option(zipFile.getEntry(entryName)) match
      case None => Right(None)
      case Some(entry) =>
        val is = zipFile.getInputStream(entry)
        try Right(Some(new String(is.readAllBytes(), "UTF-8")))
        finally is.close()

  private def parseRelationships(xmlOpt: Option[String]): Map[String, String] =
    xmlOpt match
      case None => Map.empty
      case Some(xml) =>
        XmlSecurity.parseSafe(xml, "workbook.xml.rels").toOption match
          case None => Map.empty
          case Some(elem) =>
            (elem \ "Relationship").collect { case e: Elem =>
              val id = e \@ "Id"
              val target = e \@ "Target"
              id -> target
            }.toMap

  private def parseSheetRefs(wbElem: Elem): XLResult[Vector[SheetRefInternal]] =
    (wbElem \ "sheets").headOption match
      case None => Left(XLError.ParseError("xl/workbook.xml", "Missing <sheets> element"))
      case Some(sheetsElem: Elem) =>
        val refs = (sheetsElem \ "sheet").collect { case e: Elem =>
          val name = e \@ "name"
          val sheetIdStr = e \@ "sheetId"
          val rId = e.attribute(nsRelationships, "id").map(_.text).getOrElse("")
          val state = Option(e \@ "state").filter(_.nonEmpty)

          for
            sheetName <- SheetName(name).left.map(err => XLError.InvalidSheetName(name, err))
            sheetId <- sheetIdStr.toIntOption.toRight(
              XLError.ParseError("xl/workbook.xml", s"Invalid sheetId: $sheetIdStr")
            )
          yield SheetRefInternal(sheetName, sheetId, rId, state)
        }.toVector

        refs.traverse(identity)
      case _ => Left(XLError.ParseError("xl/workbook.xml", "Invalid <sheets> element"))

  private def parseDefinedNames(wbElem: Elem): Vector[DefinedName] =
    (wbElem \ "definedNames").headOption match
      case None => Vector.empty
      case Some(elem) =>
        (elem \ "definedName").collect { case e: Elem =>
          val name = e \@ "name"
          val formula = e.text.trim
          val localSheetId = Option(e \@ "localSheetId").filter(_.nonEmpty).flatMap(_.toIntOption)
          val hidden = (e \@ "hidden") == "1"
          val comment = Option(e \@ "comment").filter(_.nonEmpty)
          DefinedName(
            name = name,
            formula = formula,
            localSheetId = localSheetId,
            hidden = hidden,
            comment = comment
          )
        }.toVector

  /**
   * Read <dimension ref="..."> from worksheet XML.
   *
   * Uses SAX parser to stop immediately after dimension element (or at sheetData). This is O(1)
   * memory and typically completes in <10ms regardless of worksheet size.
   */
  private def readDimensionFromWorksheet(zipFile: ZipFile, entryName: String): Option[CellRange] =
    Option(zipFile.getEntry(entryName)) match
      case None => None
      case Some(entry) =>
        val is = zipFile.getInputStream(entry)
        try DimensionExtractor.extract(is)
        finally is.close()

  /** Extension for traverse on Vector */
  extension [A](vec: Vector[A])
    private def traverse[B](f: A => XLResult[B]): XLResult[Vector[B]] =
      val builder = Vector.newBuilder[B]
      vec.foldLeft[XLResult[Unit]](Right(())) { (acc, a) =>
        for
          _ <- acc
          b <- f(a)
        yield builder += b
      } match
        case Right(_) => Right(builder.result())
        case Left(err) => Left(err)

/**
 * SAX-based dimension extractor.
 *
 * Parses worksheet XML just until <dimension> element, then aborts. This is much faster than
 * parsing the entire worksheet when only bounds are needed.
 */
private[metadata] object DimensionExtractor:
  import java.io.InputStream
  import javax.xml.parsers.SAXParserFactory
  import org.xml.sax.{Attributes, InputSource}
  import org.xml.sax.helpers.DefaultHandler
  import com.tjclp.xl.addressing.{ARef, CellRange}

  // Exception to abort SAX parsing early (no stacktrace for efficiency)
  @SuppressWarnings(Array("org.wartremover.warts.Null"))
  private final class FoundDimension(val range: CellRange)
      extends RuntimeException(null, null, false, false)

  @SuppressWarnings(Array("org.wartremover.warts.Null"))
  private final class NoDimension extends RuntimeException(null, null, false, false)

  /**
   * Extract dimension from worksheet XML using SAX parser.
   *
   * Stops parsing after <dimension> element or at start of <sheetData>. Memory: O(1), Time: ~10ms.
   */
  def extract(stream: InputStream): Option[CellRange] =
    try
      val factory = SAXParserFactory.newInstance()
      factory.setNamespaceAware(true)
      val parser = factory.newSAXParser()
      val handler = new DimensionHandler()
      parser.parse(InputSource(stream), handler)
      // If we get here, no dimension was found before end of document
      None
    catch
      case e: FoundDimension => Some(e.range)
      case _: NoDimension => None
      case _: Throwable => None

  private class DimensionHandler extends DefaultHandler:
    override def startElement(
      uri: String,
      localName: String,
      qName: String,
      attributes: Attributes
    ): Unit =
      localName match
        case "dimension" =>
          val ref = Option(attributes.getValue("ref"))
          ref.flatMap(parseRange) match
            case Some(range) => throw new FoundDimension(range)
            case None => () // Continue, maybe sheetData has data
        case "sheetData" =>
          // If we reached sheetData without finding dimension, abort
          throw new NoDimension
        case _ => ()

    private def parseRange(ref: String): Option[CellRange] =
      // Handle both single cell "A1" and range "A1:E1000000"
      if ref.contains(":") then
        val parts = ref.split(":", 2)
        for
          start <- ARef.parse(parts(0)).toOption
          end <- ARef.parse(parts(1)).toOption
        yield CellRange(start, end)
      else ARef.parse(ref).toOption.map(r => CellRange(r, r))
