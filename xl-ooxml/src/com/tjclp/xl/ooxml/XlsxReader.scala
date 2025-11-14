package com.tjclp.xl.ooxml

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cell.{Cell, CellError, CellValue}
import com.tjclp.xl.api.{Sheet, Workbook}
import com.tjclp.xl.error.{XLError, XLResult}

import scala.xml.*
import java.io.{ByteArrayInputStream, FileInputStream, InputStream}
import java.util.zip.ZipInputStream
import java.nio.file.{Path, Paths}
import scala.collection.mutable
import com.tjclp.xl.style.StyleRegistry
import com.tjclp.xl.style.units.StyleId

/**
 * Reader for XLSX files (ZIP parsing)
 *
 * Parses XLSX ZIP structure and converts to domain Workbook.
 *
 * **WARNING**: This is the in-memory reader that loads the entire file into memory. For large files
 * (>10k rows), use `ExcelIO.readStream()` instead for constant-memory streaming.
 *
 * Use this reader for:
 *   - Small files (<10k rows)
 *   - When you need random cell access
 *   - When you need styling information
 *
 * Use `ExcelIO.readStream()` for:
 *   - Large files (100k+ rows)
 *   - Sequential row processing
 *   - Constant-memory requirements (O(1) memory regardless of file size)
 */
object XlsxReader:

  /**
   * Read workbook from XLSX file (in-memory).
   *
   * Loads entire file into memory. For large files, use `ExcelIO.readStream()` instead.
   */
  def read(inputPath: Path): XLResult[Workbook] =
    try
      val is = new FileInputStream(inputPath.toFile)
      try
        readFromStream(is)
      finally
        is.close()
    catch case e: Exception => Left(XLError.IOError(s"Failed to read XLSX: ${e.getMessage}"))

  /** Read workbook from byte array (for testing) */
  def readFromBytes(bytes: Array[Byte]): XLResult[Workbook] =
    try readFromStream(new ByteArrayInputStream(bytes))
    catch case e: Exception => Left(XLError.IOError(s"Failed to read bytes: ${e.getMessage}"))

  /** Read workbook from input stream */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
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

    finally zip.close()

  /** Parse workbook from collected parts */
  private def parseWorkbook(parts: Map[String, String]): XLResult[Workbook] =
    for
      // Parse workbook.xml
      workbookXml <- parts
        .get("xl/workbook.xml")
        .toRight(XLError.ParseError("xl/workbook.xml", "Missing workbook.xml"))
      workbookElem <- parseXml(workbookXml, "xl/workbook.xml")
      ooxmlWb <- OoxmlWorkbook
        .fromXml(workbookElem)
        .left
        .map(err => XLError.ParseError("xl/workbook.xml", err): XLError)

      // Parse optional shared strings
      sst <- parseOptionalSST(parts)

      // Parse styles (if present)
      styles <- parseStyles(parts)

      // Parse workbook relationships
      workbookRels <- parseWorkbookRelationships(parts)

      // Parse sheets
      sheets <- parseSheets(parts, ooxmlWb.sheets, sst, styles, workbookRels)

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
          sst <- SharedStrings
            .fromXml(elem)
            .left
            .map(err => XLError.ParseError("xl/sharedStrings.xml", err): XLError)
        yield Some(sst)

  /** Parse styles table (falls back to default styles if missing) */
  private def parseStyles(parts: Map[String, String]): XLResult[WorkbookStyles] =
    parts.get("xl/styles.xml") match
      case None => Right(WorkbookStyles.default)
      case Some(xml) =>
        for
          elem <- parseXml(xml, "xl/styles.xml")
          styles <- WorkbookStyles
            .fromXml(elem)
            .left
            .map(err => XLError.ParseError("xl/styles.xml", err): XLError)
        yield styles

  /** Parse workbook relationships (used to resolve worksheet locations) */
  private def parseWorkbookRelationships(parts: Map[String, String]): XLResult[Relationships] =
    parts.get("xl/_rels/workbook.xml.rels") match
      case None => Right(Relationships.empty)
      case Some(xml) =>
        for
          elem <- parseXml(xml, "xl/_rels/workbook.xml.rels")
          rels <- Relationships
            .fromXml(elem)
            .left
            .map(err => XLError.ParseError("xl/_rels/workbook.xml.rels", err): XLError)
        yield rels

  /** Parse all worksheets */
  private def parseSheets(
    parts: Map[String, String],
    sheetRefs: Seq[SheetRef],
    sst: Option[SharedStrings],
    styles: WorkbookStyles,
    relationships: Relationships
  ): XLResult[Vector[Sheet]] =
    val relMap = relationships.relationships.map(rel => rel.id -> rel).toMap
    sheetRefs.toVector.traverse { ref =>
      val sheetPath = relMap
        .get(ref.relationshipId)
        .map(rel => resolveSheetPath(rel.target))
        .getOrElse(defaultSheetPath(ref.sheetId))

      for
        xml <- parts
          .get(sheetPath)
          .toRight(XLError.ParseError(sheetPath, s"Missing worksheet: $sheetPath"))
        elem <- parseXml(xml, sheetPath)
        ooxmlSheet <- OoxmlWorksheet
          .fromXml(elem)
          .left
          .map(err => XLError.ParseError(sheetPath, err): XLError)
        domainSheet <- convertToDomainSheet(ref.name, ooxmlSheet, sst, styles)
      yield domainSheet
    }

  private def defaultSheetPath(sheetId: Int): String = s"xl/worksheets/sheet$sheetId.xml"

  private def resolveSheetPath(target: String): String =
    val cleaned = if target.startsWith("/") then target.drop(1) else target
    val resolvedPath =
      if cleaned.startsWith("xl/") || cleaned.startsWith("xl\\") then Paths.get(cleaned)
      else Paths.get("xl").resolve(cleaned)
    resolvedPath.normalize().toString.replace('\\', '/')

  /** Convert OoxmlWorksheet to domain Sheet */
  private def convertToDomainSheet(
    name: SheetName,
    ooxmlSheet: OoxmlWorksheet,
    sst: Option[SharedStrings],
    styles: WorkbookStyles
  ): XLResult[Sheet] =
    val (cellsMap, finalRegistry) =
      ooxmlSheet.rows.foldLeft((Map.empty[ARef, Cell], StyleRegistry.default)) {
        case ((cellsAcc, registry), row) =>
          row.cells.foldLeft((cellsAcc, registry)) { case ((cellMapAcc, registryAcc), ooxmlCell) =>
            val value = (ooxmlCell.cellType, ooxmlCell.value, sst) match
              case ("s", CellValue.Text(idxStr), Some(sharedStrings)) =>
                idxStr.toIntOption match
                  case Some(idx) =>
                    sharedStrings(idx) match
                      case Some(text) => CellValue.Text(text)
                      case None => CellValue.Error(CellError.Ref)
                  case None => CellValue.Error(com.tjclp.xl.cell.CellError.Value)
              case _ => ooxmlCell.value

            val (nextRegistry, styleIdOpt) = ooxmlCell.styleIndex
              .flatMap(styles.styleAt)
              .map { style =>
                val (updatedRegistry, styleId) = registryAcc.register(style)
                (updatedRegistry, Some(styleId))
              }
              .getOrElse((registryAcc, None))

            val cell = Cell(ooxmlCell.ref, value, styleIdOpt)
            (cellMapAcc.updated(cell.ref, cell), nextRegistry)
          }
      }

    Right(
      Sheet(
        name = name,
        cells = cellsMap,
        mergedRanges = ooxmlSheet.mergedRanges,
        styleRegistry = finalRegistry
      )
    )

  /** Assemble final workbook */
  private def assembleWorkbook(sheets: Vector[Sheet]): XLResult[Workbook] =
    if sheets.isEmpty then Left(XLError.InvalidWorkbook("Workbook must have at least one sheet"))
    else Right(Workbook(sheets = sheets))

  /**
   * Parse XML string to Elem with XXE protection
   *
   * REQUIRES: xmlString is a valid XML string ENSURES:
   *   - Returns Right(elem) if parsing succeeds
   *   - Returns Left(XLError) if parsing fails
   *   - Rejects XML with DOCTYPE declarations (XXE attack prevention)
   *   - Rejects external entity references
   * DETERMINISTIC: Yes (for same input) ERROR CASES: Malformed XML, XXE attempts, parser
   * configuration failures
   */
  private def parseXml(xmlString: String, location: String): XLResult[Elem] =
    try
      // Configure SAX parser to prevent XXE (XML External Entity) attacks
      val factory = javax.xml.parsers.SAXParserFactory.newInstance()
      factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true)
      factory.setFeature("http://xml.org/sax/features/external-general-entities", false)
      factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false)
      factory.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false)
      factory.setXIncludeAware(false)
      factory.setNamespaceAware(true)

      val loader = XML.withSAXParser(factory.newSAXParser())
      Right(loader.loadString(xmlString))
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
