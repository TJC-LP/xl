package com.tjclp.xl.ooxml

import scala.xml.*
import java.io.{File, FileOutputStream, ByteArrayOutputStream}
import java.util.zip.{ZipOutputStream, ZipEntry}
import java.nio.file.{Path, Files}
import java.nio.charset.StandardCharsets
import com.tjclp.xl.{Workbook, XLError, XLResult}

/** Shared Strings Table usage policy */
enum SstPolicy derives CanEqual:
  /** Auto-detect based on heuristics (default) */
  case Auto

  /** Always use SST regardless of content */
  case Always

  /** Never use SST (inline strings only) */
  case Never

/**
 * Compression method for ZIP entries.
 *
 * DEFLATED (default) produces 5-10x smaller files with minimal CPU overhead. STORED is useful for
 * debugging (human-readable ZIP contents).
 */
enum Compression derives CanEqual:
  /** No compression (STORED) - faster writes, larger files, requires CRC32 precomputation */
  case Stored

  /** DEFLATE compression (DEFLATED) - smaller files, standard production use */
  case Deflated

  /** ZIP constant for this compression method */
  def zipMethod: Int = this match
    case Stored => ZipEntry.STORED
    case Deflated => ZipEntry.DEFLATED

/** Writer configuration options */
case class WriterConfig(
  sstPolicy: SstPolicy = SstPolicy.Auto,
  compression: Compression = Compression.Deflated,
  prettyPrint: Boolean = false // Compact XML for production
)

object WriterConfig:
  /** Default production configuration: DEFLATED compression + compact XML */
  val default: WriterConfig = WriterConfig()

  /** Debug configuration: STORED compression + pretty XML for manual inspection */
  val debug: WriterConfig = WriterConfig(
    compression = Compression.Stored,
    prettyPrint = true
  )

/**
 * Writer for XLSX files (ZIP assembly)
 *
 * Takes a domain Workbook and produces a valid XLSX file with all required parts.
 */
object XlsxWriter:

  /** Write workbook to XLSX file with default configuration */
  def write(workbook: Workbook, outputPath: Path): XLResult[Unit] =
    writeWith(workbook, outputPath, WriterConfig())

  /** Write workbook to XLSX file with custom configuration */
  def writeWith(
    workbook: Workbook,
    outputPath: Path,
    config: WriterConfig = WriterConfig()
  ): XLResult[Unit] =
    try
      // Build shared strings table based on policy
      val sst = config.sstPolicy match
        case SstPolicy.Always => Some(SharedStrings.fromWorkbook(workbook))
        case SstPolicy.Never => None
        case SstPolicy.Auto =>
          if SharedStrings.shouldUseSST(workbook) then Some(SharedStrings.fromWorkbook(workbook))
          else None

      // Build unified style index with per-sheet remappings
      val (styleIndex, sheetRemappings) = StyleIndex.fromWorkbook(workbook)
      val styles = OoxmlStyles(styleIndex)

      // Convert domain workbook to OOXML
      val ooxmlWb = OoxmlWorkbook.fromDomain(workbook)

      // Convert sheets to OOXML worksheets with style remapping
      val ooxmlSheets = workbook.sheets.zipWithIndex.map { case (sheet, sheetIdx) =>
        val remapping = sheetRemappings.getOrElse(sheetIdx, Map.empty)
        OoxmlWorksheet.fromDomainWithSST(sheet, sst, remapping)
      }

      // Create content types
      val contentTypes = ContentTypes.minimal(
        hasStyles = true, // Always include styles
        hasSharedStrings = sst.isDefined,
        sheetCount = workbook.sheets.size
      )

      // Create relationships
      val rootRels = Relationships.root()
      val workbookRels = Relationships.workbook(
        sheetCount = workbook.sheets.size,
        hasStyles = true,
        hasSharedStrings = sst.isDefined
      )

      // Assemble ZIP
      writeZip(
        outputPath,
        contentTypes,
        rootRels,
        workbookRels,
        ooxmlWb,
        ooxmlSheets,
        styles,
        sst,
        config
      )

      Right(())

    catch case e: Exception => Left(XLError.IOError(s"Failed to write XLSX: ${e.getMessage}"))

  /** Write all parts to ZIP file */
  private def writeZip(
    path: Path,
    contentTypes: ContentTypes,
    rootRels: Relationships,
    workbookRels: Relationships,
    workbook: OoxmlWorkbook,
    sheets: Vector[OoxmlWorksheet],
    styles: OoxmlStyles,
    sst: Option[SharedStrings],
    config: WriterConfig
  ): Unit =
    val zip = new ZipOutputStream(new FileOutputStream(path.toFile))
    try
      // Write parts in canonical order
      writePart(zip, "[Content_Types].xml", contentTypes.toXml, config)
      writePart(zip, "_rels/.rels", rootRels.toXml, config)
      writePart(zip, "xl/workbook.xml", workbook.toXml, config)
      writePart(zip, "xl/_rels/workbook.xml.rels", workbookRels.toXml, config)

      // Write styles
      writePart(zip, "xl/styles.xml", styles.toXml, config)

      // Write shared strings if present
      sst.foreach { sharedStrings =>
        writePart(zip, "xl/sharedStrings.xml", sharedStrings.toXml, config)
      }

      // Write worksheets
      sheets.zipWithIndex.foreach { case (sheet, idx) =>
        writePart(zip, s"xl/worksheets/sheet${idx + 1}.xml", sheet.toXml, config)
      }

    finally zip.close()

  /** Write a single XML part to ZIP */
  private def writePart(
    zip: ZipOutputStream,
    entryName: String,
    xml: Elem,
    config: WriterConfig
  ): Unit =
    val entry = new ZipEntry(entryName)
    entry.setMethod(config.compression.zipMethod)

    // Convert XML to bytes with conditional formatting
    val xmlString = if config.prettyPrint then XmlUtil.prettyPrint(xml) else XmlUtil.compact(xml)
    val bytes = xmlString.getBytes(StandardCharsets.UTF_8)

    // For STORED method, must set size and CRC before writing
    // For DEFLATED, these are computed automatically by ZipOutputStream
    config.compression match
      case Compression.Stored =>
        entry.setSize(bytes.length)
        entry.setCompressedSize(bytes.length)
        entry.setCrc(calculateCrc(bytes))
      case Compression.Deflated =>
        // ZipOutputStream computes these automatically for DEFLATED
        ()

    zip.putNextEntry(entry)
    zip.write(bytes)
    zip.closeEntry()

  /** Calculate CRC32 checksum for ZIP entry */
  private def calculateCrc(bytes: Array[Byte]): Long =
    val crc = new java.util.zip.CRC32()
    crc.update(bytes)
    crc.getValue

  /** Write workbook to bytes (for testing) */
  def writeToBytes(workbook: Workbook): XLResult[Array[Byte]] =
    try
      val baos = new ByteArrayOutputStream()
      val tempPath = Files.createTempFile("xl-", ".xlsx")
      try
        write(workbook, tempPath).map { _ =>
          Files.readAllBytes(tempPath)
        }
      finally
        Files.deleteIfExists(tempPath)
    catch case e: Exception => Left(XLError.IOError(s"Failed to write bytes: ${e.getMessage}"))
