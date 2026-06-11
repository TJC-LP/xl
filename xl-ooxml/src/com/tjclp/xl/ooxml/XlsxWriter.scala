package com.tjclp.xl.ooxml

import scala.xml.*
import java.io.{ByteArrayOutputStream, FileOutputStream, OutputStream}
import java.security.MessageDigest
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}
import java.nio.file.{Files, Path, StandardCopyOption}
import java.nio.charset.StandardCharsets
import scala.collection.immutable.ArraySeq
import scala.collection.mutable
import scala.util.{Failure, Success, Try, Using}
import com.tjclp.xl.api.{Sheet, Workbook, CellValue}
import com.tjclp.xl.drawings.ImageFormat
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.context.{ModificationTracker, SourceContext}
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.tables.TableSpec
import com.tjclp.xl.ooxml.chart.{ChartCaches, OoxmlChart}
import com.tjclp.xl.ooxml.drawing.{DrawingReader, OoxmlDrawing}
import com.tjclp.xl.ooxml.worksheet.collectHyperlinks

// Re-export writer config types for backward compatibility
export writer.{
  SstPolicy,
  Compression,
  FormulaInjectionPolicy,
  XmlBackend,
  WriterConfig,
  OutputTarget,
  OutputPath,
  OutputStreamTarget
}

// Import for internal use
import writer.*

/**
 * Unified XLSX writer with intelligent surgical modification.
 *
 * Single code path that automatically optimizes based on context:
 *   - **With source** (read → modify → write): Surgical mode
 *     - Clean → verbatim copy (11x faster)
 *     - Partially modified → regenerate changed, copy unchanged (2-5x faster)
 *     - Fully modified → regenerate all with style preservation
 *   - **Without source** (create → write): Full regeneration (standard path)
 *
 * The writer transparently chooses the optimal strategy - users don't need to decide.
 *
 * **Key Features**:
 *   - Preserves unknown parts (charts, comments, drawings)
 *   - Preserves differential formats (dxfs) for conditional formatting
 *   - Byte-identical copies for unmodified sheets
 *   - RichText in SharedStrings (not inlined)
 *   - Row-level style preservation
 *   - Excel compression level matching (defS = level 1)
 */
object XlsxWriter:

  /** Write workbook to XLSX file with default configuration */
  def write(workbook: Workbook, outputPath: Path): XLResult[Unit] =
    writeWith(workbook, outputPath, WriterConfig())

  /**
   * Write workbook to XLSX file with custom configuration.
   *
   * **Surgical Modification**: Automatically uses hybrid strategy when workbook has SourceContext:
   *   - Clean workbook → verbatim copy (11x faster)
   *   - Partially modified → regenerate changed parts, preserve rest (2-5x faster)
   *   - No SourceContext → full regeneration (current behavior)
   */
  def writeWith(
    workbook: Workbook,
    outputPath: Path,
    config: WriterConfig = WriterConfig()
  ): XLResult[Unit] =
    writeToTarget(workbook, OutputPath(outputPath), config)

  /**
   * Internal dispatch: Choose write strategy based on SourceContext.
   *
   * Strategy selection:
   *   1. SourceContext + clean + file target → verbatim copy (fastest)
   *   2. File target with source → atomic temp file + rename (prevents corruption)
   *   3. All other cases → unified write (surgical if source available, else full regeneration)
   *
   * Unified write automatically optimizes:
   *   - With source: regenerate modified, copy unmodified (surgical)
   *   - Without source: regenerate all (graceful degradation)
   *
   * **Atomic Write Safety**: When writing to a file with SourceContext (surgical mode), we always
   * write to a temp file first, then atomically rename. This prevents:
   *   - Corruption when source == destination (chained writes)
   *   - Partial writes on failure leaving broken files
   */
  private def writeToTarget(
    workbook: Workbook,
    target: OutputTarget,
    config: WriterConfig
  ): XLResult[Unit] =
    try
      val escapeFormulas = formulaEscapingRequested(config)
      // GH-243: serialize DateTime cells with the workbook's date system before any write path.
      val normalized = withDates1904Normalized(workbook)

      normalized.sourceContext match
        case Some(ctx) if ctx.isClean && !escapeFormulas =>
          // Clean workbook + file target → verbatim copy (ultra-fast)
          target match
            case OutputPath(path) =>
              copyVerbatim(ctx, path)
            case OutputStreamTarget(_) =>
              // Can't copy to stream, use unified write (will copy all parts)
              unifiedWrite(normalized, normalized.sourceContext, target, config)

        case Some(ctx) =>
          // Surgical mode with source: use atomic temp file to prevent corruption
          target match
            case OutputPath(destPath) =>
              writeAtomically(normalized, Some(ctx), destPath, config)
            case _ =>
              unifiedWrite(normalized, normalized.sourceContext, target, config)

        case None =>
          // No source context: safe to write directly (no surgical preservation)
          unifiedWrite(normalized, None, target, config)

      Right(())

    catch case e: Exception => Left(XLError.IOError(s"Failed to write XLSX: ${e.getMessage}"))

  /**
   * Re-express DateTime cells as raw serial Numbers using the workbook's date system (GH-243).
   *
   * The cell emitters (OoxmlCell/OoxmlWorksheet/DirectSaxEmitter) convert any remaining DateTime
   * values with the default 1900 epoch, so for 1904-system workbooks the conversion must happen
   * here, once, before dispatching to any write path. No-op for 1900-system workbooks (the
   * emitters' inline conversion is identical) and for workbooks without DateTime cells — in
   * particular, freshly read workbooks store date serials as Numbers, so surgical-write
   * modification tracking is unaffected.
   */
  private def withDates1904Normalized(workbook: Workbook): Workbook =
    if !workbook.metadata.date1904 then workbook
    else
      val normalizedSheets = workbook.sheets.map { sheet =>
        val hasDateTime = sheet.cells.values.exists(_.value match
          case CellValue.DateTime(_) => true
          case _ => false)
        if !hasDateTime then sheet
        else
          val cells = sheet.cells.map { case (ref, cell) =>
            cell.value match
              case CellValue.DateTime(dt) =>
                val serial = CellValue.dateTimeToExcelSerial(dt, date1904 = true)
                ref -> cell.copy(value = CellValue.Number(BigDecimal(serial)))
              case _ => ref -> cell
          }
          sheet.copy(cells = cells)
      }
      workbook.copy(sheets = normalizedSheets)

  /**
   * Write to file atomically via temp file + rename.
   *
   * This prevents corruption when:
   *   - Source and destination are the same file (chained writes)
   *   - Write fails partway through (partial file corruption)
   *
   * Process:
   *   1. Create temp file in same directory as destination
   *   2. Write complete XLSX to temp file
   *   3. Atomically rename temp → destination
   *   4. Clean up temp file on failure
   */
  private def writeAtomically(
    workbook: Workbook,
    sourceContext: Option[SourceContext],
    destPath: Path,
    config: WriterConfig
  ): Unit =
    val parent = Option(destPath.getParent).getOrElse(Path.of("."))
    val tempPath = Files.createTempFile(parent, ".xl-", ".tmp")

    try
      unifiedWrite(workbook, sourceContext, OutputPath(tempPath), config)
      Files.move(
        tempPath,
        destPath,
        StandardCopyOption.REPLACE_EXISTING,
        StandardCopyOption.ATOMIC_MOVE
      )
    catch
      case e: java.nio.file.AtomicMoveNotSupportedException =>
        // Fallback for filesystems that don't support atomic move
        Files.move(tempPath, destPath, StandardCopyOption.REPLACE_EXISTING)
      case e: Exception =>
        Files.deleteIfExists(tempPath)
        throw e

  private def formulaEscapingRequested(config: WriterConfig): Boolean =
    config.formulaInjectionPolicy == FormulaInjectionPolicy.Escape

  /**
   * Copy source file verbatim to destination (for clean workbooks).
   *
   * Fast path optimization: When a workbook has no modifications, just copy the source file
   * byte-for-byte instead of regenerating all XML. This is 10-11x faster than full regeneration.
   *
   * Handles edge case where source == dest (no-op).
   */
  private def copyVerbatim(ctx: SourceContext, dest: Path): Unit =
    val source = ctx.sourcePath
    if source != dest then
      val fingerprint = ctx.fingerprint
      val currentSize = Files.size(source)
      if currentSize != fingerprint.size then
        throw new IllegalStateException(
          s"Source file changed size since read (expected ${fingerprint.size} bytes, found $currentSize)"
        )

      val digest = MessageDigest.getInstance("SHA-256")

      val bytesCopied = usingOrThrow(Using.Manager { use =>
        val in = use(Files.newInputStream(source))
        val out = use(Files.newOutputStream(dest))
        val buffer = new Array[Byte](8192)

        def loop(total: Long): Long =
          val read = in.read(buffer)
          if read == -1 then total
          else
            digest.update(buffer, 0, read)
            out.write(buffer, 0, read)
            loop(total + read)

        loop(0L)
      })

      val computedDigest = digest.digest()
      if !fingerprint.matches(bytesCopied, computedDigest) then
        Files.deleteIfExists(dest)
        throw new IllegalStateException("Source file changed since read; refusing to copy verbatim")

  /**
   * Canonical comment author: trimmed, whitespace-only → unauthored (GH-290).
   *
   * The trim happens ONCE at write time; the `<authors>` entry and the bold author-prefix run are
   * both built from this value, so `XlsxReader.stripAuthorPrefix` (which matches the stored author
   * against the first run) always strips the prefix and it never leaks into the comment text.
   */
  private def canonicalCommentAuthor(author: Option[String]): Option[String] =
    author.map(_.trim).filter(_.nonEmpty)

  /**
   * Build per-sheet comment data for serialization.
   *
   * Returns:
   *   - Map[Int, OoxmlComments]: sheet index (0-based) -> comments to write
   *   - Set[Int]: indices (1-based) of sheets with comments for content types
   */
  private def buildCommentsData(workbook: Workbook): (Map[Int, OoxmlComments], Set[Int]) =
    val commentsBySheet = workbook.sheets.zipWithIndex.flatMap { case (sheet, idx) =>
      if sheet.comments.isEmpty then None
      else
        // Build author list (canonicalized, deduplicated and sorted for deterministic output)
        // Reserve index 0 for unauthored comments (empty string)
        val (authorSet, hasUnauthored) =
          sheet.comments.values.foldLeft((Set.empty[String], false)) {
            case ((existing, hasNone), comment) =>
              canonicalCommentAuthor(comment.author) match
                case Some(author) => (existing + author, hasNone)
                case None => (existing, true)
          }
        // Sort for deterministic output (Excel preserves insertion order; we choose determinism here)
        // to keep serialized files stable across runs.
        val realAuthors = authorSet.toVector.sorted
        val authors = if hasUnauthored then Vector("") ++ realAuthors else realAuthors

        val authorMap = authors.zipWithIndex.map { case (author, i) => author -> i }.toMap

        // Convert domain Comments to OOXML (sorted by ref for deterministic output)
        val ooxmlComments = sheet.comments.toVector.sortBy(_._1.toA1).map { case (ref, comment) =>
          val canonicalAuthor = canonicalCommentAuthor(comment.author)
          val authorId =
            canonicalAuthor.flatMap(authorMap.get).getOrElse(0) // Index 0 = "" for unauthored
          // Note: xr:uid GUIDs omitted for new comments (deterministic output)
          // GUIDs are optional per OOXML spec and only needed for revision tracking

          // Excel displays author as part of comment text (bold first run)
          val textWithAuthor = canonicalAuthor match
            case Some(author) =>
              // Prepend author name as bold run
              val authorRun = com.tjclp.xl.richtext.TextRun(
                s"$author:",
                Some(com.tjclp.xl.styles.font.Font.default.copy(bold = true))
              )
              // Prepend newline to the first run of comment text
              val textWithNewline = comment.text.runs.headOption match
                case Some(head) =>
                  // Create new TextRun with newline prepended
                  val tail = comment.text.runs.drop(1)
                  val textWithNewline = "\n" + head.text
                  val modifiedFirstRun = com.tjclp.xl.richtext.TextRun(
                    textWithNewline,
                    head.font,
                    head.rawRPrXml
                  )
                  com.tjclp.xl.richtext.RichText(Vector(authorRun, modifiedFirstRun) ++ tail)
                case None =>
                  // Empty comment text - just add author with newline
                  com.tjclp.xl.richtext.RichText(
                    Vector(
                      authorRun,
                      com.tjclp.xl.richtext.TextRun("\n")
                    )
                  )
              textWithNewline
            case None => comment.text

          OoxmlComment(
            ref = ref,
            authorId = authorId,
            text = textWithAuthor,
            guid = None // Omit for deterministic output (GUIDs optional per spec)
          )
        }

        Some(idx -> OoxmlComments(authors, ooxmlComments))
    }.toMap

    // Convert to 1-based indices for file naming and content types
    val sheetsWithComments = commentsBySheet.keySet.map(_ + 1)

    (commentsBySheet, sheetsWithComments)

  /**
   * Build per-sheet table data for serialization.
   *
   * Assigns global table IDs sequentially across all sheets (1-indexed). Tables are sorted by name
   * within each sheet for deterministic output.
   *
   * Returns:
   *   - Map[Int, Seq[(TableSpec, Long)]]: sheet index (0-based) → tables with global IDs
   *   - Int: total table count (for content types registration)
   *   - Map[String, Long]: table name → global table ID (for relationship targeting)
   */
  private def buildTablesData(
    workbook: Workbook
  ): (Map[Int, Seq[(TableSpec, Long)]], Int, Map[String, Long]) =
    // Flatten all tables from all sheets with their sheet indices
    // Sort by name within each sheet for deterministic ordering
    val allTablesWithIndices: Seq[(TableSpec, Int)] = workbook.sheets.zipWithIndex.flatMap {
      case (sheet, sheetIdx) =>
        sheet.tables.values.toSeq.sortBy(_.name).map(table => (table, sheetIdx))
    }

    // Assign sequential global IDs (1-indexed: table1.xml, table2.xml, etc.)
    val tablesWithIds: Seq[(TableSpec, Int, Long)] = allTablesWithIndices.zipWithIndex.map {
      case ((table, sheetIdx), globalIdx) => (table, sheetIdx, (globalIdx + 1).toLong)
    }

    // Group by sheet index for per-sheet processing
    val tablesBySheet: Map[Int, Seq[(TableSpec, Long)]] = tablesWithIds
      .groupMap { case (_, sheetIdx, _) => sheetIdx } { case (table, _, tableId) =>
        (table, tableId)
      }

    // Total table count for content types
    val totalTableCount = tablesWithIds.size

    // Table name → global ID mapping for lookups
    val tableIdMap: Map[String, Long] = tablesWithIds.map { case (table, _, tableId) =>
      table.name -> tableId
    }.toMap

    (tablesBySheet, totalTableCount, tableIdMap)

  /**
   * Build worksheet relationships using the actual comment file path.
   *
   * Excel numbers comment files sequentially (comments1.xml, comments2.xml...) across only sheets
   * that have comments, NOT by sheet index. This method accepts the correct path to ensure surgical
   * writes preserve the source file's comment numbering.
   */
  private def buildWorksheetRelationshipsWithCommentsPath(
    commentPath: String,
    hasComments: Boolean,
    tableIds: Seq[Long]
  ): Relationships =
    // Extract file number from comment path (e.g., "xl/comments3.xml" → "3")
    val commentFileNum = commentPath.stripPrefix("xl/comments").stripSuffix(".xml")

    val commentRels =
      if hasComments then
        Seq(
          Relationship(
            id = "rId1",
            `type` = XmlUtil.relTypeComments,
            target = s"../comments$commentFileNum.xml"
          ),
          Relationship(
            id = "rId2",
            `type` = XmlUtil.relTypeVmlDrawing,
            target = s"../drawings/vmlDrawing$commentFileNum.vml"
          )
        )
      else Seq.empty

    val rIdOffset = if hasComments then 3 else 1

    // Table relationships (sorted for deterministic output)
    val tableRels = tableIds.sorted.zipWithIndex.map { case (tableId, idx) =>
      Relationship(
        id = s"rId${rIdOffset + idx}",
        `type` = XmlUtil.relTypeTable,
        target = s"../tables/table$tableId.xml"
      )
    }

    Relationships(commentRels ++ tableRels)

  /** Write worksheet directly from domain Sheet using DirectSaxEmitter */
  private def writeWorksheetDirect(
    zip: ZipOutputStream,
    entryName: String,
    sheet: Sheet,
    sst: Option[SharedStrings],
    styleRemapping: Map[Int, Int],
    tablePartsXml: Option[scala.xml.Elem],
    escapeFormulas: Boolean,
    config: WriterConfig
  ): Unit =
    val entry = new ZipEntry(entryName)
    entry.setTime(0L)
    entry.setMethod(config.compression.zipMethod)

    config.compression match
      case Compression.Stored =>
        val baos = new ByteArrayOutputStream()
        val saxWriter = StaxSaxWriter.create(baos)
        DirectSaxEmitter.emitWorksheet(
          saxWriter,
          sheet,
          sst,
          styleRemapping,
          tablePartsXml,
          escapeFormulas
        )
        val bytes = baos.toByteArray

        entry.setSize(bytes.length)
        entry.setCompressedSize(bytes.length)
        entry.setCrc(calculateCrc(bytes))

        zip.putNextEntry(entry)
        zip.write(bytes)
        zip.closeEntry()

      case Compression.Deflated =>
        // Buffer first to avoid many small writes to compressed stream
        // (single bulk write is much more efficient for DEFLATE algorithm)
        val baos = new ByteArrayOutputStream()
        val saxWriter = StaxSaxWriter.create(baos)
        DirectSaxEmitter.emitWorksheet(
          saxWriter,
          sheet,
          sst,
          styleRemapping,
          tablePartsXml,
          escapeFormulas
        )
        saxWriter.flush()
        val bytes = baos.toByteArray

        zip.putNextEntry(entry)
        zip.write(bytes)
        zip.closeEntry()

  /** Write VML part to ZIP (VML is plain text, not standard XML) */
  private def writeVmlPart(
    zip: ZipOutputStream,
    entryName: String,
    vmlXml: String,
    config: WriterConfig
  ): Unit =
    val entry = new ZipEntry(entryName)
    entry.setTime(0L) // Deterministic timestamps
    entry.setMethod(config.compression.zipMethod)

    val bytes = vmlXml.getBytes(StandardCharsets.UTF_8)

    config.compression match
      case Compression.Stored =>
        entry.setSize(bytes.length)
        entry.setCompressedSize(bytes.length)
        val crc = new java.util.zip.CRC32()
        crc.update(bytes)
        entry.setCrc(crc.getValue)
      case Compression.Deflated => ()

    zip.putNextEntry(entry)
    zip.write(bytes)
    zip.closeEntry()

  /** Write a single XML part to ZIP */
  private def writePart(
    zip: ZipOutputStream,
    entryName: String,
    xml: Elem,
    config: WriterConfig
  ): Unit =
    val entry = new ZipEntry(entryName)
    entry.setTime(0L) // deterministic ZIP metadata for reproducible output
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

  /** Write an XML part using configured backend */
  private def writePart(
    zip: ZipOutputStream,
    entryName: String,
    part: SaxSerializable & XmlWritable,
    config: WriterConfig
  ): Unit =
    config.backend match
      case XmlBackend.ScalaXml => writePart(zip, entryName, part.toXml, config)
      case XmlBackend.SaxStax => writeGenericSax(zip, entryName, config)(part.writeSax)

  /** Write worksheet using selected backend */
  private def writeWorksheet(
    zip: ZipOutputStream,
    entryName: String,
    sheet: OoxmlWorksheet,
    config: WriterConfig
  ): Unit =
    config.backend match
      case XmlBackend.ScalaXml =>
        writePart(zip, entryName, sheet.toXml, config)
      case XmlBackend.SaxStax =>
        writeWorksheetSax(zip, entryName, sheet, config)

  private def writeWorksheetSax(
    zip: ZipOutputStream,
    entryName: String,
    sheet: OoxmlWorksheet,
    config: WriterConfig
  ): Unit =
    val entry = new ZipEntry(entryName)
    entry.setTime(0L)
    entry.setMethod(config.compression.zipMethod)

    config.compression match
      case Compression.Stored =>
        val baos = new ByteArrayOutputStream()
        val saxWriter = StaxSaxWriter.create(baos)
        sheet.writeSax(saxWriter)
        saxWriter.flush()
        val bytes = baos.toByteArray

        entry.setSize(bytes.length)
        entry.setCompressedSize(bytes.length)
        entry.setCrc(calculateCrc(bytes))

        zip.putNextEntry(entry)
        zip.write(bytes)
        zip.closeEntry()

      case Compression.Deflated =>
        // Buffer first to avoid many small writes to compressed stream
        val baos = new ByteArrayOutputStream()
        val saxWriter = StaxSaxWriter.create(baos)
        sheet.writeSax(saxWriter)
        saxWriter.flush()
        val bytes = baos.toByteArray

        zip.putNextEntry(entry)
        zip.write(bytes)
        zip.closeEntry()

  private def writeSharedStrings(
    zip: ZipOutputStream,
    entryName: String,
    sst: SharedStrings,
    config: WriterConfig
  ): Unit =
    writePart(zip, entryName, sst, config)

  private def writeStyles(
    zip: ZipOutputStream,
    entryName: String,
    styles: OoxmlStyles,
    config: WriterConfig
  ): Unit =
    writePart(zip, entryName, styles, config)

  private def writeGenericSax(
    zip: ZipOutputStream,
    entryName: String,
    config: WriterConfig
  )(body: SaxWriter => Unit): Unit =
    val entry = new ZipEntry(entryName)
    entry.setTime(0L)
    entry.setMethod(config.compression.zipMethod)

    config.compression match
      case Compression.Stored =>
        val baos = new ByteArrayOutputStream()
        val saxWriter = StaxSaxWriter.create(baos)
        body(saxWriter)
        saxWriter.flush()
        val bytes = baos.toByteArray

        entry.setSize(bytes.length)
        entry.setCompressedSize(bytes.length)
        entry.setCrc(calculateCrc(bytes))

        zip.putNextEntry(entry)
        zip.write(bytes)
        zip.closeEntry()

      case Compression.Deflated =>
        // Buffer first to avoid many small writes to compressed stream
        val baos = new ByteArrayOutputStream()
        val saxWriter = StaxSaxWriter.create(baos)
        body(saxWriter)
        saxWriter.flush()
        val bytes = baos.toByteArray

        zip.putNextEntry(entry)
        zip.write(bytes)
        zip.closeEntry()

  /** Calculate CRC32 checksum for ZIP entry */
  private def calculateCrc(bytes: Array[Byte]): Long =
    val crc = new java.util.zip.CRC32()
    crc.update(bytes)
    crc.getValue

  // ========== Surgical Modification Methods ==========

  /**
   * Determine which parts can be preserved byte-for-byte during surgical write.
   *
   * A part can be preserved if:
   *   1. It's in the unparsed set (XL didn't parse it)
   *   1. It doesn't reference any modified or deleted sheets
   *
   * @return
   *   Set of ZIP entry paths that can be copied verbatim from source
   */
  private def determinePreservableParts(
    wb: Workbook,
    ctx: SourceContext,
    graph: RelationshipGraph
  ): Set[String] =
    val tracker = ctx.modificationTracker
    val unparsed = ctx.partManifest.unparsedParts

    unparsed.filter { path =>
      val dependentSheets = graph.dependenciesFor(path)
      val touchesDeleted = dependentSheets.exists(tracker.deletedSheets.contains)
      val touchesModified = dependentSheets.exists(tracker.modifiedSheets.contains)

      // Preserve only if doesn't touch any modified/deleted sheets
      !touchesDeleted && !touchesModified
    }

  /**
   * Determine which parts must be regenerated during surgical write.
   *
   * Parts that must be regenerated:
   *   - Structural parts (always): workbook.xml, relationships, content types
   *   - Styles/SST if any sheet modified (indices may change)
   *   - Modified sheets
   *   - Relationships for modified/deleted sheets
   *
   * @return
   *   Set of ZIP entry paths that must be regenerated from domain model
   */
  private def determineRegenerateParts(
    wb: Workbook,
    ctx: SourceContext
  ): Set[String] =
    val tracker = ctx.modificationTracker
    val regenerate = mutable.Set[String]()

    // Always regenerate structural parts
    regenerate ++= Set(
      "[Content_Types].xml",
      "_rels/.rels",
      "xl/workbook.xml",
      "xl/_rels/workbook.xml.rels"
    )

    // Regenerate styles if any sheet modified or metadata changed (style indices may change)
    if tracker.modifiedSheets.nonEmpty || tracker.modifiedMetadata then
      regenerate += "xl/styles.xml"

    // Regenerate SST if any sheet modified or metadata changed (string indices may change)
    if tracker.modifiedSheets.nonEmpty || tracker.modifiedMetadata then
      regenerate += "xl/sharedStrings.xml"

    // When metadata is modified or sheets are reordered, regenerate ALL sheets.
    // New sheets don't exist in source ZIP, removed sheets change indices,
    // and reordered sheets must be written to new positions.
    if tracker.modifiedMetadata || tracker.reorderedSheets then
      // Metadata modified → regenerate all sheets and their relationships
      wb.sheets.indices.foreach { idx =>
        regenerate += s"xl/worksheets/sheet${idx + 1}.xml"
        regenerate += s"xl/worksheets/_rels/sheet${idx + 1}.xml.rels"
      }
    else
      // No metadata change → only regenerate modified sheets (surgical)
      tracker.modifiedSheets.foreach { idx =>
        regenerate += s"xl/worksheets/sheet${idx + 1}.xml"
      }

      // Regenerate relationships for modified/deleted sheets
      (tracker.modifiedSheets ++ tracker.deletedSheets).foreach { idx =>
        regenerate += s"xl/worksheets/_rels/sheet${idx + 1}.xml.rels"
      }

    regenerate.toSet

  private def usingOrThrow[A](result: Try[A]): A =
    result match
      case Success(value) => value
      case Failure(exception) => throw exception

  // ========== Drawing layer (GH-221) ==========

  /**
   * Everything the drawing layer contributes to one write: regenerated/fresh drawing parts, their
   * rels, new media, source paths superseded by regeneration, first-drawing worksheet wiring, and
   * the ContentTypes raw material. Pure value computed once per write.
   */
  private final case class DrawingWritePlan(
    parts: Vector[(String, OoxmlDrawing)],
    rels: Vector[(String, Relationships)],
    media: Vector[(String, ArraySeq[Byte])],
    skipPaths: Set[String],
    drawingRefs: Map[Int, Elem],
    sheetRelAdditions: Map[Int, Relationship],
    freshPartPaths: Set[String],
    allPartPaths: Set[String],
    imageDefaults: Map[String, String],
    manifestImageDefaults: Map[String, String],
    chartParts: Vector[(String, OoxmlChart)],
    freshChartPaths: Set[String],
    allChartPartPaths: Set[String]
  )

  private def drawingRelsPathOf(partPath: String): String =
    val slash = partPath.lastIndexOf('/')
    val (dir, name) = partPath.splitAt(slash + 1)
    s"${dir}_rels/$name.rels"

  private def fileNameOf(path: String): String =
    path.split('/').lastOption.getOrElse(path)

  private def extensionOf(path: String): String =
    val name = fileNameOf(path)
    val dot = name.lastIndexOf('.')
    if dot >= 0 && dot < name.length - 1 then name.substring(dot + 1) else ""

  private def sha256Hex(bytes: Array[Byte]): String =
    MessageDigest.getInstance("SHA-256").digest(bytes).map("%02x".format(_)).mkString

  /** Generated `<drawing r:id="rIdDr1"/>` element; GH-291 hoisting carries the r binding. */
  private def drawingRefElem: Elem =
    Elem(
      null,
      "drawing",
      new PrefixedAttribute("r", "id", "rIdDr1", Null),
      NamespaceBinding("r", XmlUtil.nsRelationships, TopScope),
      minimizeEmpty = true
    )

  /**
   * Plan all drawing-layer output for this write (GH-221).
   *
   * The hinge is the snapshot-equality dirty test, computed independently of sheetsToRegenerate: a
   * sheet's drawing part is regenerated IFF its `drawings` vector differs from the as-parsed
   * snapshot (reference-equality fast path for untouched sheets). Clean sheets contribute nothing —
   * their parts ride the byte-preservation copy loop, which is what keeps FixturePreservationSpec
   * green by construction.
   *
   * Dirty with a source part: regenerate at the SAME path (worksheet `<drawing r:id>` and sheet
   * rels stay valid), keep ALL source relationships verbatim, sha256-match picture bytes against
   * source image targets (reuse rId), else append `rId{max+1}` and write new media
   * `xl/media/image{M}.{ext}` (M = manifest max + k, ordered by need, sha-deduped per write).
   *
   * First drawings on a sheet (no source part / fresh workbook): a new part `drawing{N}.xml`, a
   * fresh rels file, a `rIdDr1` sheet-rel addition and a generated worksheet `<drawing>` element.
   * Preserved fragments carrying relationship references are dropped on this path (their targets do
   * not exist without the source part's rels — documented in LIMITATIONS).
   *
   * When sheets were deleted or reordered the index-keyed mappings are unreliable (the
   * commentPathMapping precedent has the same shape); drawing regeneration is skipped for that
   * write and parts ride preservation unchanged.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def planDrawingWrites(
    workbook: Workbook,
    sourceContext: Option[SourceContext],
    sheetsToRegenerate: Set[Int]
  ): DrawingWritePlan =
    import com.tjclp.xl.drawings.Drawing as DomainDrawing

    val manifestPaths: Set[String] =
      sourceContext.map(_.partManifest.entries.keySet).getOrElse(Set.empty)
    val manifestDrawingParts: Set[String] =
      manifestPaths.filter(_.matches("xl/drawings/drawing\\d+\\.xml"))
    val manifestChartParts: Set[String] =
      manifestPaths.filter(_.matches("xl/charts/chart\\d+\\.xml"))
    val manifestImageDefaults: Map[String, String] =
      manifestPaths
        .filter(_.startsWith("xl/media/"))
        .flatMap { p =>
          val ext = extensionOf(p).toLowerCase
          ImageFormat.fromExtension(ext).map(f => ext -> f.contentType)
        }
        .toMap

    val mappingsUnreliable = sourceContext.exists { ctx =>
      ctx.modificationTracker.deletedSheets.nonEmpty || ctx.modificationTracker.reorderedSheets
    }

    val dirtyIndices: Vector[Int] = sourceContext match
      case Some(_) if mappingsUnreliable => Vector.empty
      case Some(ctx) =>
        workbook.sheets.indices.toVector.filter { idx =>
          workbook.sheets(idx).drawings != ctx.drawingSnapshots.getOrElse(idx, Vector.empty)
        }
      case None =>
        workbook.sheets.indices.toVector.filter(idx => workbook.sheets(idx).drawings.nonEmpty)

    val maxDrawingNum = manifestDrawingParts
      .flatMap("""drawing(\d+)\.xml""".r.findFirstMatchIn(_))
      .map(_.group(1).toInt)
      .maxOption
      .getOrElse(0)
    val maxImageNum = manifestPaths
      .flatMap("""^xl/media/image(\d+)\.""".r.findFirstMatchIn(_))
      .map(_.group(1).toInt)
      .maxOption
      .getOrElse(0)
    val maxChartNum = manifestChartParts
      .flatMap("""chart(\d+)\.xml""".r.findFirstMatchIn(_))
      .map(_.group(1).toInt)
      .maxOption
      .getOrElse(0)

    var nextDrawing = maxDrawingNum + 1
    var nextImage = maxImageNum + 1
    var nextChart = maxChartNum + 1
    val newMediaByKey = mutable.Map.empty[(String, String), String] // (sha, ext) -> media path
    val mediaOut = Vector.newBuilder[(String, ArraySeq[Byte])]
    val partsOut = Vector.newBuilder[(String, OoxmlDrawing)]
    val relsOut = Vector.newBuilder[(String, Relationships)]
    val skip = Set.newBuilder[String]
    val refs = Map.newBuilder[Int, Elem]
    val relAdds = Map.newBuilder[Int, Relationship]
    val fresh = Set.newBuilder[String]
    val usedExts = mutable.Map.empty[String, String]
    val chartPartsOut = Vector.newBuilder[(String, OoxmlChart)]
    val freshCharts = Set.newBuilder[String]

    /**
     * Chart planning for one drawing part (GH-222). Per ChartFrame, in anchor order:
     *   1. EQUALITY MATCH (consume-once) against the sheet's chart snapshots, preferring the same
     *      anchorIdx: reuse the snapshot relId; the chart part is untouched and rides
     *      byte-preservation (the source rels kept verbatim make the rId resolve by construction).
     *      Robust under drawings-vector reorder/insert — the sha256-media analogue.
     *   2. No equality match but an unconsumed snapshot at the SAME anchorIdx → edited chart:
     *      regenerate at the snapshot's partPath with the same relId (no part churn).
     *   3. Otherwise → fresh `xl/charts/chartN.xml` + an appended rel via `allocRel`.
     * Leftover snapshots' parts stay on disk (the never-delete media policy; orphaned rels are
     * legal OOXML).
     */
    def planChartParts(
      drawings: Vector[DomainDrawing],
      snapshots: Vector[com.tjclp.xl.context.ChartSnapshot],
      allocRel: String => String
    ): Map[Int, String] =
      val consumed = mutable.Set.empty[Int]
      def firstUnconsumed(
        p: com.tjclp.xl.context.ChartSnapshot => Boolean
      ): Option[Int] =
        snapshots.indices.find(i => !consumed.contains(i) && p(snapshots(i)))
      val ids = Map.newBuilder[Int, String]
      drawings.zipWithIndex.foreach {
        case (frame: DomainDrawing.ChartFrame, anchorIdx) =>
          firstUnconsumed(s => s.anchorIdx == anchorIdx && s.chart == frame.chart)
            .orElse(firstUnconsumed(_.chart == frame.chart)) match
            case Some(i) =>
              consumed += i
              ids += anchorIdx -> snapshots(i).relId
            case None =>
              firstUnconsumed(_.anchorIdx == anchorIdx) match
                case Some(i) =>
                  consumed += i
                  val snap = snapshots(i)
                  skip += snap.partPath
                  chartPartsOut += snap.partPath ->
                    OoxmlChart(frame.chart, ChartCaches.resolve(workbook, frame.chart))
                  ids += anchorIdx -> snap.relId
                case None =>
                  val partPath = s"xl/charts/chart$nextChart.xml"
                  nextChart += 1
                  chartPartsOut += partPath ->
                    OoxmlChart(frame.chart, ChartCaches.resolve(workbook, frame.chart))
                  freshCharts += partPath
                  ids += anchorIdx -> allocRel(s"../charts/${fileNameOf(partPath)}")
        case _ => ()
      }
      ids.result()

    def allocMedia(image: com.tjclp.xl.drawings.ImageData): String =
      val key = (image.sha256, image.format.extension)
      newMediaByKey.getOrElseUpdate(
        key, {
          val path = s"xl/media/image$nextImage.${image.format.extension}"
          nextImage += 1
          mediaOut += path -> image.bytes
          usedExts.update(image.format.extension, image.format.contentType)
          path
        }
      )

    def emitFreshPart(idx: Int, drawings: Vector[DomainDrawing]): Unit =
      // Rel-referencing Preserved fragments are dropped: no source rels exist to resolve them.
      // Typed ChartFrames are FIRST-CLASS here (GH-222): self-contained, each allocates a fresh
      // chart part + rels entry.
      val keep = drawings.filter {
        case _: DomainDrawing.Picture => true
        case _: DomainDrawing.ChartFrame => true
        case DomainDrawing.Preserved(xml) => !OoxmlDrawing.hasRelationshipRefs(xml)
      }
      if keep.nonEmpty then
        val partPath = s"xl/drawings/drawing$nextDrawing.xml"
        nextDrawing += 1
        var relCounter = 0
        val partRels = Vector.newBuilder[Relationship]
        val targetToRel = mutable.Map.empty[String, String]
        val embeds = Map.newBuilder[Int, String]
        keep.zipWithIndex.foreach {
          case (p: DomainDrawing.Picture, i) =>
            val target = s"../media/${fileNameOf(allocMedia(p.image))}"
            val relId = targetToRel.getOrElseUpdate(
              target, {
                relCounter += 1
                val id = s"rId$relCounter"
                partRels += Relationship(id, XmlUtil.relTypeImage, target)
                id
              }
            )
            embeds += i -> relId
          case _ => ()
        }
        val chartRelIds = planChartParts(
          keep,
          Vector.empty, // fresh part: no snapshots, every chart is fresh
          target => {
            relCounter += 1
            val id = s"rId$relCounter"
            partRels += Relationship(id, XmlUtil.relTypeChart, target)
            id
          }
        )
        partsOut += partPath -> OoxmlDrawing.build(keep, embeds.result(), chartRelIds, None)
        val rels = partRels.result()
        if rels.nonEmpty then relsOut += drawingRelsPathOf(partPath) -> Relationships(rels)
        fresh += partPath
        refs += idx -> drawingRefElem
        relAdds += idx -> Relationship(
          "rIdDr1",
          XmlUtil.relTypeDrawing,
          s"../drawings/${fileNameOf(partPath)}"
        )

    def emitSamePathPart(ctx: SourceContext, idx: Int, partPath: String): Unit =
      val sheet = workbook.sheets(idx)
      val relsPath = drawingRelsPathOf(partPath)
      val sourceRelsOpt: Option[Relationships] =
        Try(
          withZipFile(ctx.sourcePath)(z => parseOptionalEntry(z, relsPath)(Relationships.fromXml))
        ).toOption.flatten
      val sourceRels = sourceRelsOpt.getOrElse(Relationships.empty)
      val sourceScope: Option[NamespaceBinding] =
        Try(withZipFile(ctx.sourcePath)(z => readZipEntry(z, partPath))).toOption.flatten
          .flatMap(xml => XmlSecurity.parseSafe(xml, partPath).toOption)
          .map(_.scope)
      // sha256 of each source image-rel target (hashed on demand, only on this dirty path)
      val shaToSourceRel: Map[String, String] =
        sourceRels.relationships
          .filter(_.`type` == XmlUtil.relTypeImage)
          .flatMap { r =>
            DrawingReader.resolveMediaTarget(r.target).flatMap { mediaPath =>
              Try(withZipFile(ctx.sourcePath) { z =>
                Option(z.getEntry(mediaPath)).map(e => z.getInputStream(e).readAllBytes())
              }).toOption.flatten.map { bytes =>
                ImageFormat
                  .fromExtension(extensionOf(mediaPath))
                  .foreach(f => usedExts.update(extensionOf(mediaPath).toLowerCase, f.contentType))
                sha256Hex(bytes) -> r.id
              }
            }
          }
          .toMap
      var relCounter =
        sourceRels.relationships.flatMap(_.id.stripPrefix("rId").toIntOption).maxOption.getOrElse(0)
      val appended = Vector.newBuilder[Relationship]
      val targetToRel = mutable.Map.empty[String, String]
      val embeds = Map.newBuilder[Int, String]
      sheet.drawings.zipWithIndex.foreach {
        case (p: DomainDrawing.Picture, anchorIdx) =>
          shaToSourceRel.get(p.image.sha256) match
            case Some(relId) => embeds += anchorIdx -> relId
            case None =>
              val target = s"../media/${fileNameOf(allocMedia(p.image))}"
              val relId = targetToRel.getOrElseUpdate(
                target, {
                  relCounter += 1
                  val id = s"rId$relCounter"
                  appended += Relationship(id, XmlUtil.relTypeImage, target)
                  id
                }
              )
              embeds += anchorIdx -> relId
        case _ => ()
      }
      // GH-222: chart planning (equality-match reuse / same-anchor regen / fresh) appends chart
      // rels through the same counter machinery
      val chartRelIds = planChartParts(
        sheet.drawings,
        ctx.chartSnapshots.getOrElse(idx, Vector.empty),
        target => {
          relCounter += 1
          val id = s"rId$relCounter"
          appended += Relationship(id, XmlUtil.relTypeChart, target)
          id
        }
      )
      // ALL source relationships are kept verbatim: preserved fragments' embedded rIds resolve by
      // construction, and orphaned image/chart rels are legal (their parts are still byte-copied —
      // parts are never deleted, the 6a media policy).
      val finalRels = Relationships(sourceRels.relationships ++ appended.result())
      partsOut += partPath -> OoxmlDrawing.build(
        sheet.drawings,
        embeds.result(),
        chartRelIds,
        sourceScope
      )
      skip += partPath
      if sourceRelsOpt.isDefined then skip += relsPath
      if finalRels.relationships.nonEmpty then relsOut += relsPath -> finalRels

    dirtyIndices.foreach { idx =>
      sourceContext match
        case Some(ctx) =>
          ctx.drawingPathMapping.get(idx) match
            case Some(partPath) => emitSamePathPart(ctx, idx, partPath)
            case None =>
              // First drawings on this sheet. Only wire when the worksheet is regenerated (it
              // must carry the <drawing> ref); untracked direct-Sheet surgery already loses
              // modification tracking for cells too — same documented caveat.
              if sheetsToRegenerate.contains(idx) then
                emitFreshPart(idx, workbook.sheets(idx).drawings)
        case None =>
          emitFreshPart(idx, workbook.sheets(idx).drawings)
    }

    val freshPaths = fresh.result()
    val freshChartPathSet = freshCharts.result()
    DrawingWritePlan(
      parts = partsOut.result(),
      rels = relsOut.result(),
      media = mediaOut.result(),
      skipPaths = skip.result(),
      drawingRefs = refs.result(),
      sheetRelAdditions = relAdds.result(),
      freshPartPaths = freshPaths,
      allPartPaths = manifestDrawingParts ++ freshPaths,
      imageDefaults = usedExts.toMap,
      manifestImageDefaults = manifestImageDefaults,
      chartParts = chartPartsOut.result(),
      freshChartPaths = freshChartPathSet,
      allChartPartPaths = manifestChartParts ++ freshChartPathSet
    )

  /** Write a raw binary part (media bytes) to the output zip. */
  private def writeBinaryPart(
    zip: ZipOutputStream,
    entryName: String,
    bytes: Array[Byte],
    config: WriterConfig
  ): Unit =
    val entry = new ZipEntry(entryName)
    entry.setTime(0L)
    entry.setMethod(config.compression.zipMethod)
    config.compression match
      case Compression.Stored =>
        entry.setSize(bytes.length)
        entry.setCompressedSize(bytes.length)
        entry.setCrc(calculateCrc(bytes))
      case Compression.Deflated => ()
    zip.putNextEntry(entry)
    zip.write(bytes)
    zip.closeEntry()

  /** Helper to open a ZipFile with automatic resource management. */
  private def withZipFile[A](path: Path)(f: ZipFile => A): A =
    usingOrThrow(Using.Manager { use =>
      val zip = use(new ZipFile(path.toFile))
      f(zip)
    })

  /**
   * GH-235: external-hyperlink relationships for a sheet (rIdHL{n} <-> URL), matching the
   * worksheet.
   */
  private def hyperlinkRelationships(sheet: Sheet): Seq[Relationship] =
    collectHyperlinks(sheet).filter(_.external).map { e =>
      Relationship(e.relId, XmlUtil.relTypeHyperlink, e.target, Some("External"))
    }

  /** Read the contents of a ZIP entry as UTF-8 string if it exists. */
  private def readZipEntry(zip: ZipFile, entryName: String): Option[String] =
    Option(zip.getEntry(entryName)).map { entry =>
      usingOrThrow(Using.Manager { use =>
        val in = use(zip.getInputStream(entry))
        new String(in.readAllBytes(), StandardCharsets.UTF_8)
      })
    }

  private def parseOptionalEntry[T](
    zip: ZipFile,
    entryName: String
  )(parse: Elem => Either[String, T]): Option[T] =
    readZipEntry(zip, entryName).flatMap { xmlString =>
      XmlSecurity.parseSafe(xmlString, entryName).toOption.flatMap(elem => parse(elem).toOption)
    }

  /**
   * Copy a single preserved part from source ZIP to output ZIP.
   *
   * Streams bytes in 8KB chunks (constant memory). Preserves compression method and metadata.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def copyPreservedPart(
    sourcePath: Path,
    partPath: String,
    outputZip: ZipOutputStream
  ): Unit =
    withZipFile(sourcePath) { sourceZip =>
      val entry = Option(sourceZip.getEntry(partPath)).getOrElse {
        throw new IllegalStateException(s"Entry missing from source: $partPath")
      }

      val newEntry = new ZipEntry(partPath)
      newEntry.setTime(0L)
      newEntry.setMethod(entry.getMethod)

      if entry.getMethod == ZipEntry.STORED then
        newEntry.setSize(entry.getSize)
        newEntry.setCompressedSize(entry.getCompressedSize)
        newEntry.setCrc(entry.getCrc)

      var entryOpen = false
      usingOrThrow(Using.Manager { use =>
        val in = use(sourceZip.getInputStream(entry))
        try
          outputZip.putNextEntry(newEntry)
          entryOpen = true
          val buffer = new Array[Byte](8192)
          var read = in.read(buffer)
          while read != -1 do
            outputZip.write(buffer, 0, read)
            read = in.read(buffer)
        finally if entryOpen then outputZip.closeEntry()
      })
    }

  /**
   * Re-parse structural files from source ZIP for preservation.
   *
   * Returns (ContentTypes, rootRels, workbookRels, workbook) parsed from the original file. If any
   * file is missing or fails to parse, returns None for that component.
   */
  private def parsePreservedStructure(
    sourcePath: Path
  ): (Option[ContentTypes], Option[Relationships], Option[Relationships], Option[OoxmlWorkbook]) =
    withZipFile(sourcePath) { zip =>
      val contentTypes = parseOptionalEntry(zip, "[Content_Types].xml")(ContentTypes.fromXml)
      val rootRels = parseOptionalEntry(zip, "_rels/.rels")(Relationships.fromXml)
      val workbookRels =
        parseOptionalEntry(zip, "xl/_rels/workbook.xml.rels")(Relationships.fromXml)
      val workbook = parseOptionalEntry(zip, "xl/workbook.xml")(OoxmlWorkbook.fromXml)
      (contentTypes, rootRels, workbookRels, workbook)
    }

  /**
   * Parse worksheet XML from source ZIP to extract preserved metadata.
   *
   * Returns the parsed OoxmlWorksheet with all metadata (cols, views, etc.) for merging during
   * regeneration.
   */
  private def parsePreservedWorksheet(
    sourcePath: Path,
    sheetPath: String
  ): XLResult[Option[OoxmlWorksheet]] =
    Try(withZipFile(sourcePath)(zip => readZipEntry(zip, sheetPath))) match
      case Failure(e) =>
        Left(XLError.IOError(s"Failed to read preserved worksheet $sheetPath: ${e.getMessage}"))
      case Success(None) => Right(None)
      case Success(Some(xmlString)) =>
        for
          elem <- XmlSecurity.parseSafe(xmlString, sheetPath)
          worksheet <- OoxmlWorksheet
            .fromXml(elem)
            .left
            .map(err => XLError.ParseError(sheetPath, err): XLError)
        yield Some(worksheet)

  /**
   * Parse SharedStrings from source ZIP for modified sheet cell encoding.
   *
   * CRITICAL: Modified sheets need the SST to encode cells as t="s" (SST references) instead of
   * t="inlineStr" (inline strings). We copy the SST verbatim to output, but parse it here so
   * modified sheets can reference the same indices as unmodified sheets.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  /**
   * Count shared-string references (sheetData cells with t="s") in a SOURCE worksheet (GH-304).
   *
   * Used to subtract a modified sheet's pre-edit contribution from the preserved SST count
   * attribute, so the sheet's actual post-edit references can be added back. Returns 0 when the
   * part is missing or unparseable (a new sheet, or a degenerate source).
   */
  private def countSourceSstReferences(sourcePath: Path, sheetPath: String): Int =
    Try(withZipFile(sourcePath)(zip => readZipEntry(zip, sheetPath))) match
      case Success(Some(xmlString)) =>
        XmlSecurity
          .parseSafe(xmlString, sheetPath)
          .toOption
          .map { elem =>
            (elem \ "sheetData" \ "row" \ "c").count(c => (c \ "@t").text == "s")
          }
          .getOrElse(0)
      case _ => 0

  private def parsePreservedSST(sourcePath: Path): Option[SharedStrings] =
    Try(withZipFile(sourcePath)(zip => readZipEntry(zip, "xl/sharedStrings.xml"))) match
      case Failure(_) => None
      case Success(None) => None
      case Success(Some(xmlString)) =>
        XmlSecurity
          .parseSafe(xmlString, "xl/sharedStrings.xml")
          .toOption
          .flatMap(SharedStrings.fromXml(_).toOption)

  /**
   * Parse preserved styles.xml to extract namespace metadata and differential formats.
   *
   * Critical for preserving Excel extension namespaces (x14ac, x16r2, xr, mc:Ignorable) and
   * differential formats used by conditional formatting.
   */
  private def parsePreservedStylesMetadata(
    sourcePath: Path
  ): (Option[MetaData], NamespaceBinding, Option[Elem]) =
    Try(withZipFile(sourcePath)(zip => readZipEntry(zip, "xl/styles.xml"))) match
      case Failure(_) => (None, TopScope, None)
      case Success(None) => (None, TopScope, None)
      case Success(Some(xmlString)) =>
        XmlSecurity.parseSafe(xmlString, "xl/styles.xml").toOption match
          case Some(elem) =>
            val dxfs = (elem \ "dxfs").headOption.collect { case e: Elem => e }
            (Some(elem.attributes), elem.scope, dxfs)
          case None => (None, TopScope, None)

  /**
   * Unified write: intelligently regenerates only what changed, preserves the rest.
   *
   * Strategy:
   *   - With source (surgical mode): Regenerate modified sheets, copy unmodified, preserve unknown
   *     parts
   *   - Without source (new workbook): Regenerate everything (graceful degradation)
   *
   * This unified path handles both surgical modification and normal writes.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def unifiedWrite(
    workbook: Workbook,
    sourceContext: Option[SourceContext],
    target: OutputTarget,
    config: WriterConfig
  ): Unit =
    val escapeFormulas = formulaEscapingRequested(config)

    // Determine modification tracking (all sheets modified if no source)
    val tracker = sourceContext.map(_.modificationTracker).getOrElse {
      ModificationTracker(
        modifiedSheets = workbook.sheets.indices.toSet,
        deletedSheets = Set.empty,
        reorderedSheets = false
      )
    }

    val (graph, preservableParts, regenerateParts) = sourceContext match
      case Some(ctx) =>
        val g = RelationshipGraph.fromManifest(ctx.partManifest)
        val pres = determinePreservableParts(workbook, ctx, g)
        val regen = determineRegenerateParts(workbook, ctx)
        (g, pres, regen)
      case None =>
        (RelationshipGraph.empty, Set.empty[String], Set.empty[String])

    // Escaping must be applied to every text-bearing worksheet part; preserved sheets can contain
    // dangerous inline strings or shared-string references that would bypass WriterConfig.secure.
    val sheetsToRegenerate =
      if escapeFormulas || tracker.modifiedMetadata || tracker.reorderedSheets then
        workbook.sheets.indices.toSet
      else tracker.modifiedSheets

    val sharedStringsPath = "xl/sharedStrings.xml"
    val sourceHasSharedStrings = sourceContext.exists(_.partManifest.contains(sharedStringsPath))

    // SST Strategy:
    // - If modified sheets have new strings not in preserved SST → regenerate (to avoid inline string corruption)
    // - If modified sheets only use existing SST strings → copy verbatim (byte-perfect preservation)
    // - If no source SST → generate according to policy
    val (sst, regenerateSharedStrings) =
      if escapeFormulas && sourceHasSharedStrings then
        (Some(SharedStrings.fromWorkbook(workbook, escapeFormulas = true)), true)
      else if sourceHasSharedStrings then
        // Parse preserved SST (sourceContext guaranteed to exist if sourceHasSharedStrings is true)
        val parsedSST = sourceContext.map(ctx => parsePreservedSST(ctx.sourcePath)).getOrElse(None)

        // Check if modified sheets contain NEW strings not in the preserved SST.
        // Entries are collected PER REFERENCE (one per text cell, GH-277: the SST count attribute
        // counts references) in deterministic ref order, and compared EXACTLY — no Unicode
        // normalization, byte-different spellings are different strings (GH-277/GH-289).
        val modifiedSheetRefEntries: Vector[Either[String, RichText]] =
          tracker.modifiedSheets.toVector.sorted.flatMap { idx =>
            workbook.sheets.lift(idx).toList.flatMap { sheet =>
              sheet.cells.toVector.sortBy(_._1.toA1).flatMap { case (_, cell) =>
                cell.value match
                  case CellValue.Text(str) => Some(Left(str))
                  case CellValue.RichText(rt) => Some(Right(rt))
                  case _ => None
              }
            }
          }

        val preservedEntries: Set[Either[String, RichText]] =
          parsedSST.map(_.strings.toSet).getOrElse(Set.empty)

        // New entries in deterministic first-reference order, storing ORIGINAL strings —
        // normalization is for comparison only, never storage (GH-277)
        val newEntries =
          modifiedSheetRefEntries.distinct.filterNot(preservedEntries.contains)

        // GH-304: the count attribute counts REFERENCES, and edits can REMOVE or DUPLICATE
        // references to existing strings, not just add new ones. Recount the modified sheets'
        // contribution exactly: preserved sheets keep their counted contribution (original count
        // minus the modified sheets' pre-edit t="s" cells), while modified sheets contribute one
        // reference per post-edit text cell (regenerated sheets encode ALL text via the SST).
        val preEditModifiedRefs = sourceContext
          .map { ctx =>
            tracker.modifiedSheets.toVector.sorted
              .map(idx => countSourceSstReferences(ctx.sourcePath, graph.pathForSheet(idx)))
              .sum
          }
          .getOrElse(0)
        val exactTotalCount = parsedSST match
          case Some(preserved) =>
            math.max(0, preserved.totalCount - preEditModifiedRefs + modifiedSheetRefEntries.size)
          case None => modifiedSheetRefEntries.size

        if newEntries.nonEmpty then
          // New strings detected → build SST from preserved + new entries only
          // Do NOT use SharedStrings.fromWorkbook (includes ALL sheets, even binary system sheets!)
          val combinedEntries =
            parsedSST.map(_.strings).getOrElse(Vector.empty) ++ newEntries

          // Create new SST with combined entries (exact-string index keys, GH-289)
          val combinedSST = SharedStrings(
            strings = combinedEntries,
            indexMap = combinedEntries.zipWithIndex.map {
              case (Left(s), idx) => s -> idx
              case (Right(rt), idx) => rt.toPlainText -> idx
            }.toMap,
            totalCount = exactTotalCount
          )
          (Some(combinedSST), true)
        else
          parsedSST match
            case Some(preserved) if preserved.totalCount != exactTotalCount =>
              // GH-304: same strings, changed reference count → re-emit the SST with the exact
              // count. Entries are never pruned or reordered: preserved sheets' t="s" indices
              // must stay valid.
              (Some(preserved.copy(totalCount = exactTotalCount)), true)
            case _ =>
              // No count drift → copy preserved SST verbatim for byte-perfect preservation
              (parsedSST, false)
      else
        // No source SST - generate if policy allows
        val generated = config.sstPolicy match
          case SstPolicy.Always => Some(SharedStrings.fromWorkbook(workbook, escapeFormulas))
          case SstPolicy.Never => None
          case SstPolicy.Auto =>
            if SharedStrings.shouldUseSST(workbook) then
              Some(SharedStrings.fromWorkbook(workbook, escapeFormulas))
            else None
        (generated, generated.isDefined)

    val sharedStringsInOutput = sourceHasSharedStrings || regenerateSharedStrings
    // Regenerated sheets ALWAYS encode against whatever SST ships in the output — preserved,
    // combined, or freshly generated. The old `if regenerateSharedStrings then sst else None`
    // orphaned a verbatim-copied SST: modified sheets re-inlined every string while
    // sharedStrings.xml still shipped (GH-277).
    val sstForSheets = sst

    // Build style index (automatic optimization based on sourceContext). Any sheet regenerated
    // from a source-backed workbook needs a local-to-workbook style remapping.
    val (styleIndex, sheetRemappings) =
      StyleIndex.fromWorkbook(workbook, sheetsRequiringRemapping = sheetsToRegenerate)

    // Parse preserved styles metadata (namespaces and dxfs) if source available
    val (preservedStylesAttrs, preservedStylesScope, preservedDxfs) = sourceContext match
      case Some(ctx) => parsePreservedStylesMetadata(ctx.sourcePath)
      case None => (None, TopScope, None)

    val styles = OoxmlStyles(styleIndex, preservedStylesAttrs, preservedStylesScope, preservedDxfs)

    // Build comments data and VML drawings
    val (commentsBySheet, sheetsWithComments) = buildCommentsData(workbook)
    val vmlDrawings = commentsBySheet.map { case (idx, comments) =>
      idx -> VmlDrawing.generateForComments(comments, idx)
    }

    // Build table data
    val (tablesBySheet, totalTableCount, tableIdMap) = buildTablesData(workbook)

    // GH-221: drawing-layer plan — snapshot-equality dirty test, media dedup, first-drawing wiring
    val drawingPlan = planDrawingWrites(workbook, sourceContext, sheetsToRegenerate)

    // Preserve structural parts from source (or fallback to minimal)
    // IMPORTANT: When metadata is modified (add/remove/rename/reorder sheets),
    // we MUST regenerate the structural parts since sheet count/order changed.
    val (preservedContentTypes, preservedRootRels, preservedWorkbookRels, preservedWorkbook) =
      sourceContext match
        case Some(ctx) if !tracker.modifiedMetadata => parsePreservedStructure(ctx.sourcePath)
        case _ => (None, None, None, None)

    // GH-314: the preserved [Content_Types].xml matters even when metadata changed — exotic
    // preserved parts (pivots, custom XML, macro payloads) still ride the verbatim copy loop and
    // must stay registered. Parsed here so the regenerate-from-minimal branch below can reconcile
    // instead of dropping their overrides.
    val preservedContentTypesForReconcile: Option[ContentTypes] =
      preservedContentTypes.orElse {
        sourceContext.flatMap { ctx =>
          withZipFile(ctx.sourcePath) { zip =>
            parseOptionalEntry(zip, "[Content_Types].xml")(ContentTypes.fromXml)
          }
        }
      }

    // Use preserved workbook structure if available, otherwise create minimal
    val ooxmlWb = preservedWorkbook match
      case Some(preserved) =>
        // Update sheets in preserved structure (names/order/visibility may have changed).
        // Named ranges (definedNames) ride through verbatim here (byte-identical). Authoring a
        // name marks metadata modified, which routes to the fromDomain branch below (GH-236).
        val updated = preserved.updateSheets(workbook.sheets, workbook.metadata.sheetStates)
        // GH-259: print names (_xlnm.Print_Area/_xlnm.Print_Titles) are modeled on Sheet.pageSetup,
        // so a sheet edit can change them WITHOUT marking metadata modified. Reconcile: keep the
        // preserved definedNames bytes when the model agrees, otherwise regenerate the element.
        val expected = PrintNames.effective(workbook)
        if OoxmlWorkbook.parseDefinedNames(updated.definedNames).toSet == expected.toSet then
          updated
        else updated.copy(definedNames = OoxmlWorkbook.buildDefinedNames(expected))
      case None =>
        // Fallback for programmatically created workbooks OR when metadata was modified — fresh
        // workbook structure; fromDomain serializes wb.metadata.definedNames (GH-236) plus the
        // PageSetup-derived print names (GH-259).
        OoxmlWorkbook.fromDomain(workbook)

    // GH-242: document properties are model-driven. The reader parses docProps/core.xml and
    // app.xml into WorkbookMetadata (and marks them parsed, so they are never copied verbatim);
    // here the model decides what ships — emitted iff at least one field is present.
    val corePropsXml = DocProps.buildCoreXml(workbook.metadata)
    val appPropsXml = DocProps.buildAppXml(workbook.metadata)

    // Content types: preserve from source when available, otherwise generate minimal.
    // IMPORTANT: Don't call withCommentOverrides when preserving - the source already has
    // correct comment entries with Excel's sequential numbering (comments1.xml, comments2.xml...)
    // which differs from sheet-index-based numbering.
    val contentTypes = preservedContentTypes match
      case Some(preserved) =>
        // Add sharedStrings override if we're generating it but source didn't have it
        val withSst =
          if sharedStringsInOutput && !sourceHasSharedStrings then
            preserved.copy(overrides =
              preserved.overrides + ("/xl/sharedStrings.xml" -> XmlUtil.ctSharedStrings)
            )
          else preserved
        // Only add table overrides (comments already in preserved Content_Types).
        // GH-221: source drawing overrides/media defaults are already in the preserved types;
        // register only fresh parts and the media extensions this write touches (idempotent).
        withSst
          .withTableOverrides(totalTableCount)
          .withDocPropsOverrides(corePropsXml.isDefined, appPropsXml.isDefined)
          .withDrawingOverrides(drawingPlan.freshPartPaths)
          .withChartOverrides(drawingPlan.freshChartPaths)
          .withImageDefaults(drawingPlan.imageDefaults)
      case None =>
        // GH-221: this branch regenerates [Content_Types].xml from scratch while preserved
        // drawing/media parts still ride the copy loop, so register EVERY drawing part shipping
        // (manifest + fresh) and every known media extension (manifest + written). GH-222: chart
        // parts follow the same allPartPaths pattern.
        val model = ContentTypes
          .minimal(
            hasStyles = true,
            hasSharedStrings = sharedStringsInOutput,
            sheetCount = workbook.sheets.size,
            sheetsWithComments = sheetsWithComments
          )
          .withCommentOverrides(sheetsWithComments)
          .withTableOverrides(totalTableCount)
          .withDocPropsOverrides(corePropsXml.isDefined, appPropsXml.isDefined)
          .withDrawingOverrides(drawingPlan.allPartPaths)
          .withChartOverrides(drawingPlan.allChartPartPaths)
          .withImageDefaults(drawingPlan.manifestImageDefaults ++ drawingPlan.imageDefaults)
        // GH-314: a source can still exist here (metadata-modified write) — merge its preserved
        // content types so exotic parts riding the copy loop keep their registrations. docProps
        // reconciliation is re-applied AFTER the merge: it removes stale preserved overrides for
        // docProps the model no longer emits (GH-242), which a plain union would resurrect.
        preservedContentTypesForReconcile match
          case Some(preserved) =>
            ContentTypes
              .reconcile(preserved, model)
              .withDocPropsOverrides(corePropsXml.isDefined, appPropsXml.isDefined)
          case None => model

    val rootRels = preservedRootRels
      .getOrElse(Relationships.root())
      .withDocProps(corePropsXml.isDefined, appPropsXml.isDefined)

    val workbookRels = preservedWorkbookRels match
      case Some(preserved) =>
        // Add sharedStrings relationship if we're generating it but source didn't have it
        if sharedStringsInOutput && !sourceHasSharedStrings then
          val nextId = preserved.relationships.size + 1
          preserved.copy(relationships =
            preserved.relationships :+
              Relationship(s"rId$nextId", XmlUtil.relTypeSharedStrings, "sharedStrings.xml")
          )
        else preserved
      case None =>
        Relationships.workbook(
          sheetCount = workbook.sheets.size,
          hasStyles = true,
          hasSharedStrings = sharedStringsInOutput
        )

    // Open output ZIP
    val zip = target match
      case OutputPath(path) => new ZipOutputStream(new FileOutputStream(path.toFile))
      case OutputStreamTarget(stream) => new ZipOutputStream(stream)

    // Set compression level to match Excel (level 1 = super fast, matches original defS)
    zip.setLevel(1)

    try
      // Write structural parts (always regenerated)
      writePart(zip, "[Content_Types].xml", contentTypes, config)
      writePart(zip, "_rels/.rels", rootRels, config)

      // GH-242: document properties — deterministic, model-driven (no GUIDs/wall-clock values).
      // The reader marks docProps as parsed, so these never collide with preserved parts.
      corePropsXml.foreach(x => writePart(zip, DocProps.corePath, x, config))
      appPropsXml.foreach(x => writePart(zip, DocProps.appPath, x, config))

      writePart(zip, "xl/workbook.xml", ooxmlWb, config)
      writePart(zip, "xl/_rels/workbook.xml.rels", workbookRels, config)
      writeStyles(zip, "xl/styles.xml", styles, config)

      // Preserve theme file from source if available
      // Theme is parsed (in "known parts") but not regenerated, so we must copy it explicitly
      val themePath = "xl/theme/theme1.xml"
      sourceContext.foreach { ctx =>
        if ctx.partManifest.contains(themePath) then
          copyPreservedPart(ctx.sourcePath, themePath, zip)
      }

      if regenerateSharedStrings then
        sst.foreach { sharedStrings =>
          writeSharedStrings(zip, sharedStringsPath, sharedStrings, config)
        }
      else if sourceHasSharedStrings then
        sourceContext.foreach(ctx => copyPreservedPart(ctx.sourcePath, sharedStringsPath, zip))

      // Write sheets: regenerate modified, copy unmodified (if source available)
      workbook.sheets.zipWithIndex.foreach { case (sheet, idx) =>
        if sheetsToRegenerate.contains(idx) then
          // Regenerate modified sheet with preserved metadata (if available)
          val sheetPath = graph.pathForSheet(idx)
          val preservedMetadata = sourceContext
            .map(ctx => parsePreservedWorksheet(ctx.sourcePath, sheetPath))
            .getOrElse(Right(None)) match
            case Right(value) => value
            case Left(err) =>
              throw new IllegalStateException(
                s"Failed to parse preserved worksheet $sheetPath: ${err.message}"
              )
          val remapping = sheetRemappings.getOrElse(idx, Map.empty)

          // Generate tableParts XML element for modified sheet
          val tablePartsXml = tablesBySheet.get(idx).flatMap { tablesForSheet =>
            if tablesForSheet.isEmpty then None
            else
              val hasComments = commentsBySheet.contains(idx)
              val rIdOffset = if hasComments then 3 else 1

              val tablePartElems =
                tablesForSheet.sortBy(_._2).zipWithIndex.map { case ((_, tableId), tableIdx) =>
                  import scala.xml.*
                  Elem(
                    prefix = null,
                    label = "tablePart",
                    attributes = new PrefixedAttribute(
                      "r",
                      "id",
                      s"rId${rIdOffset + tableIdx}",
                      Null
                    ),
                    scope = NamespaceBinding("r", XmlUtil.nsRelationships, TopScope),
                    minimizeEmpty = true
                  )
                }

              Some(
                XmlUtil
                  .elem("tableParts", "count" -> tablesForSheet.size.toString)(tablePartElems*)
              )
          }

          // For SaxStax backend with no preserved metadata, use direct SAX emission
          // (bypasses intermediate OOXML types for 5-7x performance improvement).
          // GH-221: sheets with drawings force the OoxmlWorksheet path — DirectSaxEmitter has no
          // <drawing> support yet (deferred).
          (config.backend, preservedMetadata) match
            case (XmlBackend.SaxStax, None) if sheet.drawings.isEmpty =>
              writeWorksheetDirect(
                zip,
                s"xl/worksheets/sheet${idx + 1}.xml",
                sheet,
                sstForSheets,
                remapping,
                tablePartsXml,
                escapeFormulas,
                config
              )
            case _ =>
              val ooxmlSheet =
                OoxmlWorksheet.fromDomainWithMetadata(
                  sheet,
                  sstForSheets,
                  remapping,
                  preservedMetadata,
                  tablePartsXml,
                  escapeFormulas,
                  drawingPlan.drawingRefs.get(idx)
                )
              writeWorksheet(zip, s"xl/worksheets/sheet${idx + 1}.xml", ooxmlSheet, config)

          // For modified sheets:
          // 1. Copy relationships from source (preserves printerSettings, drawings, customProperty)
          // 2. Regenerate comments/VML from domain model (handles comment add/remove/modify)
          val relsPath = s"xl/worksheets/_rels/sheet${idx + 1}.xml.rels"
          val commentPath = sourceContext
            .flatMap(_.commentPathMapping.get(idx))
            .getOrElse(s"xl/comments${idx + 1}.xml")
          val commentFileNum = commentPath.stripPrefix("xl/comments").stripSuffix(".xml")
          val vmlPath = s"xl/drawings/vmlDrawing$commentFileNum.vml"

          val hasComments = commentsBySheet.contains(idx)
          val tableIds = tablesBySheet.get(idx).map(_.map(_._2)).getOrElse(Seq.empty)
          val hlRels = hyperlinkRelationships(sheet) // GH-235

          // Copy relationships from source if available (preserves non-comment relationships).
          // Otherwise regenerate minimal relationships. Authored hyperlinks (hlRels) are merged
          // in, as is a first-drawing rel (rIdDr1, GH-221) when this sheet gained a fresh part.
          val drawingRelAdd = drawingPlan.sheetRelAdditions.get(idx)
          sourceContext match
            case Some(ctx) if ctx.partManifest.contains(relsPath) =>
              if hlRels.isEmpty && drawingRelAdd.isEmpty then
                copyPreservedPart(ctx.sourcePath, relsPath, zip)
              else
                // Merge authored hyperlink rels into the preserved sheet rels (parse + append)
                val preserved = withZipFile(ctx.sourcePath) { z =>
                  parseOptionalEntry(z, relsPath)(Relationships.fromXml)
                }.getOrElse(Relationships(Seq.empty))
                // Drop the source's hyperlink rels (we regenerate them from the model) to avoid
                // orphans, but keep everything else (printerSettings, drawings, ...).
                val kept = preserved.relationships.filterNot(_.`type` == XmlUtil.relTypeHyperlink)
                writePart(
                  zip,
                  relsPath,
                  Relationships(kept ++ hlRels ++ drawingRelAdd.toList),
                  config
                )
            case _
                if hasComments || tableIds.nonEmpty || hlRels.nonEmpty || drawingRelAdd.nonEmpty =>
              val base =
                buildWorksheetRelationshipsWithCommentsPath(commentPath, hasComments, tableIds)
              writePart(
                zip,
                relsPath,
                Relationships(base.relationships ++ hlRels ++ drawingRelAdd.toList),
                config
              )
            case _ => // No relationships needed

          // Always regenerate comments/VML from domain model for modified sheets
          // This ensures comment add/remove/modify is reflected in output
          commentsBySheet.get(idx).foreach { comments =>
            writePart(zip, commentPath, comments, config)

            vmlDrawings.get(idx).foreach { vmlXml =>
              writeVmlPart(zip, vmlPath, vmlXml, config)
            }
          }
        else
          // Copy unmodified sheet from source (only if source available)
          sourceContext.foreach { ctx =>
            val sheetPath = graph.pathForSheet(idx)
            copyPreservedPart(ctx.sourcePath, sheetPath, zip)

            // Copy comments and relationships using source's actual paths (from mapping)
            ctx.commentPathMapping.get(idx).foreach { commentPath =>
              if ctx.partManifest.contains(commentPath) then
                copyPreservedPart(ctx.sourcePath, commentPath, zip)
            }

            val relsPath = s"xl/worksheets/_rels/sheet${idx + 1}.xml.rels"
            if ctx.partManifest.contains(relsPath) then
              copyPreservedPart(ctx.sourcePath, relsPath, zip)
          }
      }

      // Write table files for all sheets (always regenerated from domain model)
      tablesBySheet.values.flatten.foreach { case (tableSpec, tableId) =>
        val ooxmlTable = TableConversions.toOoxml(tableSpec, tableId)
        writePart(zip, s"xl/tables/table$tableId.xml", ooxmlTable, config)
      }

      // GH-221: regenerated/fresh drawing parts, their rels, and new media
      drawingPlan.parts.foreach { case (path, part) => writePart(zip, path, part, config) }
      drawingPlan.rels.foreach { case (path, rels) => writePart(zip, path, rels, config) }
      drawingPlan.media.foreach { case (path, bytes) =>
        writeBinaryPart(zip, path, bytes.toArray, config)
      }
      // GH-222: regenerated/fresh chart parts (same-path regens are in skipPaths)
      drawingPlan.chartParts.foreach { case (path, part) => writePart(zip, path, part, config) }

      // Copy preserved parts (charts, drawings, images, etc.) if source available
      // Skip VML drawings for regenerated sheets - we regenerate VML from domain model
      // (or omit it if comments were removed); skip drawing parts superseded by same-path
      // regeneration (GH-221)
      sourceContext.foreach { ctx =>
        val vmlPathsToSkip = sheetsToRegenerate.flatMap { idx =>
          ctx.commentPathMapping.get(idx).map { commentPath =>
            val fileNum = commentPath.stripPrefix("xl/comments").stripSuffix(".xml")
            s"xl/drawings/vmlDrawing$fileNum.vml"
          }
        }

        preservableParts.foreach { path =>
          val isVmlDrawing = path.startsWith("xl/drawings/vmlDrawing") && path.endsWith(".vml")
          val shouldSkip =
            (isVmlDrawing && vmlPathsToSkip.contains(path)) ||
              drawingPlan.skipPaths.contains(path)

          if !shouldSkip then copyPreservedPart(ctx.sourcePath, path, zip)
        }
      }

    finally zip.close()

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
