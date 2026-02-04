package com.tjclp.xl.ooxml

import scala.xml.*
import java.io.{ByteArrayOutputStream, FileOutputStream, OutputStream}
import java.security.MessageDigest
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}
import java.nio.file.{Files, Path, StandardCopyOption}
import java.nio.charset.StandardCharsets
import scala.collection.mutable
import scala.util.{Failure, Success, Try, Using}
import com.tjclp.xl.api.{Sheet, Workbook, CellValue}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.context.{ModificationTracker, SourceContext}
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.tables.TableSpec

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
      workbook.sourceContext match
        case Some(ctx) if ctx.isClean =>
          // Clean workbook + file target → verbatim copy (ultra-fast)
          target match
            case OutputPath(path) =>
              copyVerbatim(ctx, path)
            case OutputStreamTarget(_) =>
              // Can't copy to stream, use unified write (will copy all parts)
              unifiedWrite(workbook, workbook.sourceContext, target, config)

        case Some(ctx) =>
          // Surgical mode with source: use atomic temp file to prevent corruption
          target match
            case OutputPath(destPath) =>
              writeAtomically(workbook, Some(ctx), destPath, config)
            case _ =>
              unifiedWrite(workbook, workbook.sourceContext, target, config)

        case None =>
          // No source context: safe to write directly (no surgical preservation)
          unifiedWrite(workbook, None, target, config)

      Right(())

    catch case e: Exception => Left(XLError.IOError(s"Failed to write XLSX: ${e.getMessage}"))

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
        // Build author list (deduplicated and sorted for deterministic output)
        // Reserve index 0 for unauthored comments (empty string)
        val (authorSet, hasUnauthored) =
          sheet.comments.values.foldLeft((Set.empty[String], false)) {
            case ((existing, hasNone), comment) =>
              val nextSet = comment.author match
                case Some(author) => existing + author
                case None => existing
              (nextSet, hasNone || comment.author.isEmpty)
          }
        // Sort for deterministic output (Excel preserves insertion order; we choose determinism here)
        // to keep serialized files stable across runs.
        val realAuthors = authorSet.toVector.sorted
        val authors = if hasUnauthored then Vector("") ++ realAuthors else realAuthors

        val authorMap = authors.zipWithIndex.map { case (author, i) => author -> i }.toMap

        // Convert domain Comments to OOXML (sorted by ref for deterministic output)
        val ooxmlComments = sheet.comments.toVector.sortBy(_._1.toA1).map { case (ref, comment) =>
          val authorId =
            comment.author.flatMap(authorMap.get).getOrElse(0) // Index 0 = "" for unauthored
          // Note: xr:uid GUIDs omitted for new comments (deterministic output)
          // GUIDs are optional per OOXML spec and only needed for revision tracking

          // Excel displays author as part of comment text (bold first run)
          val textWithAuthor = comment.author match
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
   * Build worksheet relationships for a sheet with comments.
   *
   * Creates relationships for:
   *   - Comments content (../commentsN.xml)
   *   - VML drawing for visual indicators (../drawings/vmlDrawingN.xml)
   *
   * N is 1-based sheet index.
   */
  private def buildWorksheetRelationships(sheetIndex: Int): Relationships =
    Relationships(
      Seq(
        Relationship(
          id = "rId1",
          `type` = XmlUtil.relTypeComments,
          target = s"../comments$sheetIndex.xml"
        ),
        Relationship(
          id = "rId2",
          `type` = XmlUtil.relTypeVmlDrawing,
          target = s"../drawings/vmlDrawing$sheetIndex.vml"
        )
      )
    )

  /**
   * Build worksheet relationships for a sheet with comments and/or tables.
   *
   * Creates relationships for:
   *   - Comments content (../commentsN.xml) - rId1 (if hasComments)
   *   - VML drawing for visual indicators (../drawings/vmlDrawingN.vml) - rId2 (if hasComments)
   *   - Tables (../tables/tableM.xml) - rId3+ (if hasComments), rId1+ (if tables only)
   *
   * N is 1-based sheet index, M is global table ID.
   *
   * @param sheetIndex
   *   Sheet index (1-based) for file naming
   * @param hasComments
   *   Whether this sheet has comments
   * @param tableIds
   *   Sequence of global table IDs for this sheet
   * @return
   *   Relationships with sequential rId allocation
   */
  private def buildWorksheetRelationshipsWithTables(
    sheetIndex: Int,
    hasComments: Boolean,
    tableIds: Seq[Long]
  ): Relationships =
    val commentPath = s"xl/comments$sheetIndex.xml"
    buildWorksheetRelationshipsWithCommentsPath(commentPath, hasComments, tableIds)

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

  /**
   * Full regeneration of all XLSX parts (current behavior).
   *
   * Used when:
   *   - Workbook has no SourceContext (created programmatically)
   *   - Output target is a stream (can't use file copy optimization)
   *
   * Regenerates all parts from domain model.
   */
  private def regenerateAll(
    workbook: Workbook,
    target: OutputTarget,
    config: WriterConfig
  ): Unit =
    // Build shared strings table based on policy
    val sst = config.sstPolicy match
      case SstPolicy.Always => Some(SharedStrings.fromWorkbook(workbook))
      case SstPolicy.Never => None
      case SstPolicy.Auto =>
        if SharedStrings.shouldUseSST(workbook) then Some(SharedStrings.fromWorkbook(workbook))
        else None

    // Build comments data and VML drawings
    val (commentsBySheet, sheetsWithComments) = buildCommentsData(workbook)
    val vmlDrawings = commentsBySheet.map { case (idx, comments) =>
      idx -> VmlDrawing.generateForComments(comments, idx)
    }

    // Build table data
    val (tablesBySheet, totalTableCount, tableIdMap) = buildTablesData(workbook)

    // Build unified style index with per-sheet remappings
    val (styleIndex, sheetRemappings) = StyleIndex.fromWorkbook(workbook)
    val styles = OoxmlStyles(styleIndex)

    // Convert domain workbook to OOXML
    val ooxmlWb = OoxmlWorkbook.fromDomain(workbook)

    // Convert sheets to OOXML worksheets with style remapping
    val ooxmlSheets = workbook.sheets.zipWithIndex.map { case (sheet, sheetIdx) =>
      val remapping = sheetRemappings.getOrElse(sheetIdx, Map.empty)

      // Generate tableParts XML element if sheet has tables
      val tablePartsXml = tablesBySheet.get(sheetIdx).flatMap { tablesForSheet =>
        if tablesForSheet.isEmpty then None
        else
          val hasComments = commentsBySheet.contains(sheetIdx)
          val rIdOffset = if hasComments then 3 else 1 // Comments use rId1-2

          val tablePartElems =
            tablesForSheet.sortBy(_._2).zipWithIndex.map { case ((_, tableId), idx) =>
              import scala.xml.*
              Elem(
                prefix = null,
                label = "tablePart",
                attributes = new PrefixedAttribute("r", "id", s"rId${rIdOffset + idx}", Null),
                scope = NamespaceBinding("r", XmlUtil.nsRelationships, TopScope),
                minimizeEmpty = true
              )
            }

          Some(
            XmlUtil.elem("tableParts", "count" -> tablesForSheet.size.toString)(tablePartElems*)
          )
      }

      val escapeFormulas = config.formulaInjectionPolicy == FormulaInjectionPolicy.Escape
      OoxmlWorksheet.fromDomainWithSST(sheet, sst, remapping, tablePartsXml, escapeFormulas)
    }

    // Create content types
    val contentTypes =
      ContentTypes
        .minimal(
          hasStyles = true, // Always include styles
          hasSharedStrings = sst.isDefined,
          sheetCount = workbook.sheets.size,
          sheetsWithComments = sheetsWithComments
        )
        .withCommentOverrides(sheetsWithComments)
        .withTableOverrides(totalTableCount)

    // Create relationships
    val rootRels = Relationships.root()
    val workbookRels = Relationships.workbook(
      sheetCount = workbook.sheets.size,
      hasStyles = true,
      hasSharedStrings = sst.isDefined
    )

    // Dispatch based on target type
    target match
      case OutputPath(path) =>
        writeZip(
          path,
          contentTypes,
          rootRels,
          workbookRels,
          ooxmlWb,
          ooxmlSheets,
          styles,
          sst,
          commentsBySheet,
          vmlDrawings,
          tablesBySheet,
          tableIdMap,
          config,
          Some(workbook),
          sheetRemappings
        )
      case OutputStreamTarget(stream) =>
        writeZipToStream(
          stream,
          contentTypes,
          rootRels,
          workbookRels,
          ooxmlWb,
          ooxmlSheets,
          styles,
          sst,
          commentsBySheet,
          vmlDrawings,
          tablesBySheet,
          tableIdMap,
          config,
          Some(workbook),
          sheetRemappings
        )

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
    commentsBySheet: Map[Int, OoxmlComments],
    vmlDrawings: Map[Int, String],
    tablesBySheet: Map[Int, Seq[(TableSpec, Long)]],
    tableIdMap: Map[String, Long],
    config: WriterConfig,
    domainWorkbook: Option[Workbook] = None,
    sheetRemappings: Map[Int, Map[Int, Int]] = Map.empty
  ): Unit =
    val zip = new ZipOutputStream(new FileOutputStream(path.toFile))
    zip.setLevel(1) // Match Excel's compression level (super fast)
    try
      writeZipContents(
        zip,
        contentTypes,
        rootRels,
        workbookRels,
        workbook,
        sheets,
        styles,
        sst,
        commentsBySheet,
        vmlDrawings,
        tablesBySheet,
        tableIdMap,
        config,
        domainWorkbook,
        sheetRemappings
      )
    finally zip.close()

  /** Write all parts to ZIP stream (for OutputStreamTarget) */
  private def writeZipToStream(
    stream: OutputStream,
    contentTypes: ContentTypes,
    rootRels: Relationships,
    workbookRels: Relationships,
    workbook: OoxmlWorkbook,
    sheets: Vector[OoxmlWorksheet],
    styles: OoxmlStyles,
    sst: Option[SharedStrings],
    commentsBySheet: Map[Int, OoxmlComments],
    vmlDrawings: Map[Int, String],
    tablesBySheet: Map[Int, Seq[(TableSpec, Long)]],
    tableIdMap: Map[String, Long],
    config: WriterConfig,
    domainWorkbook: Option[Workbook] = None,
    sheetRemappings: Map[Int, Map[Int, Int]] = Map.empty
  ): Unit =
    val zip = new ZipOutputStream(stream)
    zip.setLevel(1) // Match Excel's compression level (super fast)
    try
      writeZipContents(
        zip,
        contentTypes,
        rootRels,
        workbookRels,
        workbook,
        sheets,
        styles,
        sst,
        commentsBySheet,
        vmlDrawings,
        tablesBySheet,
        tableIdMap,
        config,
        domainWorkbook,
        sheetRemappings
      )
    finally zip.close()

  /** Common logic for writing all parts to a ZIP stream */
  private def writeZipContents(
    zip: ZipOutputStream,
    contentTypes: ContentTypes,
    rootRels: Relationships,
    workbookRels: Relationships,
    workbook: OoxmlWorkbook,
    sheets: Vector[OoxmlWorksheet],
    styles: OoxmlStyles,
    sst: Option[SharedStrings],
    commentsBySheet: Map[Int, OoxmlComments],
    vmlDrawings: Map[Int, String],
    tablesBySheet: Map[Int, Seq[(TableSpec, Long)]],
    tableIdMap: Map[String, Long],
    config: WriterConfig,
    domainWorkbook: Option[Workbook],
    sheetRemappings: Map[Int, Map[Int, Int]]
  ): Unit =
    // Write parts in canonical order
    writePart(zip, "[Content_Types].xml", contentTypes, config)
    writePart(zip, "_rels/.rels", rootRels, config)
    writePart(zip, "xl/workbook.xml", workbook, config)
    writePart(zip, "xl/_rels/workbook.xml.rels", workbookRels, config)

    // Write styles
    writeStyles(zip, "xl/styles.xml", styles, config)

    // Write shared strings if present
    sst.foreach { sharedStrings =>
      writeSharedStrings(zip, "xl/sharedStrings.xml", sharedStrings, config)
    }

    // Write worksheets
    sheets.zipWithIndex.foreach { case (sheet, idx) =>
      // For SaxStax backend with domain workbook available, use direct emission
      (config.backend, domainWorkbook) match
        case (XmlBackend.SaxStax, Some(dwb)) =>
          val domainSheet = dwb.sheets(idx)
          val remapping = sheetRemappings.getOrElse(idx, Map.empty)
          val escapeFormulas = config.formulaInjectionPolicy == FormulaInjectionPolicy.Escape
          writeWorksheetDirect(
            zip,
            s"xl/worksheets/sheet${idx + 1}.xml",
            domainSheet,
            sst,
            remapping,
            sheet.tableParts,
            escapeFormulas,
            config
          )
        case _ =>
          writeWorksheet(zip, s"xl/worksheets/sheet${idx + 1}.xml", sheet, config)

      // Write worksheet relationships if this sheet has comments or tables
      val hasComments = commentsBySheet.contains(idx)
      val tableIds = tablesBySheet.get(idx).map(_.map(_._2)).getOrElse(Seq.empty)

      if hasComments || tableIds.nonEmpty then
        val sheetRels = buildWorksheetRelationshipsWithTables(idx + 1, hasComments, tableIds)
        writePart(
          zip,
          s"xl/worksheets/_rels/sheet${idx + 1}.xml.rels",
          sheetRels,
          config
        )
    }

    // Write comment files and VML drawings for sheets with comments
    commentsBySheet.foreach { case (idx, comments) =>
      writePart(zip, s"xl/comments${idx + 1}.xml", comments, config)

      // Write VML drawing for comment indicators
      vmlDrawings.get(idx).foreach { vmlXml =>
        writeVmlPart(zip, s"xl/drawings/vmlDrawing${idx + 1}.vml", vmlXml, config)
      }
    }

    // Write table files for all sheets
    tablesBySheet.values.flatten.foreach { case (tableSpec, tableId) =>
      val ooxmlTable = TableConversions.toOoxml(tableSpec, tableId)
      writePart(zip, s"xl/tables/table$tableId.xml", ooxmlTable, config)
    }

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

    // When metadata is modified (add/remove/rename/reorder), regenerate ALL sheets.
    // New sheets don't exist in source ZIP, and removed sheets change indices.
    // This is the correct behavior - metadata changes are structural.
    if tracker.modifiedMetadata then
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

  /** Helper to open a ZipFile with automatic resource management. */
  private def withZipFile[A](path: Path)(f: ZipFile => A): A =
    usingOrThrow(Using.Manager { use =>
      val zip = use(new ZipFile(path.toFile))
      f(zip)
    })

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

    val sharedStringsPath = "xl/sharedStrings.xml"
    val sourceHasSharedStrings = sourceContext.exists(_.partManifest.contains(sharedStringsPath))

    // SST Strategy:
    // - If modified sheets have new strings not in preserved SST → regenerate (to avoid inline string corruption)
    // - If modified sheets only use existing SST strings → copy verbatim (byte-perfect preservation)
    // - If no source SST → generate according to policy
    val (sst, regenerateSharedStrings) =
      if sourceHasSharedStrings then
        // Parse preserved SST (sourceContext guaranteed to exist if sourceHasSharedStrings is true)
        val parsedSST = sourceContext.map(ctx => parsePreservedSST(ctx.sourcePath)).getOrElse(None)

        // Check if modified sheets contain NEW strings not in preserved SST
        // Extract entries from modified sheets (normalize for comparison consistency)
        val modifiedSheetEntries: Set[Either[String, RichText]] =
          tracker.modifiedSheets.flatMap { idx =>
            workbook.sheets.lift(idx).toList.flatMap { sheet =>
              sheet.cells.values.flatMap { cell =>
                cell.value match
                  case CellValue.Text(str) => Some(Left(SharedStrings.normalize(str)))
                  case CellValue.RichText(rt) => Some(Right(rt))
                  case _ => None
              }
            }
          }.toSet

        // Get preserved SST entries (normalize for comparison)
        val preservedEntries: Set[Either[String, RichText]] =
          parsedSST
            .map(
              _.strings
                .map {
                  case Left(s) => Left(SharedStrings.normalize(s))
                  case Right(rt) => Right(rt)
                }
                .toSet
            )
            .getOrElse(Set.empty)

        // Determine if modified sheets introduced new strings (normalized comparison)
        val newEntries = modifiedSheetEntries.diff(preservedEntries)

        if newEntries.nonEmpty then
          // New strings detected → build SST from preserved + new entries only
          // Do NOT use SharedStrings.fromWorkbook (includes ALL sheets, even binary system sheets!)
          val combinedEntries =
            parsedSST.map(_.strings).getOrElse(Vector.empty) ++ newEntries.toVector

          // Calculate total reference count: preserved count + new string references
          val originalTotalCount = parsedSST.map(_.totalCount).getOrElse(0)
          val newStringRefCount = newEntries.size // Each new entry used at least once
          val combinedTotalCount = originalTotalCount + newStringRefCount

          // Create new SST with combined entries
          val combinedSST = SharedStrings(
            strings = combinedEntries,
            indexMap = combinedEntries.zipWithIndex.map {
              case (Left(s), idx) => SharedStrings.normalize(s) -> idx
              case (Right(rt), idx) => SharedStrings.normalize(rt.toPlainText) -> idx
            }.toMap,
            totalCount = combinedTotalCount
          )
          (Some(combinedSST), true)
        else
          // No new strings → copy preserved SST verbatim for byte-perfect preservation
          (parsedSST, false)
      else
        // No source SST - generate if policy allows
        val generated = config.sstPolicy match
          case SstPolicy.Always => Some(SharedStrings.fromWorkbook(workbook))
          case SstPolicy.Never => None
          case SstPolicy.Auto =>
            if SharedStrings.shouldUseSST(workbook) then Some(SharedStrings.fromWorkbook(workbook))
            else None
        (generated, generated.isDefined)

    val sharedStringsInOutput = sourceHasSharedStrings || regenerateSharedStrings
    val sstForSheets = if regenerateSharedStrings then sst else None

    // Build style index (automatic optimization based on sourceContext)
    val (styleIndex, sheetRemappings) = StyleIndex.fromWorkbook(workbook)

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

    // Preserve structural parts from source (or fallback to minimal)
    // IMPORTANT: When metadata is modified (add/remove/rename/reorder sheets),
    // we MUST regenerate the structural parts since sheet count/order changed.
    val (preservedContentTypes, preservedRootRels, preservedWorkbookRels, preservedWorkbook) =
      sourceContext match
        case Some(ctx) if !tracker.modifiedMetadata => parsePreservedStructure(ctx.sourcePath)
        case _ => (None, None, None, None)

    // Use preserved workbook structure if available, otherwise create minimal
    val ooxmlWb = preservedWorkbook match
      case Some(preserved) =>
        // Update sheets in preserved structure (names/order/visibility may have changed)
        preserved.updateSheets(workbook.sheets, workbook.metadata.sheetStates)
      case None =>
        // Fallback to minimal for programmatically created workbooks
        // (or when metadata modified - we need fresh workbook structure)
        OoxmlWorkbook.fromDomain(workbook)

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
        // Only add table overrides (comments already in preserved Content_Types)
        withSst.withTableOverrides(totalTableCount)
      case None =>
        ContentTypes
          .minimal(
            hasStyles = true,
            hasSharedStrings = sharedStringsInOutput,
            sheetCount = workbook.sheets.size,
            sheetsWithComments = sheetsWithComments
          )
          .withCommentOverrides(sheetsWithComments)
          .withTableOverrides(totalTableCount)

    val rootRels = preservedRootRels.getOrElse(Relationships.root())

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

      // When metadata is modified (add/remove/rename/reorder), we must regenerate ALL sheets
      // because new sheets don't exist in source and sheet indices may have changed.
      val sheetsToRegenerate =
        if tracker.modifiedMetadata then workbook.sheets.indices.toSet
        else tracker.modifiedSheets

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

          val escapeFormulas = config.formulaInjectionPolicy == FormulaInjectionPolicy.Escape

          // For SaxStax backend with no preserved metadata, use direct SAX emission
          // (bypasses intermediate OOXML types for 5-7x performance improvement)
          (config.backend, preservedMetadata) match
            case (XmlBackend.SaxStax, None) =>
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
                  escapeFormulas
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

          // Copy relationships from source if available (preserves non-comment relationships)
          // Otherwise regenerate minimal relationships
          sourceContext match
            case Some(ctx) if ctx.partManifest.contains(relsPath) =>
              copyPreservedPart(ctx.sourcePath, relsPath, zip)
            case _ if hasComments || tableIds.nonEmpty =>
              val sheetRels =
                buildWorksheetRelationshipsWithCommentsPath(commentPath, hasComments, tableIds)
              writePart(zip, relsPath, sheetRels, config)
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

      // Copy preserved parts (charts, drawings, images, etc.) if source available
      // Skip VML drawings for regenerated sheets - we regenerate VML from domain model
      // (or omit it if comments were removed)
      sourceContext.foreach { ctx =>
        val vmlPathsToSkip = sheetsToRegenerate.flatMap { idx =>
          ctx.commentPathMapping.get(idx).map { commentPath =>
            val fileNum = commentPath.stripPrefix("xl/comments").stripSuffix(".xml")
            s"xl/drawings/vmlDrawing$fileNum.vml"
          }
        }

        preservableParts.foreach { path =>
          val isVmlDrawing = path.startsWith("xl/drawings/vmlDrawing") && path.endsWith(".vml")
          val shouldSkip = isVmlDrawing && vmlPathsToSkip.contains(path)

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
