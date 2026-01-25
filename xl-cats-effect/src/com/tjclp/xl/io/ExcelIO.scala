package com.tjclp.xl.io

import cats.effect.{Async, Sync, Resource}
import cats.syntax.all.*
import fs2.{Stream, Pipe}
import java.nio.file.{Files as JFiles, Path}
import java.io.{BufferedOutputStream, FileOutputStream, FileInputStream}
import java.util.zip.{ZipOutputStream, ZipEntry, CRC32, ZipInputStream, ZipFile}
import java.nio.charset.StandardCharsets
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.api.Workbook
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.ooxml.{XlsxReader, XlsxWriter, SharedStrings, XmlSecurity}
import com.tjclp.xl.ooxml.metadata.{LightMetadata, WorkbookMetadataReader}
import fs2.data.xml
import fs2.data.xml.XmlEvent

/**
 * Cats Effect interpreter for Excel operations.
 *
 * Implements both Excel[F] (error-raising) and ExcelR[F] (explicit error channel). Wraps pure
 * XlsxReader/Writer with effect type F[_].
 */
class ExcelIO[F[_]: Async](warningHandler: XlsxReader.Warning => F[Unit])
    extends Excel[F]
    with ExcelR[F]:

  /** Read workbook from XLSX file */
  def read(path: Path): F[Workbook] =
    readWith(path, XlsxReader.ReaderConfig.default)

  /** Read workbook from XLSX file with custom configuration */
  def readWith(path: Path, config: XlsxReader.ReaderConfig): F[Workbook] =
    Sync[F].delay(XlsxReader.readWithWarnings(path, config)).flatMap {
      case Right(result) =>
        result.warnings.traverse_(warningHandler) *> Async[F].pure(result.workbook)
      case Left(err) => Async[F].raiseError(new Exception(s"Failed to read XLSX: ${err.message}"))
    }

  /** Write workbook to XLSX file */
  def write(wb: Workbook, path: Path): F[Unit] =
    writeWith(wb, path, com.tjclp.xl.ooxml.WriterConfig.default)

  /** Write workbook to XLSX file with custom configuration */
  def writeWith(wb: Workbook, path: Path, config: com.tjclp.xl.ooxml.WriterConfig): F[Unit] =
    Sync[F].delay(XlsxWriter.writeWith(wb, path, config)).flatMap {
      case Right(_) => Async[F].unit
      case Left(err) => Async[F].raiseError(new Exception(s"Failed to write XLSX: ${err.message}"))
    }

  /** Write workbook using SAX/StAX backend for faster writes */
  @deprecated("SaxStax is now the default backend. Use write() instead.", "0.4.0")
  def writeFast(wb: Workbook, path: Path): F[Unit] =
    write(wb, path)

  /** Write workbook with custom config but forcing SAX/StAX backend */
  @deprecated("SaxStax is now the default backend. Use writeWith() instead.", "0.4.0")
  def writeFastWith(
    wb: Workbook,
    path: Path,
    config: com.tjclp.xl.ooxml.WriterConfig
  ): F[Unit] =
    writeWith(wb, path, config)

  /**
   * Stream rows from first sheet with constant memory.
   *
   * Uses SAX parser (javax.xml.parsers.SAXParser) for 3-4x speedup vs fs2-data-xml.
   */
  def readStream(path: Path): Stream[F, RowData] =
    Stream
      .bracket(
        Sync[F].delay(new ZipFile(path.toFile))
      )(zipFile => Sync[F].delay(zipFile.close()))
      .flatMap { zipFile =>
        Stream
          .eval {
            // Parse SST if present
            val sstEntry = Option(zipFile.getEntry("xl/sharedStrings.xml"))
            sstEntry match
              case Some(entry) =>
                val sstBytes = fs2.io.readInputStream[F](
                  Sync[F].delay(zipFile.getInputStream(entry)),
                  chunkSize = 4096
                )
                StreamingXmlReader.parseSharedStrings[F](sstBytes)
              case None =>
                Sync[F].pure(None)
          }
          .flatMap { sst =>
            // Stream first worksheet using SAX parser
            val wsEntry = Option(zipFile.getEntry("xl/worksheets/sheet1.xml"))
            wsEntry match
              case Some(entry) =>
                Stream
                  .bracket(Sync[F].delay(zipFile.getInputStream(entry)))(stream =>
                    Sync[F].delay(stream.close())
                  )
                  .flatMap { stream =>
                    SaxStreamingReader.parseWorksheetStream[F](stream, sst)
                  }
              case None =>
                Stream.empty
          }
      }

  /** Stream rows from first sheet within a bounded range (rows/cols). */
  def readStreamRange(path: Path, range: CellRange): Stream[F, RowData] =
    Stream
      .bracket(
        Sync[F].delay(new ZipFile(path.toFile))
      )(zipFile => Sync[F].delay(zipFile.close()))
      .flatMap { zipFile =>
        readStreamByIndex(zipFile, 1, Some(range))
      }

  /**
   * Stream rows from sheet by name with constant memory.
   *
   * Uses fs2-data-xml pull parsing for the target sheet.
   */
  def readSheetStream(path: Path, sheetName: String): Stream[F, RowData] =
    Stream
      .bracket(
        Sync[F].delay(new ZipFile(path.toFile))
      )(zipFile => Sync[F].delay(zipFile.close()))
      .flatMap { zipFile =>
        Stream
          .eval {
            // Find sheet index by parsing workbook.xml
            findSheetIndexByName(zipFile, sheetName)
          }
          .flatMap {
            case Some(sheetIndex) =>
              readStreamByIndex(zipFile, sheetIndex)
            case None =>
              Stream.raiseError[F](new Exception(s"Sheet not found: $sheetName"))
          }
      }

  /** Stream rows from sheet by name within a bounded range (rows/cols). */
  def readSheetStreamRange(path: Path, sheetName: String, range: CellRange): Stream[F, RowData] =
    Stream
      .bracket(
        Sync[F].delay(new ZipFile(path.toFile))
      )(zipFile => Sync[F].delay(zipFile.close()))
      .flatMap { zipFile =>
        Stream
          .eval {
            // Find sheet index by parsing workbook.xml
            findSheetIndexByName(zipFile, sheetName)
          }
          .flatMap {
            case Some(sheetIndex) =>
              readStreamByIndex(zipFile, sheetIndex, Some(range))
            case None =>
              Stream.raiseError[F](new Exception(s"Sheet not found: $sheetName"))
          }
      }

  /** Stream rows from sheet by index (1-based) with constant memory. */
  def readStreamByIndex(path: Path, sheetIndex: Int): Stream[F, RowData] =
    Stream
      .bracket(
        Sync[F].delay(new ZipFile(path.toFile))
      )(zipFile => Sync[F].delay(zipFile.close()))
      .flatMap { zipFile =>
        readStreamByIndex(zipFile, sheetIndex)
      }

  /** Stream rows from sheet by index within a bounded range (rows/cols). */
  def readStreamByIndexRange(
    path: Path,
    sheetIndex: Int,
    range: CellRange
  ): Stream[F, RowData] =
    Stream
      .bracket(
        Sync[F].delay(new ZipFile(path.toFile))
      )(zipFile => Sync[F].delay(zipFile.close()))
      .flatMap { zipFile =>
        readStreamByIndex(zipFile, sheetIndex, Some(range))
      }

  // ===== Lightweight Metadata Operations =====

  /**
   * Read workbook metadata only (no cell data). Instant for any file size.
   *
   * Uses pure WorkbookMetadataReader from xl-ooxml module.
   */
  def readMetadata(path: Path): F[LightMetadata] =
    Sync[F].delay(WorkbookMetadataReader.read(path)).flatMap {
      case Right(meta) => Async[F].pure(meta)
      case Left(err) =>
        Async[F].raiseError(new Exception(s"Failed to read metadata: ${err.message}"))
    }

  /**
   * Read dimension from specific worksheet (1-based index). Instant for any file size.
   */
  def readDimension(path: Path, sheetIndex: Int): F[Option[CellRange]] =
    Sync[F].delay(WorkbookMetadataReader.readDimension(path, sheetIndex)).flatMap {
      case Right(dim) => Async[F].pure(dim)
      case Left(err) =>
        Async[F].raiseError(new Exception(s"Failed to read dimension: ${err.message}"))
    }

  // Helper: Stream rows from specific sheet index using open ZipFile
  private def readStreamByIndex(zipFile: ZipFile, sheetIndex: Int): Stream[F, RowData] =
    readStreamByIndex(zipFile, sheetIndex, None)

  private def readStreamByIndex(
    zipFile: ZipFile,
    sheetIndex: Int,
    range: Option[CellRange]
  ): Stream[F, RowData] =
    Stream
      .eval {
        // Parse SST if present
        val sstEntry = Option(zipFile.getEntry("xl/sharedStrings.xml"))
        sstEntry match
          case Some(entry) =>
            val sstBytes = fs2.io.readInputStream[F](
              Sync[F].delay(zipFile.getInputStream(entry)),
              chunkSize = 4096
            )
            StreamingXmlReader.parseSharedStrings[F](sstBytes)
          case None =>
            Sync[F].pure(None)
      }
      .flatMap { sst =>
        // Stream specified worksheet
        val wsEntry = Option(zipFile.getEntry(s"xl/worksheets/sheet$sheetIndex.xml"))
        wsEntry match
          case Some(entry) =>
            Stream
              .bracket(Sync[F].delay(zipFile.getInputStream(entry)))(stream =>
                Sync[F].delay(stream.close())
              )
              .flatMap { stream =>
                val rowBounds = range.map(r => (r.start.row.index1, r.end.row.index1))
                val colBounds = range.map(r => (r.start.col.index0, r.end.col.index0))
                SaxStreamingReader.parseWorksheetStream[F](stream, sst, rowBounds, colBounds)
              }
          case None =>
            Stream.raiseError[F](
              new Exception(s"Worksheet not found: sheet$sheetIndex.xml")
            )
      }

  /**
   * Streaming write with constant O(1) memory.
   *
   * Single-pass streaming directly to ZIP. If `dimension` hint is provided, it's written to enable
   * instant metadata queries. Otherwise, dimension is omitted (use `writeStreamWithAutoDetect` for
   * automatic dimension detection at the cost of 2x I/O).
   *
   * @param path
   *   Output file path
   * @param sheetName
   *   Name of the worksheet
   * @param sheetIndex
   *   Sheet index (1-based, default 1)
   * @param config
   *   Writer configuration
   * @param dimension
   *   Optional dimension hint. When provided, enables instant bounds queries. Use
   *   `CellRange(ARef.from0(0, 0), ARef.from0(maxCol, maxRow))` to specify bounds.
   */
  def writeStream(
    path: Path,
    sheetName: String,
    sheetIndex: Int = 1,
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default,
    dimension: Option[CellRange] = None
  ): Pipe[F, RowData, Unit] =
    rows =>
      // Validate sheet index
      if sheetIndex < 1 then
        Stream.raiseError[F](
          new IllegalArgumentException(s"Sheet index must be >= 1, got: $sheetIndex")
        )
      else
        Stream
          .bracket(
            Sync[F].delay(new ZipOutputStream(new FileOutputStream(path.toFile)))
          )(zip => Sync[F].delay(zip.close()))
          .flatMap { zip =>
            // 1. Write static parts first
            Stream.eval(Sync[F].delay(writeStaticPartsSync(zip, sheetName, sheetIndex, config))) ++
              // 2. Open worksheet entry
              Stream.eval(
                Sync[F].delay {
                  val entry = new ZipEntry(s"xl/worksheets/sheet$sheetIndex.xml")
                  entry.setMethod(ZipEntry.DEFLATED)
                  zip.putNextEntry(entry)
                  // Write header with optional dimension
                  writeWorksheetHeader(zip, dimension)
                }
              ) ++
              // 3. Stream XML events → bytes → ZIP
              StreamingXmlWriter
                .worksheetBody(rows, config.formulaInjectionPolicy)
                .through(xml.render.raw())
                .through(fs2.text.utf8.encode)
                .chunks
                .evalMap(chunk => Sync[F].delay(zip.write(chunk.toArray))) ++
              // 4. Write footer and close entry
              Stream.eval(
                Sync[F].delay {
                  writeWorksheetFooter(zip)
                  zip.closeEntry()
                }
              )
          }
          .drain

  /**
   * Streaming write with automatic dimension detection.
   *
   * Uses two-phase approach: streams to temp file while tracking bounds, then assembles final ZIP
   * with accurate dimension element. This enables instant metadata queries but costs 2x I/O.
   *
   * For better performance when bounds are known, use `writeStream` with explicit `dimension`.
   */
  def writeStreamWithAutoDetect(
    path: Path,
    sheetName: String,
    sheetIndex: Int = 1,
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): Pipe[F, RowData, Unit] =
    rows =>
      if sheetIndex < 1 then
        Stream.raiseError[F](
          new IllegalArgumentException(s"Sheet index must be >= 1, got: $sheetIndex")
        )
      else
        Stream
          .bracket(
            Sync[F].delay {
              val tempFile = JFiles.createTempFile("xl-stream-", ".xml")
              val bounds = new BoundsAccumulator()
              (tempFile, bounds)
            }
          ) { case (tempFile, _) =>
            Sync[F].delay(JFiles.deleteIfExists(tempFile)).void
          }
          .flatMap { case (tempFile, bounds) =>
            // Phase 1: Stream rows to temp file, track bounds
            val writeTemp = rows
              .evalTap(row => Sync[F].delay(bounds.update(row)))
              .through(rowsToTempXml(tempFile, config))

            // Phase 2: Assemble final ZIP with dimension
            val assembleZip = Stream.eval(
              assembleWorksheetZip(path, tempFile, bounds, sheetName, sheetIndex, config)
            )

            writeTemp ++ assembleZip
          }
          .drain

  // Helper: Stream rows to temp XML file (body only, no header/footer)
  private def rowsToTempXml(
    tempFile: Path,
    config: com.tjclp.xl.ooxml.WriterConfig
  ): Pipe[F, RowData, Unit] =
    rows =>
      Stream
        .bracket(
          Sync[F].delay(new BufferedOutputStream(new FileOutputStream(tempFile.toFile)))
        )(os => Sync[F].delay(os.close()))
        .flatMap { os =>
          StreamingXmlWriter
            .worksheetBody(rows, config.formulaInjectionPolicy)
            .through(xml.render.raw())
            .through(fs2.text.utf8.encode)
            .chunks
            .evalMap(chunk => Sync[F].delay(os.write(chunk.toArray)))
        }

  // Helper: Assemble final ZIP with dimension element
  private def assembleWorksheetZip(
    path: Path,
    tempBodyFile: Path,
    bounds: BoundsAccumulator,
    sheetName: String,
    sheetIndex: Int,
    config: com.tjclp.xl.ooxml.WriterConfig
  ): F[Unit] =
    Sync[F].delay {
      val zip = new ZipOutputStream(new FileOutputStream(path.toFile))
      try
        // 1. Write static parts
        writeStaticPartsSync(zip, sheetName, sheetIndex, config)

        // 2. Write worksheet with dimension
        val wsEntry = new ZipEntry(s"xl/worksheets/sheet$sheetIndex.xml")
        wsEntry.setMethod(ZipEntry.DEFLATED)
        zip.putNextEntry(wsEntry)

        // Header with dimension (from tracked bounds)
        writeWorksheetHeader(zip, bounds.dimension)

        // Body from temp file
        JFiles.copy(tempBodyFile, zip)

        // Footer
        writeWorksheetFooter(zip)

        zip.closeEntry()
      finally zip.close()
    }

  // Helper: Write worksheet header XML to output stream
  private def writeWorksheetHeader(
    out: java.io.OutputStream,
    dimension: Option[CellRange]
  ): Unit =
    val sb = new StringBuilder
    sb.append("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
    sb.append("""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" """)
    sb.append("""xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">""")
    dimension.foreach { range =>
      sb.append(s"""<dimension ref="${range.toA1}"/>""")
    }
    sb.append("<sheetData>")
    out.write(sb.toString.getBytes(StandardCharsets.UTF_8))

  // Helper: Write worksheet footer XML to output stream
  private def writeWorksheetFooter(out: java.io.OutputStream): Unit =
    out.write("</sheetData></worksheet>".getBytes(StandardCharsets.UTF_8))

  // Helper: Sync version of writeStaticParts for use in assembleWorksheetZip
  private def writeStaticPartsSync(
    zip: ZipOutputStream,
    sheetName: String,
    sheetIndex: Int,
    config: com.tjclp.xl.ooxml.WriterConfig
  ): Unit =
    import com.tjclp.xl.ooxml.*
    import com.tjclp.xl.addressing.SheetName

    val contentTypes =
      ContentTypes.forSheetIndices(Seq(sheetIndex), hasStyles = true, hasSharedStrings = false)
    val rootRels = Relationships.root()
    val workbookRels =
      Relationships.workbookWithIndices(Seq(sheetIndex), hasStyles = true, hasSharedStrings = false)

    val ooxmlWb = OoxmlWorkbook(
      sheets = Vector(SheetRef(SheetName.unsafe(sheetName), sheetIndex, "rId1"))
    )

    val styles = OoxmlStyles.minimal

    writePartSync(zip, "[Content_Types].xml", contentTypes.toXml, config)
    writePartSync(zip, "_rels/.rels", rootRels.toXml, config)
    writePartSync(zip, "xl/workbook.xml", ooxmlWb.toXml, config)
    writePartSync(zip, "xl/_rels/workbook.xml.rels", workbookRels.toXml, config)
    writePartSync(zip, "xl/styles.xml", styles.toXml, config)

  // Helper: Write a single XML part to ZIP (sync version)
  private def writePartSync(
    zip: ZipOutputStream,
    entryName: String,
    xml: scala.xml.Elem,
    config: com.tjclp.xl.ooxml.WriterConfig
  ): Unit =
    import com.tjclp.xl.ooxml.{XmlUtil, Compression}

    val xmlString = if config.prettyPrint then XmlUtil.prettyPrint(xml) else XmlUtil.compact(xml)
    val bytes = xmlString.getBytes(StandardCharsets.UTF_8)

    val entry = new ZipEntry(entryName)
    entry.setMethod(config.compression.zipMethod)

    config.compression match
      case Compression.Stored =>
        entry.setSize(bytes.length)
        entry.setCompressedSize(bytes.length)
        val crc = new CRC32()
        crc.update(bytes)
        entry.setCrc(crc.getValue)
      case Compression.Deflated =>
        ()

    zip.putNextEntry(entry)
    zip.write(bytes)
    zip.closeEntry()

  /**
   * Write rows to XLSX file, materializing all rows first.
   *
   * @deprecated
   *   Use writeStream instead for true O(1) streaming.
   */
  @deprecated("Use writeStream for true O(1) streaming", "0.8.0")
  def writeStreamMaterialized(path: Path, sheetName: String): Pipe[F, RowData, Unit] =
    rows =>
      Stream.eval {
        rows.compile.toVector.flatMap { rowVec =>
          val sheet = rowDataToSheet(sheetName, rowVec)
          sheet match
            case Right(s) =>
              val wb = Workbook(Vector(s))
              write(wb, path)
            case Left(err) =>
              Async[F].raiseError(new Exception(s"Failed to create sheet: ${err.message}"))
        }
      }

  /**
   * Write multiple sheets sequentially with constant memory.
   *
   * Single-pass streaming directly to ZIP. Dimensions can be provided per-sheet for instant
   * metadata queries.
   *
   * @param path
   *   Output file path
   * @param sheets
   *   Sequence of (name, rows, optional dimension) tuples
   * @param config
   *   Writer configuration
   */
  def writeStreamsSeq(
    path: Path,
    sheets: Seq[(String, Stream[F, RowData])],
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): F[Unit] =
    writeStreamsSeqWithDimensions(
      path,
      sheets.map { case (name, rows) => (name, rows, None) },
      config
    )

  /**
   * Write multiple sheets with explicit dimension hints.
   *
   * Single-pass streaming directly to ZIP. When dimensions are provided, enables instant bounds
   * queries.
   */
  def writeStreamsSeqWithDimensions(
    path: Path,
    sheets: Seq[(String, Stream[F, RowData], Option[CellRange])],
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): F[Unit] =
    // Validate inputs
    if sheets.isEmpty then
      Async[F].raiseError(new IllegalArgumentException("Must provide at least one sheet"))
    else if sheets.map(_._1).distinct.size != sheets.size then
      Async[F].raiseError(
        new IllegalArgumentException(s"Duplicate sheet names: ${sheets.map(_._1).mkString(", ")}")
      )
    else
      // Auto-assign sheet indices 1, 2, 3...
      val sheetsWithIndices = sheets.zipWithIndex.map { case ((name, rows, dim), idx) =>
        (name, idx + 1, rows, dim)
      }

      Stream
        .bracket(
          Sync[F].delay(new ZipOutputStream(new FileOutputStream(path.toFile)))
        )(zip => Sync[F].delay(zip.close()))
        .flatMap { zip =>
          // 1. Write static parts with all sheet metadata
          Stream.eval(
            Sync[F].delay {
              writeStaticPartsMultiSync(
                zip,
                sheetsWithIndices.map { case (name, idx, _, _) => (name, idx) },
                config
              )
            }
          ) ++
            // 2. Stream each sheet sequentially
            Stream
              .emits(sheetsWithIndices)
              .flatMap { case (_, sheetIndex, rows, dimension) =>
                // Open worksheet entry
                Stream.eval(
                  Sync[F].delay {
                    val entry = new ZipEntry(s"xl/worksheets/sheet$sheetIndex.xml")
                    entry.setMethod(ZipEntry.DEFLATED)
                    zip.putNextEntry(entry)
                    writeWorksheetHeader(zip, dimension)
                  }
                ) ++
                  // Stream XML events → bytes → ZIP
                  StreamingXmlWriter
                    .worksheetBody(rows, config.formulaInjectionPolicy)
                    .through(xml.render.raw())
                    .through(fs2.text.utf8.encode)
                    .chunks
                    .evalMap(chunk => Sync[F].delay(zip.write(chunk.toArray))) ++
                  // Close entry
                  Stream.eval(
                    Sync[F].delay {
                      writeWorksheetFooter(zip)
                      zip.closeEntry()
                    }
                  )
              }
        }
        .compile
        .drain

  /**
   * Write multiple sheets with automatic dimension detection.
   *
   * Uses two-phase approach for each sheet. For better performance when bounds are known, use
   * `writeStreamsSeqWithDimensions`.
   */
  def writeStreamsSeqWithAutoDetect(
    path: Path,
    sheets: Seq[(String, Stream[F, RowData])],
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): F[Unit] =
    // Validate inputs
    if sheets.isEmpty then
      Async[F].raiseError(new IllegalArgumentException("Must provide at least one sheet"))
    else if sheets.map(_._1).distinct.size != sheets.size then
      Async[F].raiseError(
        new IllegalArgumentException(s"Duplicate sheet names: ${sheets.map(_._1).mkString(", ")}")
      )
    else
      // Auto-assign sheet indices 1, 2, 3...
      val sheetsWithIndices = sheets.zipWithIndex.map { case ((name, rows), idx) =>
        (name, idx + 1, rows)
      }

      // Create temp file + bounds tracker for each sheet
      Stream
        .bracket(
          Sync[F].delay {
            sheetsWithIndices.map { case (name, sheetIndex, _) =>
              val tempFile = JFiles.createTempFile(s"xl-stream-$sheetIndex-", ".xml")
              val bounds = new BoundsAccumulator()
              (name, sheetIndex, tempFile, bounds)
            }
          }
        ) { resources =>
          // Cleanup: delete all temp files
          Sync[F].delay {
            resources.foreach { case (_, _, tempFile, _) =>
              JFiles.deleteIfExists(tempFile)
            }
          }.void
        }
        .flatMap { resources =>
          // Phase 1: Stream each sheet's rows to its temp file
          val writePhase = Stream
            .emits(sheetsWithIndices.zip(resources))
            .flatMap { case ((_, _, rows), (_, _, tempFile, bounds)) =>
              rows
                .evalTap(row => Sync[F].delay(bounds.update(row)))
                .through(rowsToTempXml(tempFile, config))
            }

          // Phase 2: Assemble final ZIP with all worksheets
          val assemblePhase = Stream.eval(
            assembleMultiSheetZip(path, resources, config)
          )

          writePhase ++ assemblePhase
        }
        .compile
        .drain

  /**
   * Write workbook using streaming writer with O(1) output memory and style preservation.
   *
   * Hybrid approach: workbook is loaded in-memory (for style extraction), but output uses streaming
   * writer for O(1) output memory. Styles are fully preserved via the two-pass approach with
   * StyleIndex.
   *
   * Best for: Large modified workbooks where output memory is the bottleneck. Trade-off: Extra I/O
   * pass (temp file) for dimension detection.
   *
   * @param wb
   *   Workbook to write (in-memory)
   * @param path
   *   Output file path
   * @param config
   *   Writer configuration
   */
  def writeWorkbookStream(
    wb: Workbook,
    path: Path,
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): F[Unit] =
    import com.tjclp.xl.ooxml.style.{StyleIndex, OoxmlStyles}

    if wb.sheets.isEmpty then
      Async[F].raiseError(new IllegalArgumentException("Workbook must have at least one sheet"))
    else
      Sync[F]
        .delay {
          // Build unified style index from workbook (extracts all styles)
          val (styleIndex, remappings) = StyleIndex.fromWorkbook(wb)
          val ooxmlStyles = OoxmlStyles(styleIndex)

          // Prepare sheet data with remapped style IDs
          val sheetsWithIndices = wb.sheets.zipWithIndex.map { case (sheet, idx) =>
            val sheetIndex = idx + 1
            val remapping = remappings.getOrElse(idx, Map.empty[Int, Int])
            (sheet, sheetIndex, remapping)
          }

          (ooxmlStyles, sheetsWithIndices)
        }
        .flatMap { case (ooxmlStyles, sheetsWithIndices) =>
          // Create temp files for two-phase approach
          Stream
            .bracket(
              Sync[F].delay {
                sheetsWithIndices.map { case (sheet, sheetIndex, _) =>
                  val tempFile = JFiles.createTempFile(s"xl-wbstream-$sheetIndex-", ".xml")
                  val bounds = new BoundsAccumulator()
                  (sheet, sheetIndex, tempFile, bounds)
                }
              }
            ) { resources =>
              // Cleanup: delete all temp files
              Sync[F].delay {
                resources.foreach { case (_, _, tempFile, _) =>
                  JFiles.deleteIfExists(tempFile)
                }
              }.void
            }
            .flatMap { resources =>
              // Phase 1: Stream each sheet to temp file with style remapping
              val writePhase = Stream
                .emits(resources.zip(sheetsWithIndices))
                .flatMap { case ((_, _, tempFile, bounds), (sheet, _, remapping)) =>
                  streamSheetStyled(sheet, remapping)
                    .evalTap(row => Sync[F].delay(bounds.update(row.toRowData)))
                    .through(styledRowsToTempXml(tempFile, config))
                }

              // Phase 2: Assemble final ZIP with styles and dimensions
              val assemblePhase = Stream.eval(
                assembleWorkbookStreamZip(path, resources, ooxmlStyles, config)
              )

              writePhase ++ assemblePhase
            }
            .compile
            .drain
        }

  // Helper: Stream styled rows to temp XML file (body only)
  private def styledRowsToTempXml(
    tempFile: Path,
    config: com.tjclp.xl.ooxml.WriterConfig
  ): Pipe[F, StyledRowData, Unit] =
    rows =>
      Stream
        .bracket(
          Sync[F].delay(new BufferedOutputStream(new FileOutputStream(tempFile.toFile)))
        )(os => Sync[F].delay(os.close()))
        .flatMap { os =>
          StreamingXmlWriter
            .worksheetBodyStyled(rows, config.formulaInjectionPolicy)
            .through(xml.render.raw())
            .through(fs2.text.utf8.encode)
            .chunks
            .evalMap(chunk => Sync[F].delay(os.write(chunk.toArray)))
        }

  // Helper: Assemble workbook ZIP with styles
  private def assembleWorkbookStreamZip(
    path: Path,
    resources: Seq[(Sheet, Int, Path, BoundsAccumulator)],
    ooxmlStyles: com.tjclp.xl.ooxml.style.OoxmlStyles,
    config: com.tjclp.xl.ooxml.WriterConfig
  ): F[Unit] =
    Sync[F].delay {
      import com.tjclp.xl.ooxml.*
      import com.tjclp.xl.addressing.SheetName

      val zip = new ZipOutputStream(new FileOutputStream(path.toFile))
      try
        val sheets = resources.map { case (sheet, idx, _, _) => (sheet.name.value, idx) }

        // Content types and relationships
        val contentTypes = ContentTypes.forSheetIndices(
          sheetIndices = sheets.map(_._2),
          hasStyles = true,
          hasSharedStrings = false
        )
        val rootRels = Relationships.root()
        val workbookRels = Relationships.workbook(
          sheetCount = sheets.size,
          hasStyles = true,
          hasSharedStrings = false
        )

        // Workbook with sheet refs
        val sheetRefs = sheets.map { case (name, idx) =>
          SheetRef(SheetName.unsafe(name), idx, s"rId$idx")
        }
        val ooxmlWb = OoxmlWorkbook(sheets = sheetRefs.toVector)

        // Write static parts (with full styles, not minimal!)
        writePartSync(zip, "[Content_Types].xml", contentTypes.toXml, config)
        writePartSync(zip, "_rels/.rels", rootRels.toXml, config)
        writePartSync(zip, "xl/workbook.xml", ooxmlWb.toXml, config)
        writePartSync(zip, "xl/_rels/workbook.xml.rels", workbookRels.toXml, config)
        writePartSync(zip, "xl/styles.xml", ooxmlStyles.toXml, config) // Full styles!

        // Write each worksheet with dimension
        resources.foreach { case (_, sheetIndex, tempBodyFile, bounds) =>
          val wsEntry = new ZipEntry(s"xl/worksheets/sheet$sheetIndex.xml")
          wsEntry.setMethod(ZipEntry.DEFLATED)
          zip.putNextEntry(wsEntry)

          // Header with dimension
          writeWorksheetHeader(zip, bounds.dimension)

          // Body from temp file
          JFiles.copy(tempBodyFile, zip)

          // Footer
          writeWorksheetFooter(zip)

          zip.closeEntry()
        }
      finally zip.close()
    }

  // Helper: Convert Sheet to StyledRowData stream with remapped style IDs
  private def streamSheetStyled(
    sheet: Sheet,
    remapping: Map[Int, Int]
  ): Stream[F, StyledRowData] =
    val rowMap = sheet.cells.values
      .groupBy(_.ref.row.index1) // Group by 1-based row
      .view
      .mapValues { cells =>
        val cellValues = cells.map(c => c.ref.col.index0 -> c.value).toMap
        val cellStyles = cells.flatMap { c =>
          c.styleId.map { sid =>
            // Remap local styleId to global index
            val globalId = remapping.getOrElse(sid.value, sid.value)
            c.ref.col.index0 -> globalId
          }
        }.toMap
        (cellValues, cellStyles)
      }
      .toMap

    Stream.emits(rowMap.toSeq.sortBy(_._1)).map { case (rowIdx, (cellMap, styleMap)) =>
      StyledRowData(rowIdx, cellMap, styleMap)
    }

  // Helper: Assemble multi-sheet ZIP with dimension elements
  private def assembleMultiSheetZip(
    path: Path,
    resources: Seq[(String, Int, Path, BoundsAccumulator)],
    config: com.tjclp.xl.ooxml.WriterConfig
  ): F[Unit] =
    Sync[F].delay {
      val zip = new ZipOutputStream(new FileOutputStream(path.toFile))
      try
        // 1. Write static parts with all sheet metadata
        writeStaticPartsMultiSync(
          zip,
          resources.map { case (name, idx, _, _) => (name, idx) },
          config
        )

        // 2. Write each worksheet with its dimension
        resources.foreach { case (_, sheetIndex, tempBodyFile, bounds) =>
          val wsEntry = new ZipEntry(s"xl/worksheets/sheet$sheetIndex.xml")
          wsEntry.setMethod(ZipEntry.DEFLATED)
          zip.putNextEntry(wsEntry)

          // Header with dimension
          writeWorksheetHeader(zip, bounds.dimension)

          // Body from temp file
          JFiles.copy(tempBodyFile, zip)

          // Footer
          writeWorksheetFooter(zip)

          zip.closeEntry()
        }
      finally zip.close()
    }

  // Helper: Sync version of writeStaticPartsMulti
  private def writeStaticPartsMultiSync(
    zip: ZipOutputStream,
    sheets: Seq[(String, Int)],
    config: com.tjclp.xl.ooxml.WriterConfig
  ): Unit =
    import com.tjclp.xl.ooxml.*
    import com.tjclp.xl.addressing.SheetName

    val contentTypes = ContentTypes.forSheetIndices(
      sheetIndices = sheets.map(_._2),
      hasStyles = true,
      hasSharedStrings = false
    )
    val rootRels = Relationships.root()
    val workbookRels = Relationships.workbook(
      sheetCount = sheets.size,
      hasStyles = true,
      hasSharedStrings = false
    )

    val sheetRefs = sheets.map { case (name, idx) =>
      SheetRef(SheetName.unsafe(name), idx, s"rId$idx")
    }
    val ooxmlWb = OoxmlWorkbook(sheets = sheetRefs)
    val styles = OoxmlStyles.minimal

    writePartSync(zip, "[Content_Types].xml", contentTypes.toXml, config)
    writePartSync(zip, "_rels/.rels", rootRels.toXml, config)
    writePartSync(zip, "xl/workbook.xml", ooxmlWb.toXml, config)
    writePartSync(zip, "xl/_rels/workbook.xml.rels", workbookRels.toXml, config)
    writePartSync(zip, "xl/styles.xml", styles.toXml, config)

  // ========== ExcelR Implementation (Explicit Error Channel) ==========

  /** Read workbook with explicit error result */
  def readR(path: Path): F[XLResult[Workbook]] =
    Sync[F].delay(XlsxReader.read(path))

  /** Write workbook with explicit error result */
  def writeR(wb: Workbook, path: Path): F[XLResult[Unit]] =
    writeWithR(wb, path, com.tjclp.xl.ooxml.WriterConfig.default)

  /** Write workbook with explicit error result and custom configuration */
  def writeWithR(
    wb: Workbook,
    path: Path,
    config: com.tjclp.xl.ooxml.WriterConfig
  ): F[XLResult[Unit]] =
    Sync[F].delay(XlsxWriter.writeWith(wb, path, config))

  /** Stream rows with explicit error channel */
  def readStreamR(path: Path): Stream[F, Either[XLError, RowData]] =
    readStream(path)
      .map(Right(_))
      .handleErrorWith(e => Stream.emit(Left(XLError.IOError(e.getMessage))))

  /** Stream rows from specific sheet by name with explicit error channel */
  def readSheetStreamR(path: Path, sheetName: String): Stream[F, Either[XLError, RowData]] =
    readSheetStream(path, sheetName)
      .map(Right(_))
      .handleErrorWith(e => Stream.emit(Left(XLError.IOError(e.getMessage))))

  /** Stream rows from specific sheet by index with explicit error channel */
  def readStreamByIndexR(path: Path, sheetIndex: Int): Stream[F, Either[XLError, RowData]] =
    readStreamByIndex(path, sheetIndex)
      .map(Right(_))
      .handleErrorWith(e => Stream.emit(Left(XLError.IOError(e.getMessage))))

  /** Streaming write with explicit error channel */
  def writeStreamR(
    path: Path,
    sheetName: String,
    sheetIndex: Int = 1,
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default,
    dimension: Option[CellRange] = None
  ): Pipe[F, RowData, Either[XLError, Unit]] =
    rows =>
      writeStream(path, sheetName, sheetIndex, config, dimension)(rows)
        .map(_ => Right(()))
        .handleErrorWith(e => Stream.emit(Left(XLError.IOError(e.getMessage))))

  /** Write stream (materializing) with explicit error channel */
  @deprecated("Use writeStreamR for true O(1) streaming", "0.8.0")
  def writeStreamMaterializedR(
    path: Path,
    sheetName: String
  ): Pipe[F, RowData, Either[XLError, Unit]] =
    rows =>
      writeStreamMaterialized(path, sheetName)(rows)
        .map(_ => Right(()))
        .handleErrorWith(e => Stream.emit(Left(XLError.IOError(e.getMessage))))

  /** Write multiple sheets with explicit error channel */
  def writeStreamsSeqR(
    path: Path,
    sheets: Seq[(String, Stream[F, RowData])],
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): F[XLResult[Unit]] =
    writeStreamsSeq(path, sheets, config)
      .map(Right(_))
      .handleErrorWith(e => Async[F].pure(Left(XLError.IOError(e.getMessage))))

  // ========== Private Helpers ==========

  // Helper: Convert Sheet to RowData stream
  private def streamSheet(sheet: Sheet): Stream[F, RowData] =
    val rowMap = sheet.cells.values
      .groupBy(_.ref.row.index1) // Group by 1-based row
      .view
      .mapValues { cells =>
        cells.map(c => c.ref.col.index0 -> c.value).toMap
      }
      .toMap

    Stream.emits(rowMap.toSeq.sortBy(_._1)).map { case (rowIdx, cellMap) =>
      RowData(rowIdx, cellMap)
    }

  // Helper: Convert RowData vector to Sheet
  private def rowDataToSheet(
    sheetName: String,
    rows: Vector[RowData]
  ): Either[XLError, Sheet] =
    import com.tjclp.xl.addressing.{ARef, Column, Row, SheetName}
    import com.tjclp.xl.cells.Cell

    for
      name <- SheetName(sheetName).left.map(err => XLError.InvalidSheetName(sheetName, err))
      cells = rows.flatMap { row =>
        row.cells.map { case (colIdx, value) =>
          val ref = ARef(Column.from0(colIdx), Row.from1(row.rowIndex))
          Cell(ref, value)
        }
      }
      sheet = Sheet(name, cells.map(c => c.ref -> c).toMap)
    yield sheet

  // Helper: Write static OOXML parts to ZIP
  private def writeStaticParts(
    zip: ZipOutputStream,
    sheetName: String,
    sheetIndex: Int,
    config: com.tjclp.xl.ooxml.WriterConfig
  ): F[Unit] =
    import com.tjclp.xl.ooxml.*
    import com.tjclp.xl.addressing.SheetName

    val contentTypes =
      ContentTypes.forSheetIndices(Seq(sheetIndex), hasStyles = true, hasSharedStrings = false)
    val rootRels = Relationships.root()
    val workbookRels =
      Relationships.workbookWithIndices(Seq(sheetIndex), hasStyles = true, hasSharedStrings = false)

    // Create minimal workbook with one sheet (use provided sheetIndex)
    val ooxmlWb = OoxmlWorkbook(
      sheets = Vector(SheetRef(SheetName.unsafe(sheetName), sheetIndex, "rId1"))
    )

    // Minimal styles
    val styles = OoxmlStyles.minimal

    for
      _ <- writePart(zip, "[Content_Types].xml", contentTypes.toXml, config)
      _ <- writePart(zip, "_rels/.rels", rootRels.toXml, config)
      _ <- writePart(zip, "xl/workbook.xml", ooxmlWb.toXml, config)
      _ <- writePart(zip, "xl/_rels/workbook.xml.rels", workbookRels.toXml, config)
      _ <- writePart(zip, "xl/styles.xml", styles.toXml, config)
    yield ()

  // Helper: Write static OOXML parts for multiple sheets
  private def writeStaticPartsMulti(
    zip: ZipOutputStream,
    sheets: Seq[(String, Int)], // (name, sheetIndex)
    config: com.tjclp.xl.ooxml.WriterConfig
  ): F[Unit] =
    import com.tjclp.xl.ooxml.*
    import com.tjclp.xl.addressing.SheetName

    val contentTypes = ContentTypes.forSheetIndices(
      sheetIndices = sheets.map(_._2),
      hasStyles = true,
      hasSharedStrings = false
    )
    val rootRels = Relationships.root()
    val workbookRels = Relationships.workbook(
      sheetCount = sheets.size,
      hasStyles = true,
      hasSharedStrings = false
    )

    // Create workbook with all sheets
    val sheetRefs = sheets.map { case (name, idx) =>
      SheetRef(SheetName.unsafe(name), idx, s"rId$idx")
    }
    val ooxmlWb = OoxmlWorkbook(sheets = sheetRefs)

    // Minimal styles
    val styles = OoxmlStyles.minimal

    for
      _ <- writePart(zip, "[Content_Types].xml", contentTypes.toXml, config)
      _ <- writePart(zip, "_rels/.rels", rootRels.toXml, config)
      _ <- writePart(zip, "xl/workbook.xml", ooxmlWb.toXml, config)
      _ <- writePart(zip, "xl/_rels/workbook.xml.rels", workbookRels.toXml, config)
      _ <- writePart(zip, "xl/styles.xml", styles.toXml, config)
    yield ()

  // Helper: Write a single XML part to ZIP
  private def writePart(
    zip: ZipOutputStream,
    entryName: String,
    xml: scala.xml.Elem,
    config: com.tjclp.xl.ooxml.WriterConfig
  ): F[Unit] =
    Sync[F].delay {
      import com.tjclp.xl.ooxml.{XmlUtil, Compression}

      // Use config for XML formatting
      val xmlString = if config.prettyPrint then XmlUtil.prettyPrint(xml) else XmlUtil.compact(xml)
      val bytes = xmlString.getBytes(StandardCharsets.UTF_8)

      val entry = new ZipEntry(entryName)
      entry.setMethod(config.compression.zipMethod)

      // For STORED method, must set size and CRC before writing
      // For DEFLATED, ZipOutputStream computes these automatically
      config.compression match
        case Compression.Stored =>
          entry.setSize(bytes.length)
          entry.setCompressedSize(bytes.length)
          // Calculate CRC32
          val crc = new CRC32()
          crc.update(bytes)
          entry.setCrc(crc.getValue)
        case Compression.Deflated =>
          () // ZipOutputStream handles automatically

      zip.putNextEntry(entry)
      zip.write(bytes)
      zip.closeEntry()
    }

  // Helper: Find sheet index by name from workbook.xml
  private def findSheetIndexByName(zipFile: ZipFile, sheetName: String): F[Option[Int]] =
    Sync[F].delay {
      val wbEntry = Option(zipFile.getEntry("xl/workbook.xml"))
      wbEntry.flatMap { entry =>
        val xmlContent = new String(zipFile.getInputStream(entry).readAllBytes(), "UTF-8")
        XmlSecurity.parseSafe(xmlContent, "xl/workbook.xml").toOption.flatMap { wbXml =>
          // Parse sheet elements
          val sheets = (wbXml \\ "sheet").toSeq
          sheets
            .find { sheetElem =>
              (sheetElem \ "@name").text == sheetName
            }
            .map { sheetElem =>
              (sheetElem \ "@sheetId").text.toInt
            }
        }
      }
    }

object ExcelIO:
  /** Create default ExcelIO instance */
  def instance[F[_]: Async]: ExcelIO[F] = new ExcelIO[F](_ => Async[F].unit)

  /** Create ExcelIO with custom warning handler */
  def withWarnings[F[_]: Async](handler: XlsxReader.Warning => F[Unit]): ExcelIO[F] =
    new ExcelIO[F](handler)
