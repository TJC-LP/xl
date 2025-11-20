package com.tjclp.xl.io

import cats.effect.{Async, Sync, Resource}
import cats.syntax.all.*
import fs2.{Stream, Pipe}
import java.nio.file.Path
import java.io.{FileOutputStream, FileInputStream}
import java.util.zip.{ZipOutputStream, ZipEntry, CRC32, ZipInputStream, ZipFile}
import java.nio.charset.StandardCharsets
import com.tjclp.xl.api.Workbook
import com.tjclp.xl.errors.{XLError, XLResult}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.ooxml.{XlsxReader, XlsxWriter, SharedStrings, XmlSecurity}
import fs2.data.xml
import fs2.data.xml.XmlEvent

/**
 * Cats Effect interpreter for Excel operations.
 *
 * Implements both Excel[F] (errors-raising) and ExcelR[F] (explicit errors channel). Wraps pure
 * XlsxReader/Writer with effect type F[_].
 */
class ExcelIO[F[_]: Async](warningHandler: XlsxReader.Warning => F[Unit])
    extends Excel[F]
    with ExcelR[F]:

  /** Read workbook from XLSX file */
  def read(path: Path): F[Workbook] =
    Sync[F].delay(XlsxReader.readWithWarnings(path)).flatMap {
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

  /**
   * Stream rows from first sheet with constant memory.
   *
   * Uses fs2-data-xml pull parsing - never materializes full worksheet in memory.
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
            // Stream first worksheet
            val wsEntry = Option(zipFile.getEntry("xl/worksheets/sheet1.xml"))
            wsEntry match
              case Some(entry) =>
                val wsBytes = fs2.io.readInputStream[F](
                  Sync[F].delay(zipFile.getInputStream(entry)),
                  chunkSize = 4096
                )
                StreamingXmlReader.parseWorksheetStream(wsBytes, sst)
              case None =>
                Stream.empty
          }
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
            // Find sheets index by parsing workbook.xml
            findSheetIndexByName(zipFile, sheetName)
          }
          .flatMap {
            case Some(sheetIndex) =>
              readStreamByIndex(zipFile, sheetIndex)
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

  // Helper: Stream rows from specific sheet index using open ZipFile
  private def readStreamByIndex(zipFile: ZipFile, sheetIndex: Int): Stream[F, RowData] =
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
            val wsBytes = fs2.io.readInputStream[F](
              Sync[F].delay(zipFile.getInputStream(entry)),
              chunkSize = 4096
            )
            StreamingXmlReader.parseWorksheetStream(wsBytes, sst)
          case None =>
            Stream.raiseError[F](
              new Exception(s"Worksheet not found: sheet$sheetIndex.xml")
            )
      }

  /**
   * Write rows to XLSX file.
   *
   * Current implementation: Materializes all rows then writes. Future: Will use fs2-data-xml for
   * true streaming.
   */
  def writeStream(path: Path, sheetName: String): Pipe[F, RowData, Unit] =
    rows =>
      Stream.eval {
        rows.compile.toVector.flatMap { rowVec =>
          // Convert RowData to Sheet
          val sheet = rowDataToSheet(sheetName, rowVec)
          sheet match
            case Right(s) =>
              // Create workbook with single sheet
              val wb = Workbook(Vector(s))
              write(wb, path)
            case Left(err) =>
              Async[F].raiseError(new Exception(s"Failed to create sheet: ${err.message}"))
        }
      }

  /**
   * True streaming write with constant memory.
   *
   * Uses fs2-data-xml events - never materializes full dataset. Can write unlimited rows with ~50MB
   * memory.
   */
  def writeStreamTrue(
    path: Path,
    sheetName: String,
    sheetIndex: Int = 1,
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): Pipe[F, RowData, Unit] =
    rows =>
      // Validate sheets index
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
            Stream.eval(writeStaticParts(zip, sheetName, sheetIndex, config)) ++
              // 2. Open worksheet entry (use sheetIndex for filename)
              // Note: Must use DEFLATED for streamed worksheets (STORED requires precomputed size/CRC)
              Stream.eval(
                Sync[F].delay {
                  val entry = new ZipEntry(s"xl/worksheets/sheet$sheetIndex.xml")
                  entry.setMethod(ZipEntry.DEFLATED) // Always DEFLATED for streamed content
                  zip.putNextEntry(entry)
                }
              ) ++
              // 3. Stream XML events → bytes → ZIP
              (Stream.emit(XmlEvent.XmlDecl("1.0", Some("UTF-8"), Some(true))) ++
                StreamingXmlWriter.worksheetEvents[F](rows))
                .through(xml.render.raw())
                .through(fs2.text.utf8.encode)
                .chunks
                .evalMap(chunk => Sync[F].delay(zip.write(chunk.toArray))) ++
              // 4. Close entry
              Stream.eval(Sync[F].delay(zip.closeEntry()))
          }
          .drain

  /**
   * Write multiple sheets sequentially with constant memory.
   *
   * Each sheet is streamed in order without materializing the full dataset.
   */
  def writeStreamsSeqTrue(
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
      Stream
        .bracket(
          Sync[F].delay(new ZipOutputStream(new FileOutputStream(path.toFile)))
        )(zip => Sync[F].delay(zip.close()))
        .flatMap { zip =>
          // Auto-assign sheet indices 1, 2, 3...
          val sheetsWithIndices = sheets.zipWithIndex.map { case ((name, rows), idx) =>
            (name, idx + 1, rows)
          }

          // 1. Write static parts with all sheet metadata
          Stream.eval(
            writeStaticPartsMulti(
              zip,
              sheetsWithIndices.map { case (name, idx, _) =>
                (name, idx)
              },
              config
            )
          ) ++
            // 2. Stream each sheet sequentially
            Stream
              .emits(sheetsWithIndices)
              .flatMap { case (name, sheetIndex, rows) =>
                // Open worksheet entry
                // Note: Must use DEFLATED for streamed worksheets (STORED requires precomputed size/CRC)
                Stream.eval(
                  Sync[F].delay {
                    val entry = new ZipEntry(s"xl/worksheets/sheet$sheetIndex.xml")
                    entry.setMethod(ZipEntry.DEFLATED) // Always DEFLATED for streamed content
                    zip.putNextEntry(entry)
                  }
                ) ++
                  // Stream XML events → bytes → ZIP
                  (Stream.emit(XmlEvent.XmlDecl("1.0", Some("UTF-8"), Some(true))) ++
                    StreamingXmlWriter.worksheetEvents[F](rows))
                    .through(xml.render.raw())
                    .through(fs2.text.utf8.encode)
                    .chunks
                    .evalMap(chunk => Sync[F].delay(zip.write(chunk.toArray))) ++
                  // Close entry
                  Stream.eval(Sync[F].delay(zip.closeEntry()))
              }
        }
        .compile
        .drain

  // ========== ExcelR Implementation (Explicit Error Channel) ==========

  /** Read workbook with explicit errors result */
  def readR(path: Path): F[XLResult[Workbook]] =
    Sync[F].delay(XlsxReader.read(path))

  /** Write workbook with explicit errors result */
  def writeR(wb: Workbook, path: Path): F[XLResult[Unit]] =
    writeWithR(wb, path, com.tjclp.xl.ooxml.WriterConfig.default)

  /** Write workbook with explicit errors result and custom configuration */
  def writeWithR(
    wb: Workbook,
    path: Path,
    config: com.tjclp.xl.ooxml.WriterConfig
  ): F[XLResult[Unit]] =
    Sync[F].delay(XlsxWriter.writeWith(wb, path, config))

  /** Stream rows with explicit errors channel */
  def readStreamR(path: Path): Stream[F, Either[XLError, RowData]] =
    readStream(path)
      .map(Right(_))
      .handleErrorWith(e => Stream.emit(Left(XLError.IOError(e.getMessage))))

  /** Stream rows from specific sheet by name with explicit errors channel */
  def readSheetStreamR(path: Path, sheetName: String): Stream[F, Either[XLError, RowData]] =
    readSheetStream(path, sheetName)
      .map(Right(_))
      .handleErrorWith(e => Stream.emit(Left(XLError.IOError(e.getMessage))))

  /** Stream rows from specific sheet by index with explicit errors channel */
  def readStreamByIndexR(path: Path, sheetIndex: Int): Stream[F, Either[XLError, RowData]] =
    readStreamByIndex(path, sheetIndex)
      .map(Right(_))
      .handleErrorWith(e => Stream.emit(Left(XLError.IOError(e.getMessage))))

  /** Write stream with explicit errors channel */
  def writeStreamR(path: Path, sheetName: String): Pipe[F, RowData, Either[XLError, Unit]] =
    rows =>
      writeStream(path, sheetName)(rows)
        .map(_ => Right(()))
        .handleErrorWith(e => Stream.emit(Left(XLError.IOError(e.getMessage))))

  /** True streaming write with explicit errors channel */
  def writeStreamTrueR(
    path: Path,
    sheetName: String,
    sheetIndex: Int = 1,
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): Pipe[F, RowData, Either[XLError, Unit]] =
    rows =>
      writeStreamTrue(path, sheetName, sheetIndex, config)(rows)
        .map(_ => Right(()))
        .handleErrorWith(e => Stream.emit(Left(XLError.IOError(e.getMessage))))

  /** Write multiple sheets with explicit errors channel */
  def writeStreamsSeqTrueR(
    path: Path,
    sheets: Seq[(String, Stream[F, RowData])],
    config: com.tjclp.xl.ooxml.WriterConfig = com.tjclp.xl.ooxml.WriterConfig.default
  ): F[XLResult[Unit]] =
    writeStreamsSeqTrue(path, sheets, config)
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
