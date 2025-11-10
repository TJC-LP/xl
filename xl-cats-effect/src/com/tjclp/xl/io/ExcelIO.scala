package com.tjclp.xl.io

import cats.effect.{Async, Sync}
import cats.syntax.all.*
import fs2.Stream
import java.nio.file.Path
import com.tjclp.xl.{Workbook, XLError}
import com.tjclp.xl.ooxml.{XlsxReader, XlsxWriter}

/** Cats Effect interpreter for Excel operations.
  *
  * Wraps pure XlsxReader/Writer with effect type F[_].
  * Future: Will add fs2-data-xml streaming for memory efficiency.
  */
class ExcelIO[F[_]: Async] extends Excel[F]:

  /** Read workbook from XLSX file */
  def read(path: Path): F[Workbook] =
    Sync[F].delay(XlsxReader.read(path)).flatMap {
      case Right(wb) => Async[F].pure(wb)
      case Left(err) => Async[F].raiseError(new Exception(s"Failed to read XLSX: ${err.message}"))
    }

  /** Write workbook to XLSX file */
  def write(wb: Workbook, path: Path): F[Unit] =
    Sync[F].delay(XlsxWriter.write(wb, path)).flatMap {
      case Right(_) => Async[F].unit
      case Left(err) => Async[F].raiseError(new Exception(s"Failed to write XLSX: ${err.message}"))
    }

  /** Stream rows from first sheet.
    *
    * Current implementation: Materializes workbook then streams rows.
    * Future: Will use fs2-data-xml for true streaming (constant memory).
    */
  def readStream(path: Path): Stream[F, RowData] =
    Stream.eval(read(path)).flatMap { wb =>
      if wb.sheets.isEmpty then Stream.empty
      else streamSheet(wb.sheets(0))
    }

  /** Stream rows from sheet by name */
  def readSheetStream(path: Path, sheetName: String): Stream[F, RowData] =
    Stream.eval(read(path)).flatMap { wb =>
      wb(sheetName) match
        case Right(sheet) => streamSheet(sheet)
        case Left(err) => Stream.raiseError[F](new Exception(s"Sheet not found: $sheetName"))
    }

  /** Write rows to XLSX file.
    *
    * Current implementation: Materializes all rows then writes.
    * Future: Will use fs2-data-xml for true streaming.
    */
  def writeStream(path: Path, sheetName: String): fs2.Pipe[F, RowData, Unit] =
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

  // Helper: Convert Sheet to RowData stream
  private def streamSheet(sheet: com.tjclp.xl.Sheet): Stream[F, RowData] =
    val rowMap = sheet.cells.values
      .groupBy(_.ref.row.index1)  // Group by 1-based row
      .view.mapValues { cells =>
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
  ): Either[XLError, com.tjclp.xl.Sheet] =
    import com.tjclp.xl.{Sheet, SheetName, Cell, ARef, Column, Row}

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

object ExcelIO:
  /** Create default ExcelIO instance */
  def instance[F[_]: Async]: Excel[F] = new ExcelIO[F]
