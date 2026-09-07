package com.tjclp.xl.cli.commands

import java.nio.file.Path

import cats.effect.IO
import cats.implicits.*
import com.tjclp.xl.addressing.{ARef, CellRange, Column, SheetName}
import com.tjclp.xl.cli.contract.CliException
import com.tjclp.xl.cli.helpers.Resolve
import com.tjclp.xl.io.{ExcelIO, RowData}
import com.tjclp.xl.ooxml.metadata.LightMetadata

/**
 * The streaming `bounds` verb: the worksheet's `<dimension>` element (instant for any file size),
 * or an O(1)-memory scan of its rows. The other streaming reads — `view`, `cell`, `search`,
 * `stats`, `filter` — are [[com.tjclp.xl.cli.read.Reads]] queries over
 * [[com.tjclp.xl.cli.read.SheetSource.streaming]] (W2.4).
 */
object StreamingReadCommands:

  private val excel = ExcelIO.instance[IO]

  /**
   * Compute used range bounds using dimension element (instant for any file size).
   *
   * Uses <dimension ref="..."> from worksheet metadata. Falls back to streaming scan if dimension
   * is missing.
   */
  def bounds(
    filePath: Path,
    sheetNameOpt: Option[String]
  ): IO[String] =
    boundsDimension(filePath, sheetNameOpt)

  /**
   * Compute used range bounds using dimension element (instant).
   *
   * Reads only the <dimension ref="..."> element from worksheet XML. Typically completes in <100ms.
   */
  def boundsDimension(
    filePath: Path,
    sheetNameOpt: Option[String]
  ): IO[String] =
    excel.readMetadata(filePath).flatMap(meta => boundsFromMetadata(meta, filePath, sheetNameOpt))

  /**
   * [[boundsDimension]] over metadata the caller has already read — the runner reads it under its
   * `IO_READ` classification, so a missing or unreadable file gets the same code as on every verb
   * while a sheet the metadata does not list keeps its own failure.
   */
  def boundsFromMetadata(
    meta: LightMetadata,
    filePath: Path,
    sheetNameOpt: Option[String]
  ): IO[String] =
    boundsTarget(meta, sheetNameOpt) match
      case Left(err) => IO.raiseError(CliException(err))
      case Right((sheetName, Some(range))) =>
        val rowCount = range.end.row.index1 - range.start.row.index1 + 1
        val colCount = range.end.col.index0 - range.start.col.index0 + 1
        IO.pure(
          s"""Sheet: $sheetName
             |Used range: ${range.toA1} (from dimension element)
             |Rows: ${range.start.row.index1}-${range.end.row.index1} ($rowCount total)
             |Columns: ${range.start.col.toLetter}-${range.end.col.toLetter} ($colCount total)""".stripMargin
        )
      case Right((_, None)) =>
        // Fallback to streaming scan
        boundsScan(filePath, sheetNameOpt)

  /**
   * `bounds` as data (`bounds --json`): `{sheet, range, dimension}` — `range` the used range in A1
   * form or `null` for an empty sheet, `dimension` true when it came from the worksheet's
   * `<dimension>` element and false when from a streaming scan (`--scan`, or a sheet without one).
   * Takes the metadata already read (see [[boundsFromMetadata]]).
   */
  def boundsData(
    meta: LightMetadata,
    filePath: Path,
    sheetNameOpt: Option[String],
    scan: Boolean
  ): IO[ujson.Value] =
    def scanned: IO[ujson.Value] =
      scanBounds(filePath, sheetNameOpt).map { (sheetName, acc) =>
        boundsJson(sheetName, acc.range, fromDimension = false)
      }
    if scan then scanned
    else
      boundsTarget(meta, sheetNameOpt) match
        case Left(err) => IO.raiseError(CliException(err))
        case Right((sheetName, Some(range))) =>
          IO.pure(boundsJson(sheetName, Some(range), fromDimension = true))
        case Right((_, None)) => scanned

  private def boundsJson(
    sheetName: String,
    range: Option[CellRange],
    fromDimension: Boolean
  ): ujson.Value =
    ujson.Obj(
      "sheet" -> ujson.Str(sheetName),
      "range" -> range.fold[ujson.Value](ujson.Null)(r => ujson.Str(r.toA1)),
      "dimension" -> ujson.Bool(fromDimension)
    )

  /**
   * The sheet `bounds` reads, by THE sheet rule over the metadata ([[Resolve.sheetName]]: `-s`,
   * else the only sheet, else `SHEET_REQUIRED`) — as its display name and the `<dimension>` range
   * the metadata carries for it (None when the worksheet has none).
   */
  private def boundsTarget(
    meta: LightMetadata,
    sheetNameOpt: Option[String]
  ): Either[com.tjclp.xl.cli.contract.CliError, (String, Option[CellRange])] =
    Resolve.sheetName(meta, sheetNameOpt, None, "bounds").map { name =>
      (name.value, meta.sheets.find(_.name == name).flatMap(_.dimension))
    }

  /**
   * Compute used range bounds using streaming scan (accurate but slower).
   *
   * Tracks min/max row/col during full scan. Memory: O(1). Time: O(rows).
   */
  def boundsScan(
    filePath: Path,
    sheetNameOpt: Option[String]
  ): IO[String] =
    scanBounds(filePath, sheetNameOpt).map { (sheetName, acc) =>
      acc.format(sheetName, fromScan = true)
    }

  /**
   * The scan itself: the sheet THE rule selects (its name is what the result is reported under) and
   * the accumulated bounds.
   */
  private def scanBounds(
    filePath: Path,
    sheetNameOpt: Option[String]
  ): IO[(String, BoundsAccumulator)] =
    resolveSheetName(filePath, sheetNameOpt, None, "bounds").flatMap { sheetName =>
      excel
        .readSheetStream(filePath, sheetName)
        .compile
        .fold(BoundsAccumulator.empty)(_.update(_))
        .map(acc => (sheetName, acc))
    }

  /** Accumulator for streaming bounds computation. */
  private case class BoundsAccumulator(
    minRow: Option[Int],
    maxRow: Option[Int],
    minCol: Option[Int],
    maxCol: Option[Int],
    cellCount: Long
  ):
    // IterableOps: .min/.max safe because row.cells.isEmpty checked first
    @SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
    def update(row: RowData): BoundsAccumulator =
      if row.cells.isEmpty then this
      else
        val cols = row.cells.keys
        BoundsAccumulator(
          minRow = Some(minRow.fold(row.rowIndex)(_ min row.rowIndex)),
          maxRow = Some(maxRow.fold(row.rowIndex)(_ max row.rowIndex)),
          minCol = Some(minCol.fold(cols.min)(_ min cols.min)),
          maxCol = Some(maxCol.fold(cols.max)(_ max cols.max)),
          cellCount = cellCount + row.cells.size
        )

    /** The bounding range of every non-empty cell seen, None when the sheet is empty. */
    def range: Option[CellRange] = (minRow, maxRow, minCol, maxCol) match
      case (Some(r1), Some(r2), Some(c1), Some(c2)) =>
        Some(CellRange(ARef.from0(c1, r1 - 1), ARef.from0(c2, r2 - 1))) // rowIndex is 1-based
      case _ => None

    def format(sheetName: String, fromScan: Boolean = false): String =
      val source = if fromScan then "(from scan)" else "(streaming)"
      (minRow, maxRow, minCol, maxCol) match
        case (Some(r1), Some(r2), Some(c1), Some(c2)) =>
          val startRef = ARef.from0(c1, r1 - 1) // rowIndex is 1-based
          val endRef = ARef.from0(c2, r2 - 1)
          val rowCount = r2 - r1 + 1
          val colCount = c2 - c1 + 1
          s"""Sheet: $sheetName
             |Used range: ${startRef.toA1}:${endRef.toA1} $source
             |Rows: $r1-$r2 ($rowCount total)
             |Columns: ${Column.from0(c1).toLetter}-${Column.from0(c2).toLetter} ($colCount total)
             |Non-empty: $cellCount cells""".stripMargin
        case _ =>
          s"""Sheet: $sheetName
             |Used range: (empty) $source
             |Non-empty: 0 cells""".stripMargin

  private object BoundsAccumulator:
    def empty: BoundsAccumulator = BoundsAccumulator(None, None, None, None, 0)

  /**
   * THE sheet rule over `workbook.xml` ([[Resolve.sheetName]], ADR-017 §2.5): the ref's qualifier,
   * else `-s`, else the only sheet of a single-sheet book, else `SHEET_REQUIRED` naming `context`.
   * Reads the metadata part only.
   */
  private def resolveSheetName(
    filePath: Path,
    sheetNameOpt: Option[String],
    qualified: Option[SheetName],
    context: String
  ): IO[String] =
    excel.readMetadata(filePath).flatMap { meta =>
      IO.fromEither(
        Resolve
          .sheetName(meta, sheetNameOpt, qualified, context)
          .map(_.value)
          .left
          .map(CliException(_))
      )
    }
