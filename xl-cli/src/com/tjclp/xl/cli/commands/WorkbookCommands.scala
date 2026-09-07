package com.tjclp.xl.cli.commands

import java.nio.file.Path

import cats.effect.IO
import cats.syntax.all.*
import com.tjclp.xl.{*, given}
import com.tjclp.xl.cli.output.Markdown
import com.tjclp.xl.io.{ExcelIO, RowData}
import com.tjclp.xl.ooxml.metadata.{LightMetadata, SheetInfo}

/**
 * Workbook-level command handlers.
 *
 * Commands that operate on the workbook as a whole (no sheet context needed).
 */
object WorkbookCommands:

  private val excel = ExcelIO.instance[IO]

  /**
   * List all sheets with statistics (full mode - loads cell data).
   *
   * Shows: name, used range, cell count, formula count.
   */
  def sheets(wb: Workbook): IO[String] =
    val sheetStats = wb.sheets.map { s =>
      val usedRange = s.usedRange
      val cellCount = s.cells.size
      val formulaCount = s.cells.values.count(_.isFormula)
      (s.name.value, usedRange, cellCount, formulaCount)
    }
    IO.pure(Markdown.renderSheetList(sheetStats))

  /**
   * The full listing as data (`sheets --stats --json`): `[{name, index, state, dimension, cells,
   * formulas}]`, `index` 1-based in workbook order, `state` one of `visible`/`hidden`/`veryHidden`,
   * `dimension` the used range or `null` for an empty sheet.
   */
  def sheetsData(wb: Workbook): ujson.Value =
    ujson.Arr.from(wb.sheets.zipWithIndex.map { (s, idx) =>
      ujson.Obj(
        "name" -> ujson.Str(s.name.value),
        "index" -> ujson.Num(idx + 1),
        "state" -> ujson.Str(wb.getSheetState(s.name).getOrElse("visible")),
        "dimension" -> s.usedRange.fold[ujson.Value](ujson.Null)(r => ujson.Str(r.toA1)),
        "cells" -> ujson.Num(s.cells.size),
        "formulas" -> ujson.Num(s.cells.values.count(_.isFormula))
      )
    })

  /**
   * List all sheets with dimensions only (quick mode - no cell data).
   *
   * Uses lightweight metadata reader for instant response on large files. Shows: name, dimension
   * (from worksheet metadata or streaming scan if missing), visibility state.
   */
  def sheetsQuick(filePath: Path): IO[String] =
    sheetInfos(filePath).map(Markdown.renderSheetListQuick)

  /** The quick listing as data (`sheets --json`): `[{name, index, state, dimension}]`. */
  def sheetsQuickData(filePath: Path): IO[ujson.Value] =
    sheetInfos(filePath).map { infos =>
      ujson.Arr.from(infos.zipWithIndex.map { (info, idx) =>
        ujson.Obj(
          "name" -> ujson.Str(info.name.value),
          "index" -> ujson.Num(idx + 1),
          "state" -> ujson.Str(info.state.getOrElse("visible")),
          "dimension" -> info.dimension.fold[ujson.Value](ujson.Null)(r => ujson.Str(r.toA1))
        )
      })
    }

  /**
   * Sheet metadata with every dimension filled in. When dimension metadata is missing from the
   * worksheet XML, falls back to a streaming scan to compute accurate bounds (O(rows) time, O(1)
   * memory per sheet); a sheet the scan finds empty keeps `dimension = None`.
   */
  private def sheetInfos(filePath: Path): IO[Vector[SheetInfo]] =
    excel.readMetadata(filePath).flatMap { meta =>
      // Find sheets with missing dimensions
      val missingIndices = meta.sheets.zipWithIndex.collect {
        case (info, idx) if info.dimension.isEmpty => idx
      }

      if missingIndices.isEmpty then
        // All sheets have dimensions from metadata - fast path
        IO.pure(meta.sheets)
      else
        // Scan missing dimensions sequentially (to avoid multiple zip file handles)
        missingIndices
          .traverse { idx =>
            scanSheetBounds(filePath, idx + 1).map(bounds => (idx, bounds)) // 1-based index
          }
          .map { scannedBounds =>
            // Build map of index -> scanned dimension
            val boundsMap = scannedBounds.toMap
            // Update sheets with scanned dimensions
            meta.sheets.zipWithIndex.map { case (info, idx) =>
              if info.dimension.isEmpty then
                boundsMap.get(idx).flatten match
                  case Some(range) => info.copy(dimension = Some(range))
                  case None => info // Empty sheet, keep as unknown
              else info
            }
          }
    }

  /**
   * Scan a sheet to compute its bounds using streaming.
   *
   * Returns the bounding range of all non-empty cells, or None if sheet is empty. Uses O(1) memory
   * via streaming, O(rows) time.
   */
  @SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
  private def scanSheetBounds(filePath: Path, sheetIndex: Int): IO[Option[CellRange]] =
    excel
      .readStreamByIndex(filePath, sheetIndex)
      .compile
      .fold((Int.MaxValue, Int.MaxValue, Int.MinValue, Int.MinValue)) {
        case ((minR, minC, maxR, maxC), row) =>
          if row.cells.isEmpty then (minR, minC, maxR, maxC)
          else
            val cols = row.cells.keys
            (
              minR min row.rowIndex,
              cols.min min minC,
              maxR max row.rowIndex,
              cols.max max maxC
            )
      }
      .map { case (minR, minC, maxR, maxC) =>
        if minR == Int.MaxValue then None // No cells found - sheet is empty
        else
          Some(
            CellRange(
              ARef.from0(minC, minR - 1), // rowIndex is 1-based, ARef.from0 expects 0-based row
              ARef.from0(maxC, maxR - 1)
            )
          )
      }

  /**
   * List defined names (named ranges) from full workbook.
   */
  def names(wb: Workbook): IO[String] =
    val names = wb.metadata.definedNames
    if names.isEmpty then IO.pure("No defined names in workbook")
    else
      // Filter out hidden names unless user explicitly wants them
      val visibleNames = names.filterNot(_.hidden)
      if visibleNames.isEmpty then IO.pure("No visible defined names in workbook")
      else
        val maxNameLen = visibleNames.map(_.name.length).foldLeft(0)(_ max _)
        val lines = visibleNames.map { dn =>
          val scope = dn.localSheetId match
            case Some(idx) =>
              wb.sheets.lift(idx).map(s => s" (${s.name.value})").getOrElse(s" (sheet $idx)")
            case None => ""
          val paddedName = dn.name.padTo(maxNameLen, ' ')
          s"$paddedName  ${dn.formula}$scope"
        }
        IO.pure(lines.mkString("\n"))

  /**
   * Defined names as data (`names --json`): `[{name, refersTo, scope, hidden}]`, every name
   * including hidden ones (the text listing omits those), `scope` the sheet name for a sheet-scoped
   * name and `null` for a workbook-scoped one.
   */
  def namesData(filePath: Path): IO[ujson.Value] =
    excel.readMetadata(filePath).map { meta =>
      ujson.Arr.from(meta.definedNames.map { dn =>
        val scope = dn.localSheetId.fold[ujson.Value](ujson.Null) { idx =>
          ujson.Str(meta.sheets.lift(idx).map(_.name.value).getOrElse(s"sheet $idx"))
        }
        ujson.Obj(
          "name" -> ujson.Str(dn.name),
          "refersTo" -> ujson.Str(dn.formula),
          "scope" -> scope,
          "hidden" -> ujson.Bool(dn.hidden)
        )
      })
    }

  /**
   * List defined names using lightweight metadata reader (instant for any file size).
   */
  def namesLight(filePath: Path): IO[String] =
    excel.readMetadata(filePath).map { meta =>
      val names = meta.definedNames
      if names.isEmpty then "No defined names in workbook"
      else
        val visibleNames = names.filterNot(_.hidden)
        if visibleNames.isEmpty then "No visible defined names in workbook"
        else
          val maxNameLen = visibleNames.map(_.name.length).foldLeft(0)(_ max _)
          val lines = visibleNames.map { dn =>
            val scope = dn.localSheetId match
              case Some(idx) =>
                meta.sheets.lift(idx).map(s => s" (${s.name.value})").getOrElse(s" (sheet $idx)")
              case None => ""
            val paddedName = dn.name.padTo(maxNameLen, ' ')
            s"$paddedName  ${dn.formula}$scope"
          }
          lines.mkString("\n")
    }
