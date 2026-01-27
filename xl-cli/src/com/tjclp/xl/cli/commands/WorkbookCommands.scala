package com.tjclp.xl.cli.commands

import java.nio.file.Path

import cats.effect.IO
import com.tjclp.xl.{*, given}
import com.tjclp.xl.cli.output.Markdown
import com.tjclp.xl.io.ExcelIO
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
   * List all sheets with dimensions only (quick mode - no cell data).
   *
   * Uses lightweight metadata reader for instant response on large files. Shows: name, dimension
   * (from worksheet metadata), visibility state.
   */
  def sheetsQuick(filePath: Path): IO[String] =
    excel.readMetadata(filePath).map { meta =>
      Markdown.renderSheetListQuick(meta.sheets)
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
