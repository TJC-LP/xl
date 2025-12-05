package com.tjclp.xl.cli.commands

import cats.effect.IO
import com.tjclp.xl.{*, given}
import com.tjclp.xl.cli.output.Markdown

/**
 * Workbook-level command handlers.
 *
 * Commands that operate on the workbook as a whole (no sheet context needed).
 */
object WorkbookCommands:

  /**
   * List all sheets with statistics.
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
   * List defined names (named ranges).
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
