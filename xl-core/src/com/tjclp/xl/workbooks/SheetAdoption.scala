package com.tjclp.xl.workbooks

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.ooxml.FormulaStorage
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.font.Font

/** Conservative add-only sheet adoption until workbook-scoped package migration exists (GH-526). */
private[workbooks] object SheetAdoption:
  def adopt(destination: Workbook, sheet: Sheet, from: Workbook): XLResult[Workbook] =
    if !from.sheets.contains(sheet) then
      refuse(sheet, "sheet does not match a member of the supplied source workbook")
    else if destination == from then Right(destination)
    else
      val affected = sheet.cells.iterator.flatMap { cell =>
        cell.value match
          case CellValue.Formula(expression, _, _) =>
            val ordinals = FormulaStorage.externalOrdinals(expression).toVector.sorted
            Option.when(ordinals.nonEmpty)(
              s"${cell.ref.toA1}: $expression (external ordinals ${ordinals.mkString(", ")})"
            )
          case _ => None
      }.toVector
      if affected.nonEmpty then
        refuse(
          sheet,
          "external-workbook references could dangle or silently rebind in the destination; " +
            affected.mkString("; ")
        )
      else if from.sourceContext.nonEmpty then
        refuse(
          sheet,
          "source-backed sheet preservation requires package migration, not yet supported"
        )
      else if from.metadata.defaultFont.getOrElse(Font.default) !=
          destination.metadata.defaultFont.getOrElse(Font.default) ||
          from.metadata.theme != destination.metadata.theme
      then refuse(sheet, "source and destination workbook default font or theme differ")
      else if sheet.tables.nonEmpty || sheet.drawings.nonEmpty ||
        sheet.conditionalFormats.nonEmpty || sheet.dataValidations.nonEmpty
      then
        refuse(sheet, "tables, drawings, conditional formats and validations are not yet supported")
      else if destination.sheets.exists(_.name.value.equalsIgnoreCase(sheet.name.value)) ||
        destination.sourceContext.exists(
          _.preservedSheetNames.exists(_.value.equalsIgnoreCase(sheet.name.value))
        )
      then Left(XLError.DuplicateSheet(sheet.name.value))
      else
        // Visibility lives in workbook metadata keyed by name, not on the sheet. Write the source
        // state explicitly (None = visible) so a stale destination entry for this name is replaced.
        val state = from.getSheetState(sheet.name)
        val anyVisible = destination.sheets.exists(s => destination.getSheetState(s.name).isEmpty)
        if state.nonEmpty && !anyVisible then
          refuse(sheet, s"it is ${state.getOrElse("")} and the destination has no visible sheet")
        else
          val placed = destination.put(sheet)
          // Sheet names are case-insensitive: drop any stale entry for this name in any spelling
          val states = placed.metadata.sheetStates.filterNot((name, _) =>
            name.value.equalsIgnoreCase(sheet.name.value)
          ) + (sheet.name -> state)
          Right(placed.copy(metadata = placed.metadata.copy(sheetStates = states)))

  private def refuse(sheet: Sheet, reason: String): XLResult[Workbook] =
    Left(XLError.InvalidWorkbook(s"Cannot adopt sheet '${sheet.name.value}': $reason"))
