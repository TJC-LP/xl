package com.tjclp.xl.cli.commands

import java.nio.file.Path

import cats.effect.IO
import cats.implicits.*
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.error.XLError
import com.tjclp.xl.cli.contract.{CliError, CliException, Location}
import com.tjclp.xl.cli.helpers.Resolve
import com.tjclp.xl.cli.MemoryGuard
import com.tjclp.xl.cli.output.Format
import com.tjclp.xl.formula.eval.SheetRenamer
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * Sheet management command handlers.
 *
 * Commands for add, remove, rename, move, copy sheet operations. All methods accept a `stream`
 * parameter to use the SAX/StAX workbook writer.
 *
 * Every refusal raised in THIS object is typed (ADR-017 §2.3): a domain error keeps its own
 * `XLError.code`, a wrong flag combination is `USAGE`, and none of them falls through to
 * `INTERNAL`, the code reserved for defects (GH-608). Other command objects (`BatchParser`'s
 * remaining ops, `StreamingWriteCommands`, `ImportCommands`, …) still raise plain exceptions;
 * converting them is the follow-up. The `private[cli]` helpers are shared with the batch
 * `add-sheet`/`rename-sheet` ops (`BatchParser`), so an op fails with the same code, hint and
 * location as the verb, wrapped by `BATCH_OP_FAILED`.
 */
object SheetCommands:

  // --- Typed failures (ADR-017 §2.3): every refusal names its code ------------------------------

  /** A domain error as raised, with its own code (exit 3). */
  private[cli] def domain(error: XLError): CliException =
    CliException(CliError.fromXLError(error, None))

  /** A domain error with the verb's own wording in place of the domain message. */
  private def domain(error: XLError, message: String): CliException =
    CliException(CliError.fromXLError(error, None).copy(message = message))

  /** A wrong flag or argument combination: `USAGE` (exit 2). */
  private def usage(message: String): CliException = CliException(CliError.usage(message, None))

  /** `INVALID_SHEET_NAME` (exit 3) with the validator's own sentence. */
  private[cli] def sheetName(name: String): IO[SheetName] =
    IO.fromEither(
      SheetName(name).left.map(reason => domain(XLError.InvalidSheetName(name, reason), reason))
    )

  /**
   * `SHEET_NOT_FOUND` (exit 3): the CLI's ONE text for an unknown sheet (`Resolve.sheetNotFound`,
   * `Sheet not found: X. Available: …`, the one every read verb and the goldens carry), with the
   * domain hint and the nearest names as candidates.
   */
  private[cli] def sheetNotFound(name: String, wb: Workbook): CliException =
    CliException(Resolve.sheetNotFound(wb.sheetNames.map(_.value).toVector, name))

  /**
   * The one wording of a `move-sheet` with no position: the parser's `USAGE` and the guard here.
   */
  private[cli] val MoveSheetPositionRequired: String =
    "move-sheet requires --to, --after, or --before option"

  /**
   * Whether the book already has a sheet called `name` the way Excel compares sheet names —
   * case-insensitively. `add-sheet data` beside `Data` and `rename-sheet T s` beside `S` would each
   * write two tabs Excel treats as one name (and a rewritten `=s!A1` would resolve against `S`).
   */
  private def hasSheetNamed(wb: Workbook, name: SheetName): Boolean =
    wb.sheets.exists(_.name.value.equalsIgnoreCase(name.value))

  /** `DUPLICATE_SHEET` (exit 3) with the verb's message; nothing has been written when it fires. */
  private[cli] def duplicateSheet(name: String, message: String): CliException =
    CliException(
      CliError
        .fromXLError(XLError.DuplicateSheet(name), None)
        .copy(
          message = message,
          hint = Some("sheet names are case-insensitive in Excel; choose a name no other sheet has")
        )
    )

  /**
   * A text the renamer cannot rewrite (GH-608): the domain error's own code (`FORMULA_ERROR` for a
   * formula the parser rejects, exit 3), `location` naming the cell or the sheet whose conditional
   * format or data validation refused, and a hint saying what to change. The message already names
   * the site and carries the parser's diagnostic. Nothing has been written when it fires.
   */
  private[cli] def unrewritable(site: Option[SheetRenamer.Site], error: XLError): CliException =
    import SheetRenamer.Site
    val location = site.flatMap {
      case Site.Cell(sheet, ref) => Some(Location(None, Some(sheet.value), Some(ref.toA1), None))
      case Site.ConditionalFormat(sheet) => Some(Location(None, Some(sheet.value), None, None))
      case Site.DataValidation(sheet) => Some(Location(None, Some(sheet.value), None, None))
      case Site.Name(_) => None
    }
    // sheet names are spelled as a formula would (quoted only when needed), like `Site.describe`
    def on(sheet: SheetName): String = s"on sheet ${SheetName.quoteForFormula(sheet.value)}"
    val hint = site.map {
      case s: Site.Cell => s"fix or replace the formula at ${s.describe} before renaming"
      case Site.ConditionalFormat(sheet) =>
        s"fix or remove the conditional format ${on(sheet)} that names the renamed sheet before renaming"
      case Site.DataValidation(sheet) =>
        s"fix or remove the data validation ${on(sheet)} that names the renamed sheet before renaming"
      case Site.Name(name) => s"fix or remove the defined name '$name' before renaming"
    }
    val base = CliError.fromXLError(error, location)
    CliException(base.copy(hint = hint.orElse(base.hint)))

  /**
   * Validate both names and rename `oldName` to `newName` with every reference rewritten
   * (`SheetRenamer.renameLocated`), every refusal typed: `INVALID_SHEET_NAME`, `SHEET_NOT_FOUND`,
   * `DUPLICATE_SHEET`, or [[unrewritable]]. The one rename the verb and the batch op share, so both
   * fail with the same code, hint and location; nothing has been written when it fails.
   */
  private[cli] def renamed(wb: Workbook, oldName: String, newName: String): IO[Workbook] =
    for
      oldSheetName <- sheetName(oldName)
      newSheetName <- sheetName(newName)
      updated <- IO.fromEither(
        SheetRenamer.renameLocated(wb, oldSheetName, newSheetName).left.map {
          // `site = None` MEANS workbook-level: only `Workbook.rename`'s own refusals arrive without
          // one. Pinning `None` keeps a located refusal (whatever its error) at its site.
          case SheetRenamer.Refusal(None, XLError.SheetNotFound(_, _)) => sheetNotFound(oldName, wb)
          case SheetRenamer.Refusal(None, XLError.DuplicateSheet(_)) =>
            duplicateSheet(newName, s"Sheet '$newName' already exists")
          case SheetRenamer.Refusal(site, error) => unrewritable(site, error)
        }
      )
    yield updated

  /** Write workbook using the standard or SAX/StAX backend based on mode */
  private def writeWorkbook(
    wb: Workbook,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean
  ): IO[Unit] =
    val excel = MemoryGuard.writer
    if stream then excel.writeWorkbookStream(wb, outputPath, config)
    else excel.writeWith(wb, outputPath, config)

  /**
   * Add a new empty sheet to workbook.
   *
   * @param stream
   *   If true, uses the SAX/StAX workbook writer
   */
  def addSheet(
    wb: Workbook,
    name: String,
    afterOpt: Option[String],
    beforeOpt: Option[String],
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      newName <- sheetName(name)
      _ <- IO
        .raiseError(
          duplicateSheet(
            name,
            s"Sheet '$name' already exists. Available: ${wb.sheetNames.map(_.value).mkString(", ")}"
          )
        )
        .whenA(hasSheetNamed(wb, newName))
      newSheet = Sheet(newName)
      updatedWb <- (afterOpt, beforeOpt) match
        case (Some(after), _) =>
          // Insert after specified sheet
          for
            afterName <- sheetName(after)
            idx = wb.sheets.indexWhere(_.name == afterName)
            _ <- IO.raiseError(sheetNotFound(after, wb)).whenA(idx < 0)
            result <- IO.fromEither(wb.insertAt(idx + 1, newSheet).left.map(domain))
          yield result
        case (_, Some(before)) =>
          // Insert before specified sheet
          for
            beforeName <- sheetName(before)
            idx = wb.sheets.indexWhere(_.name == beforeName)
            _ <- IO.raiseError(sheetNotFound(before, wb)).whenA(idx < 0)
            result <- IO.fromEither(wb.insertAt(idx, newSheet).left.map(domain))
          yield result
        case (None, None) =>
          // Append at end (use put for simplicity - adds if not exists)
          IO.pure(wb.put(newSheet))
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
      position = (afterOpt, beforeOpt) match
        case (Some(after), _) => s" (after '$after')"
        case (_, Some(before)) => s" (before '$before')"
        case _ => " (at end)"
    yield s"Added sheet: $name$position\n${Format.saveSuffix(outputPath, stream)}"

  /**
   * Remove sheet from workbook.
   *
   * @param stream
   *   If true, uses the SAX/StAX workbook writer
   */
  def removeSheet(
    wb: Workbook,
    name: String,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      target <- sheetName(name)
      updatedWb <- IO.fromEither(wb.remove(target).left.map {
        case XLError.SheetNotFound(_, _) => sheetNotFound(name, wb)
        case e @ XLError.InvalidWorkbook(reason) => domain(e, reason)
        case e => domain(e)
      })
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
    yield s"Removed sheet: $name\n${Format.saveSuffix(outputPath, stream)}"

  /**
   * Rename a sheet AND every reference to it (GH-559): `SheetRenamer.rename` rewrites the sheet
   * qualifier in cell formulas on every sheet, defined names, conditional-format and
   * data-validation formulas, preserving cached values (a rename changes no value). The summary
   * counts the cell formulas whose text changed. A dependent formula that mentions the sheet but
   * cannot be parsed refuses the whole rename before anything is written — `FORMULA_ERROR` with the
   * offending cell as `location` and the parser's diagnostic (GH-608), never `INTERNAL`.
   *
   * @param stream
   *   If true, uses the SAX/StAX workbook writer
   */
  def renameSheet(
    wb: Workbook,
    oldName: String,
    newName: String,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      updatedWb <- renamed(wb, oldName, newName)
      rewritten = rewrittenFormulaCount(wb, updatedWb)
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
      suffix = if rewritten > 0 then s"; $rewritten formula(s) rewritten" else ""
    yield s"Renamed: $oldName → $newName$suffix\n${Format.saveSuffix(outputPath, stream)}"

  /**
   * Cell formulas whose text a rename changed: sheets keep their positions under a rename, so the
   * two workbooks are compared position by position.
   */
  private def rewrittenFormulaCount(before: Workbook, after: Workbook): Int =
    before.sheets
      .zip(after.sheets)
      .map { (was, is) =>
        is.cells.count { (ref, cell) =>
          (cell.value, was.cells.get(ref).map(_.value)) match
            case (CellValue.Formula(after, _, _), Some(CellValue.Formula(before, _, _))) =>
              after != before
            case _ => false
        }
      }
      .sum

  /**
   * Move sheet to new position.
   *
   * @param stream
   *   If true, uses the SAX/StAX workbook writer
   */
  def moveSheet(
    wb: Workbook,
    name: String,
    toIndexOpt: Option[Int],
    afterOpt: Option[String],
    beforeOpt: Option[String],
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      target <- sheetName(name)
      currentIdx = wb.sheets.indexWhere(_.name == target)
      _ <- IO.raiseError(sheetNotFound(name, wb)).whenA(currentIdx < 0)
      targetIdx <- (toIndexOpt, afterOpt, beforeOpt) match
        case (Some(idx), _, _) => IO.pure(idx)
        case (_, Some(after), _) =>
          for
            afterName <- sheetName(after)
            afterIdx = wb.sheets.indexWhere(_.name == afterName)
            _ <- IO.raiseError(sheetNotFound(after, wb)).whenA(afterIdx < 0)
          yield afterIdx + 1
        case (_, _, Some(before)) =>
          for
            beforeName <- sheetName(before)
            beforeIdx = wb.sheets.indexWhere(_.name == beforeName)
            _ <- IO.raiseError(sheetNotFound(before, wb)).whenA(beforeIdx < 0)
          yield beforeIdx
        case (None, None, None) =>
          // unreachable from the CLI (the parser validates first, GH-608); the library-API guard
          IO.raiseError(usage(MoveSheetPositionRequired))
      // Build new order: remove sheet from current position, insert at target
      currentNames = wb.sheetNames.toVector
      withoutSheet = currentNames.patch(currentIdx, Nil, 1)
      // Adjust target index if we removed from before the target
      adjustedIdx = if currentIdx < targetIdx then targetIdx - 1 else targetIdx
      clampedIdx = adjustedIdx.max(0).min(withoutSheet.size)
      newOrder = withoutSheet.patch(clampedIdx, Vector(target), 0)
      updatedWb <- IO.fromEither(wb.reorder(newOrder).left.map(domain))
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
      position = s"to position $clampedIdx"
    yield s"Moved: $name $position\n${Format.saveSuffix(outputPath, stream)}"

  /**
   * Copy sheet to new name.
   *
   * @param stream
   *   If true, uses the SAX/StAX workbook writer
   */
  def copySheet(
    wb: Workbook,
    sourceName: String,
    targetName: String,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      sourceSheetName <- sheetName(sourceName)
      targetSheetName <- sheetName(targetName)
      sourceSheet <- IO.fromOption(wb.sheets.find(_.name == sourceSheetName))(
        sheetNotFound(sourceName, wb)
      )
      _ <- IO
        .raiseError(duplicateSheet(targetName, s"Sheet '$targetName' already exists"))
        .whenA(hasSheetNamed(wb, targetSheetName))
      copiedSheet = sourceSheet.copy(name = targetSheetName)
      updatedWb = wb.put(copiedSheet)
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
    yield s"Copied: $sourceName → $targetName\n${Format.saveSuffix(outputPath, stream)}"

  /**
   * Hide a sheet from the sheet tabs.
   *
   * @param veryHide
   *   If true, uses "veryHidden" state (not accessible from Excel UI, only via VBA)
   * @param stream
   *   If true, uses the SAX/StAX workbook writer
   */
  def hideSheet(
    wb: Workbook,
    name: String,
    veryHide: Boolean,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      target <- sheetName(name)
      state = if veryHide then Some("veryHidden") else Some("hidden")
      updatedWb <- IO.fromEither(wb.setSheetState(target, state).left.map {
        case XLError.SheetNotFound(_, _) => sheetNotFound(name, wb)
        case e @ XLError.InvalidWorkbook(reason) => domain(e, reason)
        case e => domain(e)
      })
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
      stateDesc = if veryHide then "very hidden" else "hidden"
    yield s"Sheet '$name' is now $stateDesc\n${Format.saveSuffix(outputPath, stream)}"

  /** Add or replace a workbook-scoped named range, then write (GH-236). */
  def nameAdd(
    wb: Workbook,
    name: String,
    refersTo: String,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    val updated = wb.withDefinedName(name, refersTo)
    writeWorkbook(updated, outputPath, config, stream)
      .as(s"Added named range '$name' -> $refersTo\n${Format.saveSuffix(outputPath, stream)}")

  /** Remove a workbook-scoped named range, then write (GH-236). */
  def nameRemove(
    wb: Workbook,
    name: String,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    if !wb.metadata.definedNames.exists(d => d.name == name && d.localSheetId.isEmpty) then
      // a data condition with prose alone: `OTHER` (exit 3), as WriteCommands.refused
      IO.raiseError(domain(XLError.Other(s"Named range '$name' not found")))
    else
      val updated = wb.removeDefinedName(name)
      writeWorkbook(updated, outputPath, config, stream)
        .as(s"Removed named range '$name'\n${Format.saveSuffix(outputPath, stream)}")

  /**
   * Show a hidden sheet (make it visible).
   *
   * @param stream
   *   If true, uses the SAX/StAX workbook writer
   */
  def showSheet(
    wb: Workbook,
    name: String,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      target <- sheetName(name)
      updatedWb <- IO.fromEither(wb.setSheetState(target, None).left.map {
        case XLError.SheetNotFound(_, _) => sheetNotFound(name, wb)
        case e => domain(e)
      })
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
    yield s"Sheet '$name' is now visible\n${Format.saveSuffix(outputPath, stream)}"
