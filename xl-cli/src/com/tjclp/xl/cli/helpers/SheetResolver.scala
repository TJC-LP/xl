package com.tjclp.xl.cli.helpers

import cats.effect.IO
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, RefType, SheetName}
import com.tjclp.xl.cli.contract.{CliError, CliException}
import com.tjclp.xl.text.Suggest

/**
 * Sheet resolution utilities for CLI commands.
 *
 * Provides helpers for finding sheets by name, resolving qualified references, and validating sheet
 * context requirements. Every failure is a [[CliException]] carrying the domain error's code
 * (ADR-017 §2.3): `SHEET_NOT_FOUND` with "did you mean" candidates, `SHEET_REQUIRED` with the
 * available sheets, `INVALID_SHEET_NAME`, `INVALID_REFERENCE`. The message texts are the ones the
 * CLI has always printed.
 */
object SheetResolver:

  private def sheetNames(wb: Workbook): Vector[String] = wb.sheets.map(_.name.value)

  /** `SHEET_NOT_FOUND`, ranking the workbook's sheet names as candidates. */
  private def notFound(wb: Workbook, name: String): CliException =
    val names = sheetNames(wb)
    CliException(
      CliError
        .fromXLError(XLError.SheetNotFound(name), None)
        .copy(
          message = s"Sheet not found: $name. Available: ${names.mkString(", ")}",
          candidates = Suggest.closest(name, names)
        )
    )

  /**
   * Resolve an optional sheet name to a Sheet.
   *
   * @param wb
   *   Workbook to search
   * @param sheetNameOpt
   *   Optional sheet name (if None, returns None)
   * @return
   *   IO containing Some(Sheet) if found, None if not specified
   */
  def resolveSheet(wb: Workbook, sheetNameOpt: Option[String]): IO[Option[Sheet]] =
    sheetNameOpt match
      case Some(name) =>
        IO.fromEither(
          SheetName
            .apply(name)
            .left
            .map(reason =>
              CliException(
                CliError
                  .fromXLError(XLError.InvalidSheetName(name, reason), None)
                  .copy(message = reason)
              )
            )
        ).flatMap { sheetName =>
          IO.fromOption(wb.sheets.find(_.name == sheetName))(notFound(wb, name)).map(Some(_))
        }
      case None =>
        IO.pure(None)

  /**
   * Require a sheet to be present.
   *
   * @param wb
   *   Workbook for error message
   * @param sheetOpt
   *   Optional sheet
   * @param context
   *   Command name for error message
   * @return
   *   IO containing the sheet, or a `SHEET_REQUIRED` error naming the available sheets
   */
  def requireSheet(wb: Workbook, sheetOpt: Option[Sheet], context: String): IO[Sheet] =
    IO.fromOption(sheetOpt) {
      val names = sheetNames(wb)
      CliException(
        CliError
          .fromXLError(XLError.SheetRequired(context, names), None)
          .copy(
            message = s"$context requires --sheet or qualified ref (e.g., Sheet1!A1). " +
              s"Available sheets: ${names.mkString(", ")}"
          )
      )
    }

  /**
   * Find a sheet by name.
   *
   * @param wb
   *   Workbook to search
   * @param name
   *   Sheet name
   * @return
   *   IO containing the sheet, or a `SHEET_NOT_FOUND` error if not found
   */
  def findSheet(wb: Workbook, name: SheetName): IO[Sheet] =
    IO.fromOption(wb.sheets.find(_.name == name))(notFound(wb, name.value))

  /**
   * Resolve a reference string to a (Sheet, Either[ARef, CellRange]).
   *
   * Supports qualified refs like `Sheet1!A1` which override the default sheet context. For
   * unqualified refs, requires a default sheet or fails with helpful error.
   *
   * @param wb
   *   Workbook to search
   * @param defaultSheetOpt
   *   Default sheet context
   * @param refStr
   *   Reference string (e.g., "A1", "A1:B10", "Sheet1!A1")
   * @param context
   *   Command name for error messages
   * @return
   *   IO containing (Sheet, Either[ARef, CellRange])
   */
  def resolveRef(
    wb: Workbook,
    defaultSheetOpt: Option[Sheet],
    refStr: String,
    context: String
  ): IO[(Sheet, Either[ARef, CellRange])] =
    IO.fromEither(
      RefType
        .parse(refStr)
        .left
        .map(reason =>
          CliException(
            CliError.fromXLError(XLError.InvalidReference(reason), None).copy(message = reason)
          )
        )
    ).flatMap {
      case RefType.Cell(ref) =>
        requireSheet(wb, defaultSheetOpt, s"$context with unqualified ref '$refStr'")
          .map(s => (s, Left(ref)))
      case RefType.Range(range) =>
        requireSheet(wb, defaultSheetOpt, s"$context with unqualified range '$refStr'")
          .map(s => (s, Right(range)))
      case RefType.QualifiedCell(sheetName, ref) =>
        findSheet(wb, sheetName).map(s => (s, Left(ref)))
      case RefType.QualifiedRange(sheetName, range) =>
        findSheet(wb, sheetName).map(s => (s, Right(range)))
    }
