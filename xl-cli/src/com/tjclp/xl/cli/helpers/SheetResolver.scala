package com.tjclp.xl.cli.helpers

import cats.effect.IO
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cli.contract.{CliError, CliException}

/**
 * The IO face of THE sheet rule ([[Resolve]], ADR-017 §2.5) with the signatures every handler
 * calls. Each method is a one-line forwarder: the rule lives in `Resolve`, and every failure is a
 * [[CliException]] carrying the domain error's code — `SHEET_NOT_FOUND` with "did you mean"
 * candidates, `SHEET_REQUIRED` with the available sheets, `INVALID_SHEET_NAME`, `INVALID_REFERENCE`
 * — with the message texts the CLI has always printed.
 */
object SheetResolver:

  private def lift[A](result: Either[CliError, A]): IO[A] =
    IO.fromEither(result.left.map(CliException(_)))

  /** The `-s` sheet by name, or `None` without one: step 2 alone (no auto-select). */
  def resolveSheet(wb: Workbook, sheetNameOpt: Option[String]): IO[Option[Sheet]] =
    lift(Resolve.default(wb, sheetNameOpt, takesSheet = false))

  /**
   * The given sheet, else steps 3–4: the only sheet of a single-sheet book, or a `SHEET_REQUIRED`
   * error naming the available sheets (`context` names the command).
   */
  def requireSheet(wb: Workbook, sheetOpt: Option[Sheet], context: String): IO[Sheet] =
    lift(sheetOpt.fold(Resolve.sheet(wb, None, context))(Right(_)))

  /** The sheet a name denotes, or a `SHEET_NOT_FOUND` error. */
  def findSheet(wb: Workbook, name: SheetName): IO[Sheet] =
    lift(Resolve.named(wb, name))

  /**
   * A reference string to its `(Sheet, Either[ARef, CellRange])` by THE rule: a qualified ref
   * (`Sheet1!A1`) names the sheet, else the default, else the only sheet of a single-sheet book,
   * else `SHEET_REQUIRED` (`context` names the command).
   */
  def resolveRef(
    wb: Workbook,
    defaultSheetOpt: Option[Sheet],
    refStr: String,
    context: String
  ): IO[(Sheet, Either[ARef, CellRange])] =
    lift(
      Resolve
        .targetWith(wb, defaultSheetOpt, refStr, context)
        .map(r => (r.sheet, r.target.toEither))
    )
