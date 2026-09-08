package com.tjclp.xl.cli.helpers

import cats.syntax.all.*

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, RefType, SheetName}
import com.tjclp.xl.cli.contract.{CliError, ErrorCode, Location, Warning, WarningCode}
import com.tjclp.xl.ooxml.metadata.LightMetadata

/**
 * THE sheet rule (ADR-017 §2.5), stated once and used by every verb, batch op and streaming path:
 *
 *   1. a sheet-qualified ref (`'Q1 Report'!A1:D9`) names the sheet;
 *   2. otherwise `-s`/`--sheet` (for a batch op, its `sheet` key comes before `-s`);
 *   3. otherwise, if the workbook has exactly one sheet, that sheet — the runner announces it as a
 *      `SHEET_AUTOSELECTED` warning under `--json` only ([[autoSelected]]);
 *   4. otherwise `SHEET_REQUIRED` (exit 3 — the file has been read) with the sheet names as
 *      candidates.
 *
 * Pure and total: every entry point returns `Either[CliError, _]`; the IO forwarders live in
 * [[SheetResolver]]. The two message texts the CLI has always printed for `SHEET_REQUIRED` and
 * `SHEET_NOT_FOUND` are kept verbatim ([[sheetRequired]], [[sheetNotFound]]).
 */
object Resolve:

  /** What a ref string addressed, once its sheet is known. */
  enum Target derives CanEqual:
    case Cell(ref: ARef)
    case Range(range: CellRange)

    /** The shape the handlers pattern-match on. */
    def toEither: Either[ARef, CellRange] = this match
      case Cell(ref) => Left(ref)
      case Range(range) => Right(range)

  /**
   * A resolved ref: which step decided it — `viaQualifiedRef` for step 1, `autoSelected` for step 3
   * — so a caller can tell an explicit choice from a default.
   */
  final case class Resolved(
    sheet: Sheet,
    target: Target,
    viaQualifiedRef: Boolean,
    autoSelected: Boolean
  ) derives CanEqual

  // ---------------------------------------------------------------------------------------------
  // The rule over a loaded workbook
  // ---------------------------------------------------------------------------------------------

  /** Steps 1–4 for a ref string; `verb` names the caller in the `SHEET_REQUIRED` message. */
  def target(
    wb: Workbook,
    sheetFlag: Option[String],
    refStr: String,
    verb: String
  ): Either[CliError, Resolved] =
    ref(refStr).flatMap {
      case (Some(qualifier), tgt) => named(wb, qualifier).map(Resolved(_, tgt, true, false))
      case (None, tgt) =>
        sheet(wb, sheetFlag, unqualified(verb, refStr, tgt))
          .map(Resolved(_, tgt, false, sheetFlag.isEmpty))
    }

  /**
   * [[target]] with the `-s` sheet already looked up (the shape every handler receives): step 2 is
   * the given default, steps 3–4 apply only without one.
   */
  def targetWith(
    wb: Workbook,
    default: Option[Sheet],
    refStr: String,
    verb: String
  ): Either[CliError, Resolved] =
    ref(refStr).flatMap {
      case (Some(qualifier), tgt) => named(wb, qualifier).map(Resolved(_, tgt, true, false))
      case (None, tgt) =>
        default match
          case Some(s) => Right(Resolved(s, tgt, false, false))
          case None =>
            sheet(wb, None, unqualified(verb, refStr, tgt)).map(Resolved(_, tgt, false, true))
    }

  /** Steps 2–4 for a verb that needs a sheet but takes no ref (`bounds`, `row`, `unfreeze`, …). */
  def sheet(wb: Workbook, sheetFlag: Option[String], verb: String): Either[CliError, Sheet] =
    sheetFlag match
      case Some(name) => lookup(wb, name)
      case None =>
        onlySheet(wb).toRight(sheetRequired(verb, wb.sheets.map(_.name.value)))

  /** Step 1's lookup on its own: the sheet a qualifier names, else `SHEET_NOT_FOUND`. */
  def named(wb: Workbook, name: SheetName): Either[CliError, Sheet] =
    wb.sheets.find(_.name == name).toRight(sheetNotFound(wb.sheets.map(_.name.value), name.value))

  /**
   * The run's default sheet, resolved once before dispatch: `-s` by name (`INVALID_SHEET_NAME` /
   * `SHEET_NOT_FOUND`); else, for a verb that takes a sheet, the only sheet of a single-sheet book
   * (step 3); else `None` — an AllSheets verb reads the whole book, a sheet verb reaches step 4
   * through [[targetWith]] or [[sheet]] when it needs one.
   */
  def default(
    wb: Workbook,
    sheetFlag: Option[String],
    takesSheet: Boolean
  ): Either[CliError, Option[Sheet]] =
    sheetFlag match
      case Some(name) => lookup(wb, name).map(Some(_))
      case None => Right(if takesSheet then onlySheet(wb) else None)

  /** Step 3's premise: the only sheet of a single-sheet book. */
  def onlySheet(wb: Workbook): Option[Sheet] = wb.sheets match
    case Vector(single) => Some(single)
    case _ => None

  // ---------------------------------------------------------------------------------------------
  // The rule over metadata (the streaming paths: a name, no Sheet in memory)
  // ---------------------------------------------------------------------------------------------

  /** Steps 1–4 over `workbook.xml` alone; `qualified` is the sheet the ref string named, if any. */
  def sheetName(
    meta: LightMetadata,
    sheetFlag: Option[String],
    qualified: Option[SheetName],
    verb: String
  ): Either[CliError, SheetName] =
    sheetNameAmong(meta.sheets.map(_.name), sheetFlag, qualified, verb)

  /**
   * Steps 1–4 over the sheet names alone — the form every [[com.tjclp.xl.cli.read.SheetSource]]
   * resolves through, so a loaded workbook and the streaming reader apply one rule with one text.
   */
  def sheetNameAmong(
    names: Vector[SheetName],
    sheetFlag: Option[String],
    qualified: Option[SheetName],
    verb: String
  ): Either[CliError, SheetName] =
    def known(name: SheetName): Either[CliError, SheetName] =
      if names.contains(name) then Right(name)
      else Left(sheetNotFound(names.map(_.value), name.value))
    qualified match
      case Some(qualifier) => known(qualifier)
      case None =>
        sheetFlag match
          case Some(flag) => validSheetName(flag).flatMap(known)
          case None => onlyAmong(names).toRight(sheetRequired(verb, names.map(_.value)))

  /**
   * A name given on the command line (`-s`, one of `--sheets`) over metadata: `INVALID_SHEET_NAME`
   * when the validator refuses it, `SHEET_NOT_FOUND` with the nearest names when the book lacks it.
   */
  def knownName(meta: LightMetadata, name: String): Either[CliError, SheetName] =
    knownAmong(meta.sheets.map(_.name), name)

  /** [[knownName]] over the sheet names alone. */
  def knownAmong(names: Vector[SheetName], name: String): Either[CliError, SheetName] =
    validSheetName(name).flatMap { sn =>
      if names.contains(sn) then Right(sn) else Left(sheetNotFound(names.map(_.value), sn.value))
    }

  /** Step 3's premise over metadata. */
  def only(meta: LightMetadata): Option[SheetName] = onlyAmong(meta.sheets.map(_.name))

  /** Step 3's premise over the sheet names alone. */
  def onlyAmong(names: Vector[SheetName]): Option[SheetName] = names match
    case Vector(single) => Some(single)
    case _ => None

  // ---------------------------------------------------------------------------------------------
  // The rule for one batch op
  // ---------------------------------------------------------------------------------------------

  /**
   * Steps 1–4 for one batch op: qualified ref > the op's `sheet` key > the batch default (`-s`, as
   * retargeted by an earlier `rename-sheet`) > the only sheet. A qualified ref that disagrees with
   * `sheet` is a malformed op (`BATCH_OP_INVALID`, exit 2), so is an invalid `sheet` value; the
   * `SHEET_REQUIRED` of step 4 carries the op's 1-based `index` as its location.
   */
  def forOp(
    wb: Workbook,
    defaultSheet: Option[SheetName],
    opSheet: Option[String],
    qualified: Option[SheetName],
    index: Int,
    op: String
  ): Either[CliError, SheetName] =
    for
      declared <- opSheet.traverse { name =>
        SheetName(name).left.map(reason =>
          invalidOp(index, op, s"invalid 'sheet' '$name': $reason")
        )
      }
      _ <- opSheetAgreement(index, op, "the qualified ref", qualified, declared)
      name <- qualified.orElse(declared).orElse(defaultSheet) match
        case Some(chosen) => Right(chosen)
        case None =>
          onlySheet(wb)
            .map(_.name)
            .toRight(
              sheetRequired(s"batch $op", wb.sheets.map(_.name.value)).copy(location = at(index))
            )
    yield name

  /**
   * A target ref qualified with a sheet other than the op's `sheet` key is a contradiction —
   * `BATCH_OP_INVALID` naming the op and its index; `subject` says which ref (`ref 'Other!A1'`).
   */
  def opSheetAgreement(
    index: Int,
    op: String,
    subject: String,
    qualified: Option[SheetName],
    declared: Option[SheetName]
  ): Either[CliError, Unit] =
    (qualified, declared) match
      case (Some(named), Some(sheet)) if named != sheet =>
        Left(
          invalidOp(
            index,
            op,
            s"$subject names sheet '${named.value}' but \"sheet\" is '${sheet.value}'"
          )
        )
      case _ => Right(())

  // ---------------------------------------------------------------------------------------------
  // Diagnostics
  // ---------------------------------------------------------------------------------------------

  /** Step 3 as the run reports it under `--json`: which sheet stood in for the missing `-s`. */
  def autoSelected(name: SheetName): Warning =
    Warning(
      WarningCode.SHEET_AUTOSELECTED,
      s"no -s given: '${name.value}' is the workbook's only sheet, so it was selected",
      Some(Location.none.copy(sheet = Some(name.value)))
    )

  /** `SHEET_REQUIRED` with the text the CLI has always printed; candidates are the sheet names. */
  def sheetRequired(context: String, names: Vector[String]): CliError =
    CliError
      .fromXLError(XLError.SheetRequired(context, names), None)
      .copy(
        message = s"$context requires --sheet or qualified ref (e.g., Sheet1!A1). " +
          s"Available sheets: ${names.mkString(", ")}"
      )

  /**
   * `SHEET_NOT_FOUND` with the text the CLI has always printed; the nearest names are the domain
   * error's own `candidates` (`SheetNotFound(name, available)`, GH-589), so `xl` and a script's
   * `orExit` offer the same "did you mean".
   */
  def sheetNotFound(names: Vector[String], name: String): CliError =
    CliError
      .fromXLError(XLError.SheetNotFound(name, names), None)
      .copy(message = s"Sheet not found: $name. Available: ${names.mkString(", ")}")

  /**
   * A ref of the wrong shape for the verb — a range where one cell is needed, a column that is not
   * one: `INVALID_REFERENCE` as the domain renders it (`Invalid reference: <reason>`), the text
   * `deps` has always used.
   */
  def invalidReference(reason: String): CliError =
    CliError.fromXLError(XLError.InvalidReference(reason), None)

  private def at(index: Int): Option[Location] = Some(Location.none.copy(opIndex = Some(index)))

  private def invalidOp(index: Int, op: String, message: String): CliError =
    CliError(ErrorCode.BATCH_OP_INVALID, s"Object $index ($op): $message", location = at(index))

  /** A `-s` value as a sheet name: `INVALID_SHEET_NAME` with the validator's own text. */
  def validSheetName(name: String): Either[CliError, SheetName] =
    SheetName(name).left.map(reason =>
      CliError.fromXLError(XLError.InvalidSheetName(name, reason), None).copy(message = reason)
    )

  private def lookup(wb: Workbook, name: String): Either[CliError, Sheet] =
    validSheetName(name).flatMap(named(wb, _))

  /**
   * `INVALID_REFERENCE` with the parser's own text, else the qualifier (if any) and the target. A
   * whole-column (`AM:AM`, `A:C`) or whole-row (`3:3`, `1:5`) span, bare or sheet-qualified, is the
   * range over every row (or column) of the sheet (GH-641), the form `view` and `stats` take over a
   * column of unknown height; the source streams it lazily.
   */
  def ref(refStr: String): Either[CliError, (Option[SheetName], Target)] =
    RefType.parse(refStr) match
      case Right(RefType.Cell(ref)) => Right((None, Target.Cell(ref)))
      case Right(RefType.Range(range)) => Right((None, Target.Range(range)))
      case Right(RefType.QualifiedCell(sheet, ref)) => Right((Some(sheet), Target.Cell(ref)))
      case Right(RefType.QualifiedRange(sheet, range)) => Right((Some(sheet), Target.Range(range)))
      case Left(reason) =>
        wholeSpan(refStr).toRight(
          CliError.fromXLError(XLError.InvalidReference(reason), None).copy(message = reason)
        )

  /** `[Sheet!]A:C` or `[Sheet!]1:5` as the full-height (full-width) range, else `None`. */
  private def wholeSpan(refStr: String): Option[(Option[SheetName], Target)] =
    val bang = refStr.lastIndexOf('!')
    val (qualifier, spanPart) =
      if bang < 0 then (None, refStr)
      else (Some(refStr.substring(0, bang)), refStr.substring(bang + 1))
    val sheet: Option[Option[SheetName]] = qualifier match
      case None => Some(None)
      case Some(quoted) if quoted.length >= 2 && quoted.startsWith("'") && quoted.endsWith("'") =>
        SheetName(quoted.substring(1, quoted.length - 1).replace("''", "'")).toOption.map(Some(_))
      case Some(bare) => SheetName(bare).toOption.map(Some(_))
    for
      s <- sheet
      range <- CellRange.parse(spanPart.trim).toOption
      if range.isFullColumn || range.isFullRow
    yield (s, Target.Range(range))

  /** `A:C` for a whole-column range, `1:5` for a whole-row one, `A1:B2` otherwise. */
  def rangeLabel(range: CellRange): String =
    if range.isFullColumn then s"${range.start.col.toLetter}:${range.end.col.toLetter}"
    else if range.isFullRow then s"${range.start.row.index1}:${range.end.row.index1}"
    else range.toA1

  /** The `SHEET_REQUIRED` context for an unqualified ref: today's exact wording. */
  def unqualified(verb: String, refStr: String, target: Target): String = target match
    case Target.Cell(_) => s"$verb with unqualified ref '$refStr'"
    case Target.Range(_) => s"$verb with unqualified range '$refStr'"
