package com.tjclp.xl.cli

import java.nio.file.Path

import cats.effect.{IO, Ref}

import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.error.{XLError, XLResult}

/**
 * Session state for the xl CLI (Phase 2 - REPL mode).
 *
 * NOTE: This module is scaffolding for future REPL mode implementation. The current CLI (Phase 1)
 * is stateless - each command specifies --file/--sheet. This will be used when we add interactive
 * session support with commands like `open`, `select`, `save`, etc.
 *
 * Tracks the currently open workbook, active sheet context, and whether there are unsaved changes.
 */
final case class Session(
  workbook: Option[Workbook],
  path: Option[Path],
  activeSheet: Option[SheetName],
  isDirty: Boolean,
  isReadOnly: Boolean
):
  /** Whether a workbook is currently loaded */
  def isOpen: Boolean = workbook.isDefined

  /** Get the active sheet, or the first sheet if none selected */
  def currentSheet: Option[Sheet] =
    workbook.flatMap { wb =>
      activeSheet
        .flatMap(name => wb.sheets.find(_.name == name))
        .orElse(wb.sheets.headOption)
    }

  /** Get sheet by name */
  def sheet(name: SheetName): Option[Sheet] =
    workbook.flatMap(_.sheets.find(_.name == name))

  /** Update the workbook, marking session as dirty */
  def withWorkbook(wb: Workbook): Session =
    copy(workbook = Some(wb), isDirty = true)

  /** Update a specific sheet in the workbook */
  def updateSheet(sheet: Sheet): Session =
    workbook match
      case Some(wb) =>
        val updated =
          wb.copy(sheets = wb.sheets.map(s => if s.name == sheet.name then sheet else s))
        copy(workbook = Some(updated), isDirty = true)
      case None => this

  /** Mark session as clean (after save) */
  def markClean: Session = copy(isDirty = false)

  /** Select a different active sheet */
  def selectSheet(name: SheetName): Session =
    copy(activeSheet = Some(name))

object Session:
  /** Empty session with no workbook loaded */
  def empty: Session = Session(
    workbook = None,
    path = None,
    activeSheet = None,
    isDirty = false,
    isReadOnly = false
  )

  /** Create session with a loaded workbook */
  def withWorkbook(wb: Workbook, path: Path, readOnly: Boolean = false): Session =
    Session(
      workbook = Some(wb),
      path = Some(path),
      activeSheet = wb.sheets.headOption.map(_.name),
      isDirty = false,
      isReadOnly = readOnly
    )

  /** Create session with a new (unsaved) workbook */
  def newWorkbook(wb: Workbook): Session =
    Session(
      workbook = Some(wb),
      path = None,
      activeSheet = wb.sheets.headOption.map(_.name),
      isDirty = true, // New workbooks need to be saved
      isReadOnly = false
    )

/**
 * Mutable session state holder for the CLI.
 *
 * Uses cats-effect Ref for thread-safe state management.
 */
final class SessionRef private (ref: Ref[IO, Session]):
  def get: IO[Session] = ref.get

  def modify[A](f: Session => (Session, A)): IO[A] = ref.modify(f)

  def update(f: Session => Session): IO[Unit] = ref.update(f)

  def set(s: Session): IO[Unit] = ref.set(s)

object SessionRef:
  def empty: IO[SessionRef] =
    Ref.of[IO, Session](Session.empty).map(new SessionRef(_))

  def apply(session: Session): IO[SessionRef] =
    Ref.of[IO, Session](session).map(new SessionRef(_))
