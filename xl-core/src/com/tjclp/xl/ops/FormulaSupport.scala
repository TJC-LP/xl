package com.tjclp.xl.ops

import com.tjclp.xl.addressing.{Column, Row, SheetName}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.workbooks.Workbook

/**
 * The formula-aware work the core [[Edit]] interpreter delegates (ADR-017 §2.12): xl-core has no
 * parser, so shifting a dragged formula, rewriting references through a structural edit and
 * renaming a sheet inside formula text are capabilities a caller supplies. xl-evaluator's
 * `EvalFormulaSupport` implements every method; the scripting prelude provides it as the `given`.
 * There is deliberately no `given` default in xl-core: a silent formula-blind interpreter would
 * write books that open with `#REF!`.
 */
trait FormulaSupport:
  /** `Right(())` when `formula` (with or without its leading `=`) is acceptable to store. */
  def validate(formula: String): XLResult[Unit]

  /** Shift every relative reference by `(colDelta, rowDelta)` the way a fill-drag does. */
  def shift(formula: String, colDelta: Int, rowDelta: Int): XLResult[String]

  /** Insert `count` rows before `at` on `sheet`, rewriting references on every sheet. */
  def insertRows(wb: Workbook, sheet: SheetName, at: Row, count: Int): XLResult[Workbook]

  /** Delete `count` rows from `at` on `sheet`, rewriting references on every sheet. */
  def deleteRows(wb: Workbook, sheet: SheetName, at: Row, count: Int): XLResult[Workbook]

  /** Insert `count` columns before `at` on `sheet`, rewriting references on every sheet. */
  def insertCols(wb: Workbook, sheet: SheetName, at: Column, count: Int): XLResult[Workbook]

  /** Delete `count` columns from `at` on `sheet`, rewriting references on every sheet. */
  def deleteCols(wb: Workbook, sheet: SheetName, at: Column, count: Int): XLResult[Workbook]

  /** Rename `from` to `to`, rewriting every formula, defined name and rule that names it. */
  def renameSheet(wb: Workbook, from: SheetName, to: SheetName): XLResult[Workbook]

object FormulaSupport:

  private val Capability = "formula support"

  private val Hint =
    "provide xl-evaluator's EvalFormulaSupport (given FormulaSupport = EvalFormulaSupport) " +
      "or import com.tjclp.xl.scripting.{*, given}"

  /** The refusal every text-only capability returns, naming the edit that needed it. */
  def unsupported(op: String): XLError = XLError.UnsupportedCapability(op, Capability, Hint)

  /**
   * Formula text is stored as written and nothing is rewritten: `validate` accepts non-empty text;
   * `shift`, the structural edits and `renameSheet` REFUSE with [[XLError.UnsupportedCapability]] —
   * xl-core cannot rewrite what it cannot parse, and writing unshifted or stale references would be
   * a silent `#REF!`.
   */
  val textOnly: FormulaSupport = new FormulaSupport:
    def validate(formula: String): XLResult[Unit] =
      if formula.trim.stripPrefix("=").trim.isEmpty then
        Left(XLError.FormulaError(formula, "empty formula"))
      else Right(())

    def shift(formula: String, colDelta: Int, rowDelta: Int): XLResult[String] =
      Left(unsupported("shift"))

    def insertRows(wb: Workbook, sheet: SheetName, at: Row, count: Int): XLResult[Workbook] =
      Left(unsupported("insert-rows"))

    def deleteRows(wb: Workbook, sheet: SheetName, at: Row, count: Int): XLResult[Workbook] =
      Left(unsupported("delete-rows"))

    def insertCols(wb: Workbook, sheet: SheetName, at: Column, count: Int): XLResult[Workbook] =
      Left(unsupported("insert-cols"))

    def deleteCols(wb: Workbook, sheet: SheetName, at: Column, count: Int): XLResult[Workbook] =
      Left(unsupported("delete-cols"))

    def renameSheet(wb: Workbook, from: SheetName, to: SheetName): XLResult[Workbook] =
      Left(unsupported("rename-sheet"))
