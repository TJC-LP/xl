package com.tjclp.xl.cli.helpers

import com.tjclp.xl.addressing.Column
import com.tjclp.xl.sheets.Sheet

/**
 * Forwarder onto `Sheet.autoFitWidth` (ADR-017 §2.12, W2.1: the GH-156 font-metric width — Calibri
 * 11 max-digit-width convention, char-count heuristic when fonts are unavailable — lives in xl-core
 * now), shared by `col --auto-fit`, `autofit` and batch `autofit` so the three cannot drift.
 */
object ColumnAutoFit:

  /** The width `col` needs for its formatted content, in Excel character units. */
  def calculateWidth(sheet: Sheet, col: Column): Double = sheet.autoFitWidth(col)
