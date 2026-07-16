package com.tjclp.xl.workbooks

/**
 * Workbook calculation mode — OOXML ST_CalcMode, the `calcMode` attribute of `<calcPr>` (GH-400).
 *
 * `AutoNoTable` (OOXML `autoNoTable`) recalculates everything EXCEPT data tables — the standard
 * professional-model setting when sensitivity tables must hold their last computed values instead
 * of recomputing on every edit. The codec maps cases to the OOXML tokens `manual` / `auto` /
 * `autoNoTable`; unknown tokens parse to None (a read never fails on them).
 */
enum CalcMode derives CanEqual:
  case Manual, Auto, AutoNoTable

/**
 * Workbook calculation properties: the modeled subset of `<calcPr>` (GH-373, GH-400).
 *
 * Two attribute families with distinct write semantics:
 *
 *   - The iterative-calculation triple (GH-373) — `iterativeCalculation` / `maxIterations` /
 *     `maxChange` — is a coherent authored unit: OOXML spells the attributes `iterate` /
 *     `iterateCount` / `iterateDelta` and an absent field means the Excel default (100 / 0.001), so
 *     on write the model is the full truth for these three (a None field emits no attribute and
 *     removes a stale preserved one). Professional financial models are routinely circular by
 *     design (interest on average debt balances) and ship with iterative calculation enabled.
 *   - The independent facts (GH-400) — `calcMode` / `fullCalcOnLoad` / `calcId` — where None means
 *     "no opinion": nothing is emitted on a scratch build and a preserved attribute rides through
 *     untouched on a source-backed write. Some(x) overlays x. This is what lets a from-scratch
 *     build author the TJC house `<calcPr calcId="191029" calcMode="autoNoTable" iterate="1"/>`
 *     without a post-write zip patch.
 *
 * Still-unmodeled calcPr attributes (refMode, concurrentCalc, calcOnSave, ...) are preserved on
 * write and never surface here.
 *
 * @param iterativeCalculation
 *   True enables bounded iterative evaluation of circular references (`iterate="1"`)
 * @param maxIterations
 *   Iteration cap (`iterateCount`, Excel default 100 when absent)
 * @param maxChange
 *   Convergence threshold between iterations (`iterateDelta`, Excel default 0.001 when absent)
 * @param calcMode
 *   Calculation mode (`calcMode`): manual | auto | autoNoTable; None = not authored
 * @param fullCalcOnLoad
 *   Force a full recalculation when the file opens (`fullCalcOnLoad`); None = not authored
 * @param calcId
 *   Calculation engine build stamp (`calcId`, e.g. 191029); diff tooling already excludes it. None =
 *   not authored (scratch builds need none — the LibreOffice precedent)
 */
final case class CalcPr(
  iterativeCalculation: Boolean = false,
  maxIterations: Option[Int] = None,
  maxChange: Option[BigDecimal] = None,
  calcMode: Option[CalcMode] = None,
  fullCalcOnLoad: Option[Boolean] = None,
  calcId: Option[Int] = None
)
