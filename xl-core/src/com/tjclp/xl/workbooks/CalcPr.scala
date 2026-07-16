package com.tjclp.xl.workbooks

/**
 * Workbook calculation properties: the iterative-calculation subset of `<calcPr>` (GH-373).
 *
 * Professional financial models are routinely circular by design (interest on average debt
 * balances) and ship with iterative calculation enabled; a replica without it opens with circular
 * warnings the original never shows. The model uses friendly names — OOXML CT_CalcPr spells the
 * attributes `iterate` / `iterateCount` / `iterateDelta` and the codec maps between them. Unmodeled
 * calcPr attributes (calcId, fullCalcOnLoad, refMode, ...) are preserved on write and never surface
 * here.
 *
 * @param iterativeCalculation
 *   True enables bounded iterative evaluation of circular references (`iterate="1"`)
 * @param maxIterations
 *   Iteration cap (`iterateCount`, Excel default 100 when absent)
 * @param maxChange
 *   Convergence threshold between iterations (`iterateDelta`, Excel default 0.001 when absent)
 */
final case class CalcPr(
  iterativeCalculation: Boolean = false,
  maxIterations: Option[Int] = None,
  maxChange: Option[BigDecimal] = None
)
