package com.tjclp.xl.cells

import com.tjclp.xl.addressing.{ARef, CellRange}

/**
 * OOXML CT_CellFormula record kind for non-shared formula records (GH-430).
 *
 * `Normal` is a plain `<f>expr</f>` plus the calc flags any formula may carry; the other kinds add
 * the attributes of a real record. All of them must survive read -> model -> write on exactly the
 * cells whose XML carried them (per-cell faithfulness, no group inference). Shared formulas stay
 * expanded per GH-370 and never appear here.
 *
 * Note: the CSE case is named `ArrayFormula`, never `Array` — this enum is exported through
 * `com.tjclp.xl.api` and a member named `Array` would shadow `scala.Array` at wildcard-import
 * sites.
 */
enum FormulaKind derives CanEqual:
  /**
   * Plain `<f>expr</f>`, carrying the two calc flags CT_CellFormula allows on any formula: `ca`
   * ("calculate cell" — Excel's volatile marking) and `aca` ("always calculate array"). Flags
   * normalize like every other record attr, so LibreOffice's explicit `aca="false"` reads as false
   * and re-emits omitted — the schema default (GH-435).
   */
  case Normal(aca: Boolean = false, ca: Boolean = false)

  /**
   * `<f t="array" ref="..">expr</f>` — an array-formula anchor whose result fills `ref`. `mode`
   * tells a legacy CSE record (`{=…}`, Ctrl+Shift+Enter) from an Excel 365 dynamic array (a
   * spilling formula, flagged on disk by the cell's `cm` attribute pointing at XLDAPR dynamic-array
   * properties in `xl/metadata.xml`, GH-714). Both evaluate in array mode; only the mode decides
   * which one Excel shows and whether it re-spills.
   */
  case ArrayFormula(
    ref: CellRange,
    aca: Boolean = false,
    ca: Boolean = false,
    mode: ArrayMode = ArrayMode.Legacy
  )

  /**
   * `<f t="dataTable" .../>` — carries no formula text in XML; the owning Formula's expression is
   * the derived display text (see [[FormulaKind.displayExpression]]). `r1`/`r2` are `Option`
   * because files with `del1`/`del2` legitimately omit the deleted input cell.
   */
  case DataTable(
    ref: CellRange,
    dt2D: Boolean,
    dtr: Boolean,
    r1: Option[ARef],
    r2: Option[ARef],
    del1: Boolean = false,
    del2: Boolean = false,
    ca: Boolean = false
  )

object FormulaKind:
  /** An Excel 365 dynamic-array anchor whose spill fills `extent` (anchor = `extent.start`). */
  def dynamicArray(extent: CellRange): FormulaKind.ArrayFormula =
    ArrayFormula(extent, mode = ArrayMode.Dynamic())

  extension (k: FormulaKind)
    /** True for an Excel 365 dynamic-array anchor (GH-714); false for CSE and every other kind. */
    def isDynamicArray: Boolean = k match
      case ArrayFormula(_, _, _, ArrayMode.Dynamic(_)) => true
      case _ => false

  /**
   * Excel formula-bar text for a data table record, without braces or a leading `=`: a 2-D table
   * renders `TABLE(r1,r2)`; a 1-D row-oriented table (`dtr`) renders `TABLE(r1,)`; a 1-D
   * column-oriented table renders `TABLE(,r1)`. A missing input renders an empty slot.
   */
  def displayExpression(dt: FormulaKind.DataTable): String =
    val first = dt.r1.map(_.toA1).getOrElse("")
    val second = dt.r2.map(_.toA1).getOrElse("")
    if dt.dt2D then s"TABLE($first,$second)"
    else if dt.dtr then s"TABLE($first,)"
    else s"TABLE(,$first)"

/**
 * How an [[FormulaKind.ArrayFormula]] record is stored and shown (GH-714).
 *
 * `Legacy` is the CSE array Excel shows in braces (`{=SUM(A1:A3*B1:B3)}`): `t="array"` and no `cm`.
 * `Dynamic` is the Excel 365 spilling formula: the same `<f t="array">` record plus the cell's `cm`
 * attribute, resolved through `xl/metadata.xml` to XLDAPR dynamic-array properties with `fDynamic`
 * set. `collapsed` carries XLDAPR's `fCollapsed` losslessly. The model never holds the raw `cm`
 * index; the writer allocates it from the metadata part it ships.
 */
enum ArrayMode derives CanEqual:
  case Legacy
  case Dynamic(collapsed: Boolean = false)
