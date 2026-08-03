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

  /** `<f t="array" ref="..">expr</f>` — legacy CSE anchor. */
  case ArrayFormula(ref: CellRange, aca: Boolean = false, ca: Boolean = false)

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
