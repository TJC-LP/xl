package com.tjclp.xl.formula.graph

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef

/**
 * #790: the parse of each formula cell a recalculation graph was built from, for the evaluation of
 * the same pass.
 *
 * [[get]] answers only for the cell the text was read from, and only while the caller's text equals
 * it, so a formula edited after the build is parsed again. Lookups go through the cell's node id:
 * no map here is keyed by a `TExpr`, whose `BigDecimal` literals can make `hashCode` throw (#780).
 */
private[formula] trait FormulaParses:

  /** The build's parse of `text` at `ref` on `sheet`; None when the build holds no parse of it. */
  def get(sheet: SheetName, ref: ARef, text: String): Option[TExpr[?]]

private[formula] object FormulaParses:

  /** Holds nothing: every caller parses. */
  val none: FormulaParses = (_, _, _) => None

  /**
   * The parses of `nodes`' formulas: `texts(n)` is formula `n`'s text and `parsed(n)` its parse
   * (None when it does not parse).
   */
  def of(
    nodes: NodeTable,
    texts: IArray[String],
    parsed: IArray[Option[TExpr[?]]]
  ): FormulaParses =
    (sheet, ref, text) =>
      nodes
        .lookup(QualifiedRef(sheet, ref))
        .flatMap(n => if texts(n) == text then parsed(n) else None)
