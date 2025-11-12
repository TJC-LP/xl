package com.tjclp.xl

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.sheet.Sheet
import scala.quoted.*

/** Batch put macro for elegant multi-cell updates */
object putMacro:
  /**
   * Batch put with automatic CellValue conversion.
   *
   * Usage:
   * {{{
   * import com.tjclp.xl.putMacro.put
   *
   * sheet.put(
   *   cell"A1" -> "Hello",
   *   cell"B1" -> 42,
   *   cell"C1" -> true
   * )
   * }}}
   *
   * Expands to chained .put() calls with zero intermediate allocations.
   */
  extension (sheet: Sheet)
    transparent inline def put(inline pairs: (ARef, Any)*): Sheet =
      ${ putImpl('sheet, 'pairs) }

  private def putImpl(sheet: Expr[Sheet], pairs: Expr[Seq[(ARef, Any)]])(using
    Quotes
  ): Expr[Sheet] =
    import quotes.reflect.*

    pairs match
      case Varargs(pairExprs) =>
        // Build chained put calls
        pairExprs.foldLeft(sheet) { (sheetExpr, pairExpr) =>
          // Extract ArrowAssoc(ref).->(value) components
          val (ref, value) = pairExpr.asTerm match
            // Match: ref -> value (which becomes ArrowAssoc(ref).->(value))
            case Apply(TypeApply(Select(Apply(_, List(r)), "->"), _), List(v)) =>
              (r.asExprOf[ARef], v.asExpr)
            // Also try inlined version
            case Inlined(_, _, Apply(TypeApply(Select(Apply(_, List(r)), "->"), _), List(v))) =>
              (r.asExprOf[ARef], v.asExpr)
            case _ =>
              report.errorAndAbort("Batch put requires tuple pairs: cell\"A1\" -> value")

          // Generate CellValue based on runtime value (use CellValue.from)
          '{ $sheetExpr.put($ref, com.tjclp.xl.cell.CellValue.from($value)) }
        }

      case _ =>
        report.errorAndAbort("Batch put requires literal tuple arguments")

end putMacro
