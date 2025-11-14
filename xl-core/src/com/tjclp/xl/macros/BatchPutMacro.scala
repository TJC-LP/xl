package com.tjclp.xl.macros

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.sheet.Sheet

import scala.quoted.*

/** Batch put macro for elegant multi-cell updates */
object BatchPutMacro:
  /**
   * Batch put with automatic CellValue conversion.
   *
   * Usage:
   * {{{
   * import com.tjclp.xl.putMacro.put
   *
   * sheet.put(
   *   ref"A1" -> "Hello",
   *   ref"B1" -> 42,
   *   ref"C1" -> true
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

          // Generate put call - check if value is Formatted to preserve NumFmt
          value.asTerm.tpe.asType match
            // Case 1: Formatted value - preserve NumFmt metadata
            case '[com.tjclp.xl.formatted.Formatted] =>
              val formattedValue = value.asExprOf[com.tjclp.xl.formatted.Formatted]
              '{
                import com.tjclp.xl.sheet.styleSyntax.withCellStyle
                val formatted = $formattedValue
                val style =
                  com.tjclp.xl.style.CellStyle.default.withNumFmt(formatted.numFmt)
                $sheetExpr
                  .put($ref, formatted.value)
                  .withCellStyle($ref, style)
              }

            // Case 2: All other types - use CellValue.from converter
            case _ =>
              '{ $sheetExpr.put($ref, com.tjclp.xl.cell.CellValue.from($value)) }
        }

      case _ =>
        report.errorAndAbort("Batch put requires literal tuple arguments")

end BatchPutMacro

export BatchPutMacro.*
