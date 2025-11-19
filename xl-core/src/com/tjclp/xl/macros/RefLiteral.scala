package com.tjclp.xl.macros

import com.tjclp.xl.addressing.{ARef, CellRange, RefParser, RefType, SheetName}
import com.tjclp.xl.addressing.RefParser.ParsedRef

import scala.quoted.*

/**
 * Compile-time validated unified reference literal.
 *
 * Supports all Excel reference formats with transparent return types:
 *   - `ref"A1"` → ARef (unwrapped)
 *   - `ref"A1:B10"` → CellRange (unwrapped)
 *   - `ref"Sales!A1"` → RefType.QualifiedCell
 *   - `ref"'Q1 Sales'!A1:B10"` → RefType.QualifiedRange
 *   - `ref"'It''s Q1'!A1"` → RefType.QualifiedCell (escaped quotes: '' → ')
 *
 * Uses zero-allocation parsing at compile time for maximum performance.
 *
 * Return type is transparent: simple refs return unwrapped ARef/CellRange for backwards
 * compatibility, qualified refs return RefType for sheets information.
 */
@SuppressWarnings(
  Array(
    "org.wartremover.warts.Var",
    "org.wartremover.warts.While"
  )
)
object RefLiteral:

  // -------- Public API --------
  extension (inline sc: StringContext)
    transparent inline def ref(): ARef | CellRange | RefType = ${ refImpl0('sc) }

    transparent inline def ref(
      inline args: Any*
    ): ARef | CellRange | RefType | Either[com.tjclp.xl.errors.XLError, RefType] =
      ${ refImplN('sc, 'args) }

  // -------- Implementation --------
  private def refImpl0(sc: Expr[StringContext])(using Quotes): Expr[ARef | CellRange | RefType] =
    import quotes.reflect.report
    val s = literal(sc)

    RefParser.parse(s) match
      case Right(parsed) => emitParsedLiteral(parsed)
      case Left(err) => report.errorAndAbort(s"Invalid ref literal '$s': $err")

  private def errorNoInterpolation(sc: Expr[StringContext], args: Expr[Seq[Any]], kind: String)(
    using Quotes
  ): Expr[Nothing] =
    import quotes.reflect.report
    args match
      case Varargs(Nil) => report.errorAndAbort(s"""Use $kind"...": no interpolation supported""")
      case _ => report.errorAndAbort(s"""$kind"...": interpolation not supported""")

  /**
   * Macro implementation supporting compile-time literals, compile-time optimization, and runtime
   * interpolation.
   *
   * Three branches:
   *   - No args (literal): Compile-time validation (Phase 1)
   *   - All args are compile-time constants: Compile-time optimization (Phase 2)
   *   - Any args are runtime variables: Runtime parsing (Phase 1)
   */
  private def refImplN(
    sc: Expr[StringContext],
    args: Expr[Seq[Any]]
  )(using Quotes): Expr[ARef | CellRange | RefType | Either[com.tjclp.xl.errors.XLError, RefType]] =
    import quotes.reflect.*

    args match
      case Varargs(exprs) if exprs.isEmpty =>
        // Branch 1: No interpolation - compile-time literal (Phase 1)
        refImpl0(sc)

      case Varargs(exprs) =>
        // Check if all arguments are compile-time constants
        MacroUtil.allLiterals(args) match
          case Some(literals) =>
            // Branch 2: All compile-time constants - OPTIMIZE (Phase 2)
            compileTimeOptimizedPath(sc, literals)

          case None =>
            // Branch 3: Has runtime variables - runtime parsing (Phase 1)
            runtimePath(sc, args)

  /**
   * Phase 2: Compile-time optimization for all-literal interpolations.
   *
   * Reconstructs the full string, parses it at compile time, and emits a constant. Returns
   * unwrapped ARef, CellRange, or RefType (same as non-interpolated literals).
   *
   * If parsing fails, aborts compilation with helpful errors message.
   */
  private def compileTimeOptimizedPath(
    sc: Expr[StringContext],
    literals: Seq[Any]
  )(using Quotes): Expr[ARef | CellRange | RefType] =
    import quotes.reflect.report

    val parts = sc.valueOrAbort.parts
    val fullString = MacroUtil.reconstructString(parts, literals)

    RefParser.parse(fullString) match
      case Right(parsed) => emitParsedLiteral(parsed)
      case Left(err) =>
        report.errorAndAbort(MacroUtil.formatCompileError("ref", fullString, err))

  private def emitParsedLiteral(parsed: ParsedRef)(using
    Quotes
  ): Expr[ARef | CellRange | RefType] =
    parsed match
      case ParsedRef.Cell(None, col0, row0) => constARef(col0, row0)
      case ParsedRef.Range(None, cs, rs, ce, re) => constCellRange(cs, rs, ce, re)
      case ParsedRef.Cell(Some(sheet), col0, row0) => constQualifiedCell(sheet, col0, row0)
      case ParsedRef.Range(Some(sheet), cs, rs, ce, re) =>
        constQualifiedRange(sheet, cs, rs, ce, re)

  /**
   * Phase 1: Runtime parsing for dynamic interpolations.
   *
   * Builds string at runtime and parses it with existing runtime parser. Returns Either[XLError,
   * RefType] for explicit errors handling.
   */
  private def runtimePath(
    sc: Expr[StringContext],
    args: Expr[Seq[Any]]
  )(using Quotes): Expr[Either[com.tjclp.xl.errors.XLError, RefType]] =
    '{
      val str = $sc.s($args*)
      RefType.parseToXLError(str)
    }.asExprOf[Either[com.tjclp.xl.errors.XLError, RefType]]

  private def literal(sc: Expr[StringContext])(using Quotes): String =
    val parts = sc.valueOrAbort.parts
    if parts.lengthCompare(1) != 0 then
      quotes.reflect.report.errorAndAbort("literal must be a single part")
    parts(0) // Safe: length == 1 verified above

  // -------- Const Emitters --------

  /**
   * Emit unwrapped ARef constant (for simple refs).
   *
   * Uses ARef.from0 public API for encapsulation. The inline function expands to the packed Long
   * representation at compile time with zero runtime overhead.
   */
  private def constARef(col0: Int, row0: Int)(using Quotes): Expr[ARef] =
    '{ ARef.from0(${ Expr(col0) }, ${ Expr(row0) }) }

  /** Emit unwrapped CellRange constant (for simple ranges) */
  private def constCellRange(cs: Int, rs: Int, ce: Int, re: Int)(using Quotes): Expr[CellRange] =
    '{ CellRange(${ constARef(cs, rs) }, ${ constARef(ce, re) }) }

  /** Emit RefType.QualifiedCell (for qualified refs) */
  private def constQualifiedCell(sheetName: String, col0: Int, row0: Int)(using
    Quotes
  ): Expr[RefType] =
    '{
      RefType.QualifiedCell(
        SheetName.unsafe(${ Expr(sheetName) }),
        ${ constARef(col0, row0) }
      )
    }

  /** Emit RefType.QualifiedRange (for qualified ranges) */
  private def constQualifiedRange(
    sheetName: String,
    cs: Int,
    rs: Int,
    ce: Int,
    re: Int
  )(using Quotes): Expr[RefType] =
    '{
      RefType.QualifiedRange(
        SheetName.unsafe(${ Expr(sheetName) }),
        CellRange(${ constARef(cs, rs) }, ${ constARef(ce, re) })
      )
    }

end RefLiteral

/** Export unified ref literal */
export RefLiteral.*
