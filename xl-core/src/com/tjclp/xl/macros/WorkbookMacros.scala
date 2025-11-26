package com.tjclp.xl.macros

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

import scala.quoted.*

/**
 * Compile-time validated string refs for Workbook string-based methods.
 *
 * When `workbook("Sales")` is called with a string literal, the sheet name format is validated at
 * compile time. Invalid literals like `"Invalid:Name"` fail to compile.
 *
 * Sheet existence is always checked at runtime (returns XLResult for get operations).
 */
object WorkbookMacros:

  // Re-use SheetName validation logic
  private val InvalidChars = Set(':', '\\', '/', '?', '*', '[', ']')
  private val MaxLength = 31
  private val ReservedNames = Set("history")

  /** Validate sheet name at compile time (returns None if valid, Some(error) if invalid) */
  private def validateSheetName(name: String): Option[String] =
    if name.isEmpty then Some("Sheet name cannot be empty")
    else if name.length > MaxLength then Some(s"Sheet name too long (max $MaxLength): $name")
    else if name.exists(InvalidChars.contains) then
      Some(s"Sheet name contains invalid characters: $name")
    else if ReservedNames.contains(name.toLowerCase) then
      Some(s"Sheet name '$name' is reserved by Excel")
    else None

  /**
   * Macro impl for workbook(name: String) - validates literal name format at compile time.
   *
   * Always returns XLResult[Sheet] because sheet existence is runtime-dependent. Invalid literal
   * names fail at compile time.
   */
  def applyImpl(
    workbook: Expr[Workbook],
    name: Expr[String]
  )(using Quotes): Expr[XLResult[Sheet]] =
    import quotes.reflect.*

    name.value match
      case Some(literal) =>
        // Compile-time literal - validate name format now
        validateSheetName(literal) match
          case Some(err) =>
            report.errorAndAbort(s"Invalid sheet name '$literal': $err")
          case None =>
            // Valid name format - emit call with pre-validated SheetName
            // Still returns XLResult because sheet might not exist
            '{ $workbook.apply(SheetName.unsafe(${ Expr(literal) })) }

      case None =>
        // Runtime expression - defer all validation
        '{
          SheetName($name).left
            .map(err => XLError.InvalidSheetName($name, err): XLError)
            .flatMap(sn => $workbook.apply(sn))
        }

  /**
   * Macro impl for workbook.put(name: String, sheet: Sheet).
   *
   * Returns Workbook for literal refs (name format validated at compile time, operation is
   * infallible). Returns XLResult[Workbook] for runtime refs.
   */
  def putStringNameImpl(
    workbook: Expr[Workbook],
    name: Expr[String],
    sheet: Expr[Sheet]
  )(using Quotes): Expr[Workbook | XLResult[Workbook]] =
    import quotes.reflect.*

    name.value match
      case Some(literal) =>
        // Compile-time literal - validate name format now
        validateSheetName(literal) match
          case Some(err) =>
            report.errorAndAbort(s"Invalid sheet name '$literal': $err")
          case None =>
            // Valid name format - emit direct put call (returns Workbook)
            '{ $workbook.put(SheetName.unsafe(${ Expr(literal) }), $sheet) }

      case None =>
        // Runtime expression - defer validation (returns XLResult[Workbook])
        '{
          SheetName($name) match
            case Right(sn) => Right($workbook.put(sn, $sheet))
            case Left(err) => Left(XLError.InvalidSheetName($name, err))
        }

  /**
   * Macro impl for workbook.remove(name: String).
   *
   * Validates name format at compile time for literals. Always returns XLResult[Workbook] because
   * sheet existence is runtime-dependent.
   */
  def removeImpl(
    workbook: Expr[Workbook],
    name: Expr[String]
  )(using Quotes): Expr[XLResult[Workbook]] =
    import quotes.reflect.*

    name.value match
      case Some(literal) =>
        // Compile-time literal - validate name format now
        validateSheetName(literal) match
          case Some(err) =>
            report.errorAndAbort(s"Invalid sheet name '$literal': $err")
          case None =>
            // Valid name format - emit direct remove call (still XLResult due to existence check)
            '{ $workbook.remove(SheetName.unsafe(${ Expr(literal) })) }

      case None =>
        // Runtime expression - defer validation
        '{
          SheetName($name).left
            .map(err => XLError.InvalidSheetName($name, err))
            .flatMap(sn => $workbook.remove(sn))
        }

  /**
   * Macro impl for workbook.update(name: String, f: Sheet => Sheet).
   *
   * Validates name format at compile time for literals. Always returns XLResult[Workbook] because
   * sheet existence is runtime-dependent.
   */
  def updateImpl(
    workbook: Expr[Workbook],
    name: Expr[String],
    f: Expr[Sheet => Sheet]
  )(using Quotes): Expr[XLResult[Workbook]] =
    import quotes.reflect.*

    name.value match
      case Some(literal) =>
        // Compile-time literal - validate name format now
        validateSheetName(literal) match
          case Some(err) =>
            report.errorAndAbort(s"Invalid sheet name '$literal': $err")
          case None =>
            // Valid name format - emit direct update call (still XLResult due to existence check)
            '{ $workbook.update(SheetName.unsafe(${ Expr(literal) }), $f) }

      case None =>
        // Runtime expression - defer validation
        '{
          SheetName($name).left
            .map(err => XLError.InvalidSheetName($name, err))
            .flatMap(sn => $workbook.update(sn, $f))
        }

  /**
   * Macro impl for workbook.delete(name: String).
   *
   * Validates name format at compile time for literals. Always returns XLResult[Workbook] because
   * sheet existence is runtime-dependent (same as remove).
   */
  def deleteImpl(
    workbook: Expr[Workbook],
    name: Expr[String]
  )(using Quotes): Expr[XLResult[Workbook]] =
    import quotes.reflect.*

    name.value match
      case Some(literal) =>
        // Compile-time literal - validate name format now
        validateSheetName(literal) match
          case Some(err) =>
            report.errorAndAbort(s"Invalid sheet name '$literal': $err")
          case None =>
            // Valid name format - emit direct delete call (still XLResult due to existence check)
            '{ $workbook.delete(SheetName.unsafe(${ Expr(literal) })) }

      case None =>
        // Runtime expression - defer validation
        '{
          SheetName($name).left
            .map(err => XLError.InvalidSheetName($name, err))
            .flatMap(sn => $workbook.delete(sn))
        }

  /**
   * Macro impl for workbook.get(name: String) extension.
   *
   * Validates name format at compile time for literals. Always returns Option[Sheet] because sheet
   * existence is runtime-dependent.
   */
  def getImpl(
    workbook: Expr[Workbook],
    name: Expr[String]
  )(using Quotes): Expr[Option[Sheet]] =
    import quotes.reflect.*

    name.value match
      case Some(literal) =>
        // Compile-time literal - validate name format now
        validateSheetName(literal) match
          case Some(err) =>
            report.errorAndAbort(s"Invalid sheet name '$literal': $err")
          case None =>
            // Valid name format - emit direct lookup
            '{ $workbook.apply(SheetName.unsafe(${ Expr(literal) })).toOption }

      case None =>
        // Runtime expression - defer validation
        '{ $workbook.apply($name).toOption }

end WorkbookMacros
