package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import com.tjclp.xl.cells.CellValue
import TExpr.*

trait TExprCoercions:
  // ===== PolyRef Conversion Helpers =====

  /**
   * Convert any TExpr to String type with coercion.
   *
   * Used by text functions (LEFT, RIGHT, UPPER, etc.) to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asStringExpr(expr: TExpr[?]): TExpr[String] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeAsString)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeAsString)
    case other => other.asInstanceOf[TExpr[String]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to LocalDate type with coercion.
   *
   * Used by date functions (YEAR, MONTH, DAY) to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asDateExpr(expr: TExpr[?]): TExpr[java.time.LocalDate] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeAsDate)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeAsDate)
    case other =>
      other.asInstanceOf[TExpr[java.time.LocalDate]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to Int type with coercion.
   *
   * Used by functions requiring integer arguments (LEFT, RIGHT, DATE). Automatically converts
   * BigDecimal expressions (like YEAR/MONTH/DAY) to Int.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asIntExpr(expr: TExpr[?]): TExpr[Int] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeAsInt)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeAsInt)
    case TExpr.Lit(bd: BigDecimal) if bd.isValidInt => TExpr.Lit(bd.toInt)
    // Convert BigDecimal expressions to Int (YEAR/MONTH/DAY/LEN return BigDecimal)
    case call: TExpr.Call[?] if call.spec == FunctionSpecs.year =>
      ToInt(call.asInstanceOf[TExpr[BigDecimal]])
    case call: TExpr.Call[?] if call.spec == FunctionSpecs.month =>
      ToInt(call.asInstanceOf[TExpr[BigDecimal]])
    case call: TExpr.Call[?] if call.spec == FunctionSpecs.day =>
      ToInt(call.asInstanceOf[TExpr[BigDecimal]])
    case call: TExpr.Call[?] if call.spec == FunctionSpecs.len =>
      ToInt(call.asInstanceOf[TExpr[BigDecimal]])
    case other => other.asInstanceOf[TExpr[Int]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to BigDecimal type (numeric).
   *
   * Used by arithmetic functions to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asNumericExpr(expr: TExpr[?]): TExpr[BigDecimal] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeNumeric)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeNumeric)
    // Date functions return LocalDate/LocalDateTime - convert to Excel serial number
    case call: TExpr.Call[?] if call.spec.flags.returnsDate =>
      DateToSerial(call.asInstanceOf[TExpr[java.time.LocalDate]])
    case call: TExpr.Call[?] if call.spec.flags.returnsTime =>
      DateTimeToSerial(call.asInstanceOf[TExpr[java.time.LocalDateTime]])
    case other =>
      other.asInstanceOf[TExpr[BigDecimal]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to Boolean type.
   *
   * Used by logical functions (AND, OR, NOT, IF) to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asBooleanExpr(expr: TExpr[?]): TExpr[Boolean] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeBool)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeBool)
    case other => other.asInstanceOf[TExpr[Boolean]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to CellValue type.
   *
   * Used by error handling functions (IFERROR, ISERROR) to preserve raw cell values.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asCellValueExpr(expr: TExpr[?]): TExpr[CellValue] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeCellValue)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeCellValue)
    case other =>
      other.asInstanceOf[TExpr[CellValue]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to resolved CellValue type.
   *
   * Used for standalone cell references (e.g., =A1, =Sheet1!B2) where we need the cell's
   * "effective" value with cached formula results extracted and empty cells converted to 0.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asResolvedValueExpr(expr: TExpr[?]): TExpr[CellValue] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeResolvedValue)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeResolvedValue)
    case other =>
      other.asInstanceOf[TExpr[CellValue]] // Safe: non-PolyRef already has correct type
