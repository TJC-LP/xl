package com.tjclp.xl.codec

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.error.XLError

/**
 * Cell codec error types
 *
 * These errors are specific to cell-level encoding/decoding and can be converted to XLError when
 * needed.
 *
 * There is deliberately no `UncachedFormula` case (GH-589 revisited the GH-477 decision): typed
 * reads decode through `Cell.effectiveValue`, so a formula with no cached value surfaces as
 * `TypeMismatch(expected, Formula(_, None, _))` — `Cell.isUncachedFormula` and
 * `Sheet.readTypedStrict` make the distinction without a new error shape, and every exhaustive
 * `CodecError` match stays closed (in the evaluator: `EvalError.toXLError`,
 * `SheetEvaluator.evalErrorToXLError`; the CLI renders `XLError` and has no `CodecError` site).
 * Adding the case means changing those evaluator sites in the same commit; the recalculating writes
 * (`Excel.writeChecked` / `writeRecalculated`) remove the condition at the source instead.
 */
enum CodecError:
  /** Type mismatch when reading a cell */
  case TypeMismatch(expected: String, actual: CellValue)

  /** Parse error when converting cell value to target type */
  case ParseError(value: String, targetType: String, detail: String)

object CodecError:
  extension (error: CodecError)
    /** Convert CodecError to XLError for compatibility with general error handling */
    def toXLError(ref: ARef): XLError = error match
      case TypeMismatch(expected, actual) =>
        XLError.TypeMismatch(expected, actual.toString, ref.toA1)
      case ParseError(value, targetType, detail) =>
        XLError.ParseError(ref.toA1, s"Cannot parse '$value' as $targetType: $detail")
