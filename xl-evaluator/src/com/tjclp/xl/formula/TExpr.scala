package com.tjclp.xl.formula

import com.tjclp.xl.{ARef, Anchor, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.codec.CodecError

import scala.math.BigDecimal

/**
 * Typed expression GADT for Excel formulas.
 *
 * TExpr[A] represents a formula expression that evaluates to type A. This is a purely functional,
 * total representation - all operations return Either for error handling, no exceptions thrown.
 *
 * Type parameter A captures the result type:
 *   - TExpr[BigDecimal] for numeric expressions
 *   - TExpr[Boolean] for boolean/logical expressions
 *   - TExpr[String] for text expressions
 *
 * Laws satisfied:
 *   1. If-fusion: If(c, Lit(x), Lit(y)) ≡ Lit(if ⟦c⟧ then x else y)
 *   2. Ring laws: Add/Mul form commutative semiring over BigDecimal
 *   3. Short-circuit: And/Or respect left-to-right semantics
 *   4. Round-trip: parse(print(expr)) == Right(expr)
 */
enum TExpr[A] derives CanEqual:
  /**
   * Literal value - compile-time constant.
   *
   * Example: Lit(42) represents the number 42
   */
  case Lit[A](value: A) extends TExpr[A]

  /**
   * Cell reference - reads value from a cell at runtime.
   *
   * @param at
   *   The cell address to read from
   * @param anchor
   *   Anchoring mode for formula dragging (default: Relative)
   * @param decode
   *   Function to decode the cell's value to type A
   *
   * Example: Ref(ARef("A1"), Anchor.Relative, decodeNumber) reads numeric value from A1
   */
  case Ref[A](at: ARef, anchor: Anchor, decode: Cell => Either[CodecError, A]) extends TExpr[A]

  /**
   * Polymorphic cell reference - defers type commitment until function context.
   *
   * This is used during parsing when the expected type of a cell reference is not yet known.
   * Function parsers convert PolyRef to Ref with appropriate type-coercing decoder based on
   * context.
   *
   * Matches Excel's type coercion semantics: A1=42 in =LEFT(A1,1) converts 42→"4".
   *
   * @param at
   *   The cell address to read from
   * @param anchor
   *   Anchoring mode for formula dragging (default: Relative)
   *
   * Example: PolyRef(ARef("A1"), Anchor.Relative) - type determined by enclosing function
   */
  case PolyRef(at: ARef, anchor: Anchor = Anchor.Relative) extends TExpr[Nothing]

  /**
   * Sheet-qualified cell reference - reads value from a cell in another sheet.
   *
   * @param sheet
   *   The target sheet name
   * @param at
   *   The cell address within the target sheet
   * @param anchor
   *   Anchoring mode for formula dragging
   * @param decode
   *   Function to decode the cell's value to type A
   *
   * Example: SheetRef("Sales", ARef("F4"), Anchor.Relative, decodeNumber) reads from Sales!F4
   */
  case SheetRef[A](
    sheet: SheetName,
    at: ARef,
    anchor: Anchor,
    decode: Cell => Either[CodecError, A]
  ) extends TExpr[A]

  /**
   * Polymorphic sheet-qualified cell reference - defers type commitment until function context.
   *
   * @param sheet
   *   The target sheet name
   * @param at
   *   The cell address within the target sheet
   * @param anchor
   *   Anchoring mode for formula dragging
   *
   * Example: SheetPolyRef("Sales", ARef("F4"), Anchor.Relative) - type determined by context
   */
  case SheetPolyRef(sheet: SheetName, at: ARef, anchor: Anchor = Anchor.Relative)
      extends TExpr[Nothing]

  /**
   * Local range reference.
   *
   * Represents a range like A1:B10 before being consumed by a function.
   */
  case RangeRef(range: CellRange) extends TExpr[Nothing]

  /**
   * Sheet-qualified range reference - references a range in another sheet.
   *
   * @param sheet
   *   The target sheet name
   * @param range
   *   The cell range within the target sheet
   *
   * Example: SheetRange("Sales", CellRange(A1, A10)) represents Sales!A1:A10
   */
  case SheetRange(sheet: SheetName, range: CellRange) extends TExpr[Nothing]

  // Arithmetic operators (form commutative semiring over BigDecimal)

  /**
   * Addition: x + y
   *
   * Laws:
   *   - Commutative: Add(x, y) = Add(y, x)
   *   - Associative: Add(Add(x, y), z) = Add(x, Add(y, z))
   *   - Identity: Add(x, Lit(0)) = x
   */
  case Add(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Subtraction: x - y
   */
  case Sub(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Multiplication: x * y
   *
   * Laws:
   *   - Commutative: Mul(x, y) = Mul(y, x)
   *   - Associative: Mul(Mul(x, y), z) = Mul(x, Mul(y, z))
   *   - Identity: Mul(x, Lit(1)) = x
   *   - Distributive: Mul(x, Add(y, z)) = Add(Mul(x, y), Mul(x, z))
   */
  case Mul(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Division: x / y
   *
   * Note: Division by zero returns evaluation error (not parse error)
   */
  case Div(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  // Range aggregation

  // Comparison operators (return Boolean)

  /**
   * Equality: x == y
   */
  case Eq[A](x: TExpr[A], y: TExpr[A]) extends TExpr[Boolean]

  /**
   * Inequality: x != y
   */
  case Neq[A](x: TExpr[A], y: TExpr[A]) extends TExpr[Boolean]

  /**
   * Less than: x < y (numeric comparison)
   */
  case Lt(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]

  /**
   * Less than or equal: x <= y
   */
  case Lte(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]

  /**
   * Greater than: x > y
   */
  case Gt(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]

  /**
   * Greater than or equal: x >= y
   */
  case Gte(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]

  // Type conversions

  /**
   * Convert BigDecimal to Int.
   *
   * Required for DATE(YEAR(A1), MONTH(A1), DAY(A1)) pattern where date extraction functions return
   * BigDecimal but DATE expects Int arguments.
   *
   * Validates at evaluation time that value is a valid integer.
   *
   * Example: ToInt(YEAR(A1)) where YEAR returns BigDecimal(2025) → Int(2025)
   */
  case ToInt(expr: TExpr[BigDecimal]) extends TExpr[Int]

  /**
   * Convert LocalDate to Excel serial number.
   *
   * Excel represents dates as the number of days since December 30, 1899. This enables date
   * arithmetic like TODAY()+30.
   *
   * Example: DateToSerial(TODAY()) where TODAY() = 2025-12-20 → BigDecimal(46011)
   */
  case DateToSerial(expr: TExpr[java.time.LocalDate]) extends TExpr[BigDecimal]

  /**
   * Convert LocalDateTime to Excel serial number.
   *
   * Excel represents dates/times as days since epoch + fractional day for time. This enables
   * datetime arithmetic like NOW()+0.5.
   *
   * Example: DateTimeToSerial(NOW()) where NOW() = 2025-12-20 12:00 → BigDecimal(46011.5)
   */
  case DateTimeToSerial(expr: TExpr[java.time.LocalDateTime]) extends TExpr[BigDecimal]

  /**
   * Generic aggregate function over a range using the Aggregator typeclass.
   *
   * This case unifies all simple aggregate functions (SUM, COUNT, MIN, MAX, AVERAGE) into a single
   * representation. The aggregatorId determines which Aggregator instance to use for evaluation.
   *
   * Benefits:
   *   - Single pattern match case in Evaluator, FormulaPrinter, FormulaShifter, DependencyGraph
   *   - Adding new aggregates only requires a new Aggregator given instance
   *   - Automatically supports full row/column references via RangeLocation
   *
   * @param aggregatorId
   *   The name of the aggregator (SUM, COUNT, MIN, MAX, AVERAGE)
   * @param location
   *   The range to aggregate (local or cross-sheet)
   *
   * Example: Aggregate("SUM", RangeLocation.Local(A1:A10))
   */
  case Aggregate(aggregatorId: String, location: TExpr.RangeLocation) extends TExpr[BigDecimal]

  /**
   * Generic function call backed by FunctionSpec.
   *
   * Centralizes function definitions so parser, evaluator, printer, shifter, and dependency
   * analysis can share a single spec.
   */
  case Call[A](spec: FunctionSpec[A], args: spec.Args) extends TExpr[A]

object TExpr:

  // ===== Range Location Abstraction =====

  /**
   * Where a range is located - same sheet or cross-sheet.
   *
   * This enum unifies local and cross-sheet range references, eliminating the need for paired TExpr
   * cases (e.g., Min + SheetMin). Used by TExpr.Aggregate for unified aggregation.
   */
  enum RangeLocation derives CanEqual:
    case Local(range: CellRange)
    case CrossSheet(sheet: SheetName, range: CellRange)

  object RangeLocation:
    extension (loc: RangeLocation)
      /** Extract the CellRange regardless of location */
      def range: CellRange = loc match
        case Local(r) => r
        case CrossSheet(_, r) => r

      /** Get sheet name for cross-sheet, None for local */
      def sheetName: Option[SheetName] = loc match
        case CrossSheet(s, _) => Some(s)
        case _ => None

      /** Get cells for local ranges only (for intra-sheet dependency graphs) */
      def localCells: Set[ARef] = loc match
        case Local(r) => r.cells.toSet
        case CrossSheet(_, _) => Set.empty

      /**
       * Get cells for local ranges, bounded by the sheet's used range.
       *
       * Preferred for dependency graph construction to avoid materializing 1M+ cells for full
       * column/row references like A:A or 1:1.
       *
       * @param bounds
       *   Optional bounding range (typically sheet.usedRange)
       * @return
       *   Set of cell references in the intersection of this range and bounds
       */
      def localCellsBounded(bounds: Option[CellRange]): Set[ARef] = loc match
        case Local(r) =>
          bounds match
            case Some(b) => r.intersect(b).map(_.cells.toSet).getOrElse(Set.empty)
            case None => r.cells.toSet
        case CrossSheet(_, _) => Set.empty

      /** Check if this is a cross-sheet reference */
      def isCrossSheet: Boolean = loc match
        case CrossSheet(_, _) => true
        case _ => false

      /** Get all cells in the range (delegates to underlying CellRange) */
      def cells: Iterator[ARef] = loc.range.cells

      /** Get width of the range */
      def width: Int = loc.range.width

      /** Get height of the range */
      def height: Int = loc.range.height

      /** Get A1 string representation */
      def toA1: String = loc match
        case Local(r) => r.toA1
        case CrossSheet(s, r) => s"${s.value}!${r.toA1}"

      /** Get starting column of the range */
      def colStart: Column = loc.range.colStart

      /** Get ending column of the range */
      def colEnd: Column = loc.range.colEnd

      /** Get starting row of the range */
      def rowStart: Row = loc.range.rowStart

      /** Get ending row of the range */
      def rowEnd: Row = loc.range.rowEnd

      /** Get start cell reference of the range */
      def start: ARef = loc.range.start

      /** Get end cell reference of the range */
      def end: ARef = loc.range.end

  // ===== Smart Constructors =====

  /**
   * Smart constructor for literals.
   *
   * Example: TExpr.lit(42)
   */
  def lit[A](value: A): TExpr[A] = Lit(value)

  /**
   * Smart constructor for cell references.
   *
   * Example: TExpr.ref(ARef("A1"), Anchor.Relative, codec)
   */
  def ref[A](at: ARef, anchor: Anchor, decode: Cell => Either[CodecError, A]): TExpr[A] =
    Ref(at, anchor, decode)

  /**
   * Smart constructor for cell references with default Relative anchor.
   *
   * Example: TExpr.ref(ARef("A1"), codec)
   */
  def ref[A](at: ARef, decode: Cell => Either[CodecError, A]): TExpr[A] =
    Ref(at, Anchor.Relative, decode)

  /**
   * Smart constructor for conditionals.
   *
   * Example: TExpr.cond(test, ifTrue, ifFalse)
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def cond[A](test: TExpr[Boolean], ifTrue: TExpr[A], ifFalse: TExpr[A]): TExpr[A] =
    Call(
      FunctionSpecs.ifFn,
      (test, ifTrue.asInstanceOf[TExpr[Any]], ifFalse.asInstanceOf[TExpr[Any]])
    ).asInstanceOf[TExpr[A]]

  // Convenience constructors for common operations

  /**
   * SUM aggregation: sum all numeric values in range.
   *
   * Example: TExpr.sum(CellRange("A1:A10"))
   */
  def sum(range: CellRange): TExpr[BigDecimal] =
    Call(FunctionSpecs.sum, RangeLocation.Local(range))

  /**
   * COUNT aggregation: count numeric cells in range.
   *
   * Example: TExpr.count(CellRange("A1:A10"))
   */
  def count(range: CellRange): TExpr[BigDecimal] =
    Call(FunctionSpecs.count, RangeLocation.Local(range))

  /**
   * AVERAGE aggregation: average of numeric values in range.
   *
   * Example: TExpr.average(CellRange("A1:A10"))
   */
  def average(range: CellRange): TExpr[BigDecimal] =
    Call(FunctionSpecs.average, RangeLocation.Local(range))

  /**
   * MIN aggregation: minimum numeric value in range.
   *
   * Example: TExpr.min(CellRange("A1:A10"))
   */
  def min(range: CellRange): TExpr[BigDecimal] =
    Call(FunctionSpecs.min, RangeLocation.Local(range))

  /**
   * MAX aggregation: maximum numeric value in range.
   *
   * Example: TExpr.max(CellRange("A1:A10"))
   */
  def max(range: CellRange): TExpr[BigDecimal] =
    Call(FunctionSpecs.max, RangeLocation.Local(range))

  // Financial function smart constructors

  /**
   * Smart constructor for NPV over a range of cash flows.
   *
   * Example: TExpr.npv(TExpr.Lit(BigDecimal("0.1")), CellRange("A2:A6"))
   */
  def npv(rate: TExpr[BigDecimal], values: CellRange): TExpr[BigDecimal] =
    Call(FunctionSpecs.npv, (rate, values))

  /**
   * Smart constructor for IRR with optional guess.
   *
   * Example: TExpr.irr(CellRange("A1:A6"), Some(TExpr.Lit(BigDecimal("0.15"))))
   */
  def irr(values: CellRange, guess: Option[TExpr[BigDecimal]] = None): TExpr[BigDecimal] =
    Call(FunctionSpecs.irr, (values, guess))

  /**
   * Smart constructor for XNPV with irregular dates.
   *
   * @param rate
   *   Discount rate
   * @param values
   *   Cash flow values range
   * @param dates
   *   Corresponding dates range
   *
   * Example: TExpr.xnpv(TExpr.Lit(0.1), valuesRange, datesRange)
   */
  def xnpv(
    rate: TExpr[BigDecimal],
    values: CellRange,
    dates: CellRange
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.xnpv, (rate, values, dates))

  /**
   * Smart constructor for XIRR with irregular dates.
   *
   * @param values
   *   Cash flow values range (must have positive and negative)
   * @param dates
   *   Corresponding dates range
   * @param guess
   *   Optional starting rate for Newton-Raphson (default 0.1)
   *
   * Example: TExpr.xirr(valuesRange, datesRange, Some(TExpr.Lit(0.15)))
   */
  def xirr(
    values: CellRange,
    dates: CellRange,
    guess: Option[TExpr[BigDecimal]] = None
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.xirr, (values, dates, guess))

  // ===== TVM Smart Constructors =====

  /**
   * PMT: calculate payment per period.
   *
   * Example: TExpr.pmt(TExpr.Lit(0.05/12), TExpr.Lit(24), TExpr.Lit(10000))
   */
  def pmt(
    rate: TExpr[BigDecimal],
    nper: TExpr[BigDecimal],
    pv: TExpr[BigDecimal],
    fv: Option[TExpr[BigDecimal]] = None,
    pmtType: Option[TExpr[BigDecimal]] = None
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.pmt, (rate, nper, pv, fv, pmtType))

  /**
   * FV: calculate future value.
   *
   * Example: TExpr.fv(TExpr.Lit(0.05/12), TExpr.Lit(24), TExpr.Lit(-100))
   */
  def fv(
    rate: TExpr[BigDecimal],
    nper: TExpr[BigDecimal],
    pmt: TExpr[BigDecimal],
    pv: Option[TExpr[BigDecimal]] = None,
    pmtType: Option[TExpr[BigDecimal]] = None
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.fv, (rate, nper, pmt, pv, pmtType))

  /**
   * PV: calculate present value.
   *
   * Example: TExpr.pv(TExpr.Lit(0.05/12), TExpr.Lit(24), TExpr.Lit(-500))
   */
  def pv(
    rate: TExpr[BigDecimal],
    nper: TExpr[BigDecimal],
    pmt: TExpr[BigDecimal],
    fv: Option[TExpr[BigDecimal]] = None,
    pmtType: Option[TExpr[BigDecimal]] = None
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.pv, (rate, nper, pmt, fv, pmtType))

  /**
   * NPER: calculate number of periods.
   *
   * Example: TExpr.nper(TExpr.Lit(0.05/12), TExpr.Lit(-500), TExpr.Lit(10000))
   */
  def nper(
    rate: TExpr[BigDecimal],
    pmt: TExpr[BigDecimal],
    pv: TExpr[BigDecimal],
    fv: Option[TExpr[BigDecimal]] = None,
    pmtType: Option[TExpr[BigDecimal]] = None
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.nper, (rate, pmt, pv, fv, pmtType))

  /**
   * RATE: calculate interest rate per period.
   *
   * Example: TExpr.rate(TExpr.Lit(24), TExpr.Lit(-500), TExpr.Lit(10000))
   */
  def rate(
    nper: TExpr[BigDecimal],
    pmt: TExpr[BigDecimal],
    pv: TExpr[BigDecimal],
    fv: Option[TExpr[BigDecimal]] = None,
    pmtType: Option[TExpr[BigDecimal]] = None,
    guess: Option[TExpr[BigDecimal]] = None
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.rate, (nper, pmt, pv, fv, pmtType, guess))

  /**
   * Smart constructor for VLOOKUP (supports text and numeric lookups).
   *
   * Example: TExpr.vlookup(TExpr.Lit("Widget A"), CellRange("A1:D10"), TExpr.Lit(4),
   * TExpr.Lit(false))
   */
  def vlookup(
    lookup: TExpr[?],
    table: CellRange,
    colIndex: TExpr[Int],
    rangeLookup: TExpr[Boolean] = Lit(true)
  ): TExpr[CellValue] =
    Call(
      FunctionSpecs.vlookup,
      (asCellValueExpr(lookup), RangeLocation.Local(table), colIndex, Some(rangeLookup))
    )

  /**
   * Smart constructor for VLOOKUP with explicit RangeLocation (supports cross-sheet lookups).
   *
   * Example: TExpr.vlookupWithLocation(TExpr.Lit("Widget A"),
   * RangeLocation.CrossSheet(SheetName("Lookup"), CellRange("A1:D10")), TExpr.Lit(4),
   * TExpr.Lit(false))
   */
  def vlookupWithLocation(
    lookup: TExpr[?],
    table: RangeLocation,
    colIndex: TExpr[Int],
    rangeLookup: TExpr[Boolean] = Lit(true)
  ): TExpr[CellValue] =
    Call(
      FunctionSpecs.vlookup,
      (asCellValueExpr(lookup), table, colIndex, Some(rangeLookup))
    )

  // Conditional aggregation function smart constructors

  /**
   * SUMIF: sum cells where criteria matches.
   *
   * Example: TExpr.sumIf(CellRange("A1:A10"), TExpr.Lit("Apple"), Some(CellRange("B1:B10")))
   */
  def sumIf(
    range: CellRange,
    criteria: TExpr[?],
    sumRange: Option[CellRange] = None
  ): TExpr[BigDecimal] =
    Call(
      FunctionSpecs.sumif,
      (range, criteria.asInstanceOf[TExpr[Any]], sumRange)
    )

  /**
   * COUNTIF: count cells where criteria matches.
   *
   * Example: TExpr.countIf(CellRange("A1:A10"), TExpr.Lit(">100"))
   */
  def countIf(range: CellRange, criteria: TExpr[?]): TExpr[BigDecimal] =
    Call(FunctionSpecs.countif, (range, criteria.asInstanceOf[TExpr[Any]]))

  /**
   * SUMIFS: sum with multiple criteria (AND logic).
   *
   * Example: TExpr.sumIfs(CellRange("C1:C10"), List((CellRange("A1:A10"), TExpr.Lit("Apple"))))
   */
  def sumIfs(
    sumRange: CellRange,
    conditions: List[(CellRange, TExpr[?])]
  ): TExpr[BigDecimal] =
    Call(
      FunctionSpecs.sumifs,
      (sumRange, conditions.map { case (r, c) => (r, c.asInstanceOf[TExpr[Any]]) })
    )

  /**
   * COUNTIFS: count with multiple criteria (AND logic).
   *
   * Example: TExpr.countIfs(List((CellRange("A1:A10"), TExpr.Lit("Apple"))))
   */
  def countIfs(conditions: List[(CellRange, TExpr[?])]): TExpr[BigDecimal] =
    Call(
      FunctionSpecs.countifs,
      conditions.map { case (r, c) => (r, c.asInstanceOf[TExpr[Any]]) }
    )

  /**
   * AVERAGEIF: average cells where criteria matches.
   *
   * Example: TExpr.averageIf(CellRange("A1:A10"), TExpr.Lit("Apple"), Some(CellRange("B1:B10")))
   */
  def averageIf(
    range: CellRange,
    criteria: TExpr[?],
    averageRange: Option[CellRange] = None
  ): TExpr[BigDecimal] =
    Call(
      FunctionSpecs.averageif,
      (range, criteria.asInstanceOf[TExpr[Any]], averageRange)
    )

  /**
   * AVERAGEIFS: average with multiple criteria (AND logic).
   *
   * Example: TExpr.averageIfs(CellRange("C1:C10"), List((CellRange("A1:A10"), TExpr.Lit("Apple"))))
   */
  def averageIfs(
    averageRange: CellRange,
    conditions: List[(CellRange, TExpr[?])]
  ): TExpr[BigDecimal] =
    Call(
      FunctionSpecs.averageifs,
      (averageRange, conditions.map { case (r, c) => (r, c.asInstanceOf[TExpr[Any]]) })
    )

  // Error handling function smart constructors

  /**
   * IFERROR: return value_if_error if value results in error.
   *
   * Example: TExpr.iferror(TExpr.Div(...), TExpr.Lit(CellValue.Number(0)))
   */
  def iferror(value: TExpr[CellValue], valueIfError: TExpr[CellValue]): TExpr[CellValue] =
    Call(FunctionSpecs.iferror, (value, valueIfError))

  /**
   * ISERROR: check if expression results in error.
   *
   * Example: TExpr.iserror(TExpr.Div(...))
   */
  def iserror(value: TExpr[CellValue]): TExpr[Boolean] =
    Call(FunctionSpecs.iserror, value)

  /**
   * ISERR: check if expression results in error (excluding #N/A).
   *
   * Example: TExpr.iserr(TExpr.Div(...))
   */
  def iserr(value: TExpr[CellValue]): TExpr[Boolean] =
    Call(FunctionSpecs.iserr, value)

  /**
   * ISNUMBER: check if value is numeric.
   *
   * Example: TExpr.isnumber(TExpr.ref(ARef("A1")))
   */
  def isnumber(value: TExpr[CellValue]): TExpr[Boolean] =
    Call(FunctionSpecs.isnumber, value)

  /**
   * ISTEXT: check if value is text.
   *
   * Example: TExpr.istext(TExpr.ref(ARef("A1")))
   */
  def istext(value: TExpr[CellValue]): TExpr[Boolean] =
    Call(FunctionSpecs.istext, value)

  /**
   * ISBLANK: check if cell is empty.
   *
   * Example: TExpr.isblank(TExpr.ref(ARef("A1")))
   */
  def isblank(value: TExpr[CellValue]): TExpr[Boolean] =
    Call(FunctionSpecs.isblank, value)

  // Rounding and math function smart constructors

  /**
   * ROUND: round number to specified digits.
   *
   * Example: TExpr.round(TExpr.Lit(2.5), TExpr.Lit(0))
   */
  def round(value: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.round, (value, numDigits))

  /**
   * ROUNDUP: round away from zero.
   *
   * Example: TExpr.roundUp(TExpr.Lit(2.1), TExpr.Lit(0))
   */
  def roundUp(value: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.roundUp, (value, numDigits))

  /**
   * ROUNDDOWN: round toward zero (truncate).
   *
   * Example: TExpr.roundDown(TExpr.Lit(2.9), TExpr.Lit(0))
   */
  def roundDown(value: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.roundDown, (value, numDigits))

  /**
   * ABS: absolute value.
   *
   * Example: TExpr.abs(TExpr.Lit(-5))
   */
  def abs(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.abs, value)

  /**
   * SQRT: square root.
   *
   * Example: TExpr.sqrt(TExpr.Lit(16))
   */
  def sqrt(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.sqrt, value)

  /**
   * MOD: modulo (remainder after division).
   *
   * Example: TExpr.mod(TExpr.Lit(5), TExpr.Lit(3))
   */
  def mod(number: TExpr[BigDecimal], divisor: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.mod, (number, divisor))

  /**
   * POWER: number raised to a power.
   *
   * Example: TExpr.power(TExpr.Lit(2), TExpr.Lit(3))
   */
  def power(number: TExpr[BigDecimal], power: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.power, (number, power))

  /**
   * LOG: logarithm to specified base.
   *
   * Example: TExpr.log(TExpr.Lit(100), TExpr.Lit(10))
   */
  def log(number: TExpr[BigDecimal], base: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.log, (number, Some(base)))

  /**
   * LN: natural logarithm (base e).
   *
   * Example: TExpr.ln(TExpr.Lit(2.718281828))
   */
  def ln(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.ln, value)

  /**
   * EXP: e raised to a power.
   *
   * Example: TExpr.exp(TExpr.Lit(1))
   */
  def exp(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.exp, value)

  /**
   * FLOOR: round down to nearest multiple of significance.
   *
   * Example: TExpr.floor(TExpr.Lit(2.5), TExpr.Lit(1))
   */
  def floor(number: TExpr[BigDecimal], significance: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.floor, (number, significance))

  /**
   * CEILING: round up to nearest multiple of significance.
   *
   * Example: TExpr.ceiling(TExpr.Lit(2.5), TExpr.Lit(1))
   */
  def ceiling(number: TExpr[BigDecimal], significance: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.ceiling, (number, significance))

  /**
   * TRUNC: truncate to specified number of decimal places.
   *
   * Example: TExpr.trunc(TExpr.Lit(8.9), TExpr.Lit(0))
   */
  def trunc(number: TExpr[BigDecimal], numDigits: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.trunc, (number, Some(numDigits)))

  /**
   * SIGN: sign of a number (1, -1, or 0).
   *
   * Example: TExpr.sign(TExpr.Lit(-5))
   */
  def sign(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.sign, value)

  /**
   * INT: round down to nearest integer (floor).
   *
   * Example: TExpr.int_(TExpr.Lit(8.9))
   */
  def int_(value: TExpr[BigDecimal]): TExpr[BigDecimal] =
    Call(FunctionSpecs.int, value)

  // Reference information function smart constructors

  /**
   * Create ROW expression.
   *
   * Example: TExpr.row(TExpr.PolyRef(ref"A5", Anchor.Relative))
   */
  def row(ref: TExpr[?]): TExpr[BigDecimal] =
    Call(FunctionSpecs.row, ref.asInstanceOf[TExpr[Any]])

  /**
   * Create COLUMN expression.
   *
   * Example: TExpr.column(TExpr.PolyRef(ref"C1", Anchor.Relative))
   */
  def column(ref: TExpr[?]): TExpr[BigDecimal] =
    Call(FunctionSpecs.column, ref.asInstanceOf[TExpr[Any]])

  /**
   * Create ROWS expression.
   *
   * Example: TExpr.rows(TExpr.RangeRef(range))
   */
  def rows(range: TExpr[?]): TExpr[BigDecimal] =
    Call(FunctionSpecs.rows, range.asInstanceOf[TExpr[Any]])

  /**
   * Create COLUMNS expression.
   *
   * Example: TExpr.columns(TExpr.RangeRef(range))
   */
  def columns(range: TExpr[?]): TExpr[BigDecimal] =
    Call(FunctionSpecs.columns, range.asInstanceOf[TExpr[Any]])

  /**
   * Create ADDRESS expression.
   *
   * Example: TExpr.address(TExpr.Lit(1), TExpr.Lit(1), TExpr.Lit(1), TExpr.Lit(true), None)
   */
  def address(
    row: TExpr[BigDecimal],
    col: TExpr[BigDecimal],
    absNum: TExpr[BigDecimal] = Lit(BigDecimal(1)),
    a1Style: TExpr[Boolean] = Lit(true),
    sheetName: Option[TExpr[String]] = None
  ): TExpr[String] =
    Call(FunctionSpecs.address, (row, col, Some(absNum), Some(a1Style), sheetName))

  // Array and advanced lookup function smart constructors

  /**
   * SUMPRODUCT: multiply corresponding elements and sum.
   *
   * Example: TExpr.sumProduct(List(CellRange.parse("A1:A3").toOption.get,
   * CellRange.parse("B1:B3").toOption.get))
   */
  def sumProduct(arrays: List[CellRange]): TExpr[BigDecimal] =
    Call(FunctionSpecs.sumproduct, arrays)

  /**
   * XLOOKUP: advanced lookup with flexible matching.
   *
   * @param lookupValue
   *   The value to search for
   * @param lookupArray
   *   The range to search in
   * @param returnArray
   *   The range to return values from (same dimensions as lookupArray)
   * @param ifNotFound
   *   Optional value to return if no match (default: #N/A error)
   * @param matchMode
   *   0=exact (default), -1=next smaller, 1=next larger, 2=wildcard
   * @param searchMode
   *   1=first-to-last (default), -1=last-to-first, 2=binary asc, -2=binary desc
   *
   * Example: TExpr.xlookup(TExpr.Lit("Apple"), lookupRange, returnRange)
   */
  def xlookup(
    lookupValue: TExpr[?],
    lookupArray: CellRange,
    returnArray: CellRange,
    ifNotFound: Option[TExpr[?]] = None,
    matchMode: TExpr[Int] = Lit(0),
    searchMode: TExpr[Int] = Lit(1)
  ): TExpr[CellValue] =
    val matchModeOpt = ifNotFound.map(_ => matchMode)
    val searchModeOpt = ifNotFound.map(_ => searchMode)
    Call(
      FunctionSpecs.xlookup,
      (
        lookupValue.asInstanceOf[TExpr[Any]],
        lookupArray,
        returnArray,
        ifNotFound.map(_.asInstanceOf[TExpr[Any]]),
        matchModeOpt,
        searchModeOpt
      )
    )

  /**
   * INDEX: get value at position in array.
   *
   * @param array
   *   The range to index into
   * @param rowNum
   *   1-based row position
   * @param colNum
   *   Optional 1-based column position (defaults to 1 for single-column ranges)
   *
   * Example: TExpr.index(range, TExpr.Lit(2), Some(TExpr.Lit(3)))
   */
  def index(
    array: CellRange,
    rowNum: TExpr[BigDecimal],
    colNum: Option[TExpr[BigDecimal]] = None
  ): TExpr[CellValue] =
    Call(FunctionSpecs.index, (array, rowNum, colNum))

  /**
   * MATCH: find position of value in array.
   *
   * @param lookupValue
   *   The value to search for
   * @param lookupArray
   *   The range to search in
   * @param matchType
   *   1=largest <= (default), 0=exact, -1=smallest >=
   *
   * Example: TExpr.matchExpr(TExpr.Lit("B"), range, TExpr.Lit(0))
   */
  def matchExpr(
    lookupValue: TExpr[?],
    lookupArray: CellRange,
    matchType: TExpr[BigDecimal] = Lit(BigDecimal(1))
  ): TExpr[BigDecimal] =
    Call(
      FunctionSpecs.matchFn,
      (lookupValue.asInstanceOf[TExpr[Any]], lookupArray, Some(matchType))
    )

  // Text function smart constructors

  /**
   * CONCATENATE text values.
   *
   * Example: TExpr.concatenate(List(TExpr.Lit("Hello"), TExpr.Lit(" "), TExpr.Lit("World")))
   */
  def concatenate(xs: List[TExpr[String]]): TExpr[String] =
    Call(FunctionSpecs.concatenate, xs)

  /**
   * LEFT substring extraction.
   *
   * Example: TExpr.left(TExpr.Lit("Hello"), TExpr.Lit(3))
   */
  def left(text: TExpr[String], n: TExpr[Int]): TExpr[String] =
    Call(FunctionSpecs.left, (text, n))

  /**
   * RIGHT substring extraction.
   *
   * Example: TExpr.right(TExpr.Lit("Hello"), TExpr.Lit(3))
   */
  def right(text: TExpr[String], n: TExpr[Int]): TExpr[String] =
    Call(FunctionSpecs.right, (text, n))

  /**
   * LEN text length.
   *
   * Returns BigDecimal to match Excel semantics.
   *
   * Example: TExpr.len(TExpr.Lit("Hello"))
   */
  def len(text: TExpr[String]): TExpr[BigDecimal] =
    Call(FunctionSpecs.len, text)

  /**
   * UPPER convert to uppercase.
   *
   * Example: TExpr.upper(TExpr.Lit("hello"))
   */
  def upper(text: TExpr[String]): TExpr[String] =
    Call(FunctionSpecs.upper, text)

  /**
   * LOWER convert to lowercase.
   *
   * Example: TExpr.lower(TExpr.Lit("HELLO"))
   */
  def lower(text: TExpr[String]): TExpr[String] =
    Call(FunctionSpecs.lower, text)

  // Date/Time function smart constructors

  /**
   * TODAY current date.
   *
   * Example: TExpr.today()
   */
  def today(): TExpr[java.time.LocalDate] = Call(FunctionSpecs.today, EmptyTuple)

  /**
   * NOW current date and time.
   *
   * Example: TExpr.now()
   */
  def now(): TExpr[java.time.LocalDateTime] = Call(FunctionSpecs.now, EmptyTuple)

  /**
   * PI mathematical constant.
   *
   * Example: TExpr.pi()
   */
  def pi(): TExpr[BigDecimal] = Call(FunctionSpecs.pi, EmptyTuple)

  /**
   * DATE construct from year, month, day.
   *
   * Example: TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21))
   */
  def date(year: TExpr[Int], month: TExpr[Int], day: TExpr[Int]): TExpr[java.time.LocalDate] =
    Call(FunctionSpecs.date, (year, month, day))

  /**
   * YEAR extract year from date.
   *
   * Returns BigDecimal to match Excel semantics and enable arithmetic composition.
   *
   * Example: TExpr.year(TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21)))
   */
  def year(date: TExpr[java.time.LocalDate]): TExpr[BigDecimal] = Call(FunctionSpecs.year, date)

  /**
   * MONTH extract month from date.
   *
   * Returns BigDecimal to match Excel semantics and enable arithmetic composition.
   *
   * Example: TExpr.month(TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21)))
   */
  def month(date: TExpr[java.time.LocalDate]): TExpr[BigDecimal] = Call(FunctionSpecs.month, date)

  /**
   * DAY extract day from date.
   *
   * Returns BigDecimal to match Excel semantics and enable arithmetic composition.
   *
   * Example: TExpr.day(TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21)))
   */
  def day(date: TExpr[java.time.LocalDate]): TExpr[BigDecimal] = Call(FunctionSpecs.day, date)

  /**
   * EOMONTH end of month N months from start.
   *
   * @param startDate
   *   The starting date
   * @param months
   *   Number of months to add (can be negative)
   *
   * Example: TExpr.eomonth(dateExpr, TExpr.Lit(1))
   */
  def eomonth(
    startDate: TExpr[java.time.LocalDate],
    months: TExpr[Int]
  ): TExpr[java.time.LocalDate] =
    Call(FunctionSpecs.eomonth, (startDate, months))

  /**
   * EDATE add months to date.
   *
   * @param startDate
   *   The starting date
   * @param months
   *   Number of months to add (can be negative)
   *
   * Example: TExpr.edate(dateExpr, TExpr.Lit(3))
   */
  def edate(startDate: TExpr[java.time.LocalDate], months: TExpr[Int]): TExpr[java.time.LocalDate] =
    Call(FunctionSpecs.edate, (startDate, months))

  /**
   * DATEDIF difference between dates.
   *
   * @param startDate
   *   The starting date
   * @param endDate
   *   The ending date
   * @param unit
   *   Unit: "Y", "M", "D", "MD", "YM", "YD"
   *
   * Example: TExpr.datedif(start, end, TExpr.Lit("Y"))
   */
  def datedif(
    startDate: TExpr[java.time.LocalDate],
    endDate: TExpr[java.time.LocalDate],
    unit: TExpr[String]
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.datedif, (startDate, endDate, unit))

  /**
   * NETWORKDAYS count working days between dates.
   *
   * @param startDate
   *   The starting date (inclusive)
   * @param endDate
   *   The ending date (inclusive)
   * @param holidays
   *   Optional range of holiday dates to exclude
   *
   * Example: TExpr.networkdays(start, end, Some(holidayRange))
   */
  def networkdays(
    startDate: TExpr[java.time.LocalDate],
    endDate: TExpr[java.time.LocalDate],
    holidays: Option[CellRange] = None
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.networkdays, (startDate, endDate, holidays))

  /**
   * WORKDAY add working days to date.
   *
   * @param startDate
   *   The starting date
   * @param days
   *   Number of working days to add (can be negative)
   * @param holidays
   *   Optional range of holiday dates to exclude
   *
   * Example: TExpr.workday(start, TExpr.Lit(5), Some(holidayRange))
   */
  def workday(
    startDate: TExpr[java.time.LocalDate],
    days: TExpr[Int],
    holidays: Option[CellRange] = None
  ): TExpr[java.time.LocalDate] =
    Call(FunctionSpecs.workday, (startDate, days, holidays))

  /**
   * YEARFRAC year fraction between dates.
   *
   * @param startDate
   *   The starting date
   * @param endDate
   *   The ending date
   * @param basis
   *   Day count basis: 0=US 30/360, 1=Actual/actual, 2=Actual/360, 3=Actual/365, 4=EU 30/360
   *
   * Example: TExpr.yearfrac(start, end, TExpr.Lit(1))
   */
  def yearfrac(
    startDate: TExpr[java.time.LocalDate],
    endDate: TExpr[java.time.LocalDate],
    basis: TExpr[Int] = Lit(0)
  ): TExpr[BigDecimal] =
    Call(FunctionSpecs.yearfrac, (startDate, endDate, Some(basis)))

  // Decoder functions for cell coercion

  /**
   * Decode cell as numeric value (Double or BigDecimal).
   *
   * Handles Formula cells by extracting the cached numeric value when available. This enables
   * nested formula evaluation where a cell reference points to another formula cell with a cached
   * result.
   */
  def decodeNumeric(cell: Cell): Either[CodecError, BigDecimal] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Number(value) => scala.util.Right(value)
      case CellValue.Formula(_, Some(CellValue.Number(cached))) =>
        // Extract cached numeric value from formula cell
        scala.util.Right(cached)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Numeric",
            actual = other
          )
        )

  /**
   * Decode cell as text/string value.
   *
   * Handles Formula cells by extracting the cached text value when available.
   */
  def decodeString(cell: Cell): Either[CodecError, String] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Text(value) => scala.util.Right(value)
      case CellValue.Formula(_, Some(CellValue.Text(cached))) =>
        scala.util.Right(cached)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Text",
            actual = other
          )
        )

  /**
   * Decode cell as integer value.
   *
   * Handles Formula cells by extracting the cached integer value when available.
   */
  def decodeInt(cell: Cell): Either[CodecError, Int] =
    import com.tjclp.xl.cells.CellValue
    def tryInt(value: BigDecimal, orig: CellValue): Either[CodecError, Int] =
      if value.isValidInt then scala.util.Right(value.toInt)
      else
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Int",
            actual = orig
          )
        )
    cell.value match
      case CellValue.Number(value) => tryInt(value, CellValue.Number(value))
      case CellValue.Formula(_, Some(CellValue.Number(cached))) =>
        tryInt(cached, CellValue.Number(cached))
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Int",
            actual = other
          )
        )

  /**
   * Decode cell as LocalDate value (extracts date from DateTime).
   *
   * Handles Formula cells by extracting the cached DateTime value when available.
   */
  def decodeDate(cell: Cell): Either[CodecError, java.time.LocalDate] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.DateTime(value) => scala.util.Right(value.toLocalDate)
      case CellValue.Formula(_, Some(CellValue.DateTime(cached))) =>
        scala.util.Right(cached.toLocalDate)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Date",
            actual = other
          )
        )

  /**
   * Decode cell as LocalDateTime value.
   *
   * Handles Formula cells by extracting the cached DateTime value when available.
   */
  def decodeDateTime(cell: Cell): Either[CodecError, java.time.LocalDateTime] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.DateTime(value) => scala.util.Right(value)
      case CellValue.Formula(_, Some(CellValue.DateTime(cached))) =>
        scala.util.Right(cached)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "DateTime",
            actual = other
          )
        )

  /**
   * Decode cell as Boolean value.
   *
   * Handles Formula cells by extracting the cached Boolean value when available.
   */
  def decodeBool(cell: Cell): Either[CodecError, Boolean] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Bool(value) => scala.util.Right(value)
      case CellValue.Formula(_, Some(CellValue.Bool(cached))) =>
        scala.util.Right(cached)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Boolean",
            actual = other
          )
        )

  /**
   * Decode cell as CellValue (always succeeds).
   *
   * Used for IFERROR/ISERROR which need to preserve the raw cell value.
   */
  def decodeCellValue(cell: Cell): Either[CodecError, CellValue] =
    scala.util.Right(cell.value)

  /**
   * Decode cell as resolved CellValue (extracts cached values, converts empty to 0).
   *
   * Used for standalone cell references (e.g., =A1, =Sheet1!B2) where the formula returns the
   * cell's "effective" value:
   *   - Number, Text, Bool, DateTime, RichText → returned as-is
   *   - Formula → returns cached value if present, or Number(0) if no cache
   *   - Empty → returns Number(0) (Excel treats empty as 0 in numeric contexts)
   *   - Error → returns the error
   *
   * This matches Excel semantics for standalone cell references.
   */
  def decodeResolvedValue(cell: Cell): Either[CodecError, CellValue] =
    import com.tjclp.xl.cells.CellValue
    val resolved = cell.value match
      case CellValue.Number(n) => CellValue.Number(n)
      case CellValue.Text(s) => CellValue.Text(s)
      case CellValue.Bool(b) => CellValue.Bool(b)
      case CellValue.DateTime(dt) => CellValue.DateTime(dt)
      case CellValue.RichText(rt) => CellValue.Text(rt.toPlainText)
      case CellValue.Formula(_, cached) =>
        cached match
          case Some(CellValue.Number(n)) => CellValue.Number(n)
          case Some(CellValue.Text(s)) => CellValue.Text(s)
          case Some(CellValue.Bool(b)) => CellValue.Bool(b)
          case Some(CellValue.DateTime(dt)) => CellValue.DateTime(dt)
          case Some(CellValue.RichText(rt)) => CellValue.Text(rt.toPlainText)
          case _ => CellValue.Number(BigDecimal(0))
      case CellValue.Error(err) => CellValue.Error(err)
      case CellValue.Empty => CellValue.Number(BigDecimal(0))
    scala.util.Right(resolved)

  // ===== Type-Coercing Decoders (Excel-compatible automatic conversion) =====

  /**
   * Decode cell as String with automatic type coercion.
   *
   * Matches Excel semantics:
   *   - Text → as-is
   *   - Number → toString (42 → "42")
   *   - Boolean → toString (true → "TRUE", false → "FALSE")
   *   - DateTime → ISO format
   *   - Formula → text representation
   *   - Empty → empty string
   */
  def decodeAsString(cell: Cell): Either[CodecError, String] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Empty => scala.util.Right("")
      case CellValue.Text(s) => scala.util.Right(s)
      case CellValue.Number(n) => scala.util.Right(n.toString)
      case CellValue.Bool(b) => scala.util.Right(if b then "TRUE" else "FALSE")
      case CellValue.DateTime(dt) => scala.util.Right(dt.toString)
      case CellValue.Formula(text, _) => scala.util.Right(text)
      case CellValue.RichText(rt) => scala.util.Right(rt.toPlainText)
      case other => scala.util.Left(CodecError.TypeMismatch("String", other))

  /**
   * Decode cell as LocalDate with automatic type coercion.
   *
   * Matches Excel semantics:
   *   - DateTime → extract date component
   *   - Number → interpret as Excel serial number (not yet implemented)
   *   - Text → parse as ISO date (not yet implemented)
   *   - Other → error
   */
  def decodeAsDate(cell: Cell): Either[CodecError, java.time.LocalDate] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.DateTime(dt) => scala.util.Right(dt.toLocalDate)
      case other =>
        // Future: could add Excel serial number → date conversion for Number
        // Future: could add ISO date string parsing for Text
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Date",
            actual = other
          )
        )

  /**
   * Decode cell as Int with automatic type coercion.
   *
   * Matches Excel semantics:
   *   - Number → toInt if valid
   *   - Boolean → 1 for TRUE, 0 for FALSE
   *   - Text → parse as number (not yet implemented)
   *   - Other → error
   */
  def decodeAsInt(cell: Cell): Either[CodecError, Int] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Number(n) if n.isValidInt => scala.util.Right(n.toInt)
      case CellValue.Bool(b) => scala.util.Right(if b then 1 else 0)
      case CellValue.Number(n) =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Int",
            actual = CellValue.Number(n)
          )
        )
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Int",
            actual = other
          )
        )

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

  // ===== Date Function Detection =====
  // Used by CLI to auto-apply date formatting when writing formulas

  /**
   * Check if expression contains any date-returning functions.
   *
   * Used to determine if a formula result should be formatted as a date in Excel. Recursively
   * checks compound expressions (arithmetic, conditionals, etc.).
   */
  def containsDateFunction(expr: TExpr[?]): Boolean = expr match
    case call: Call[?] =>
      call.spec.flags.returnsDate ||
      call.spec.argSpec
        .toValues(call.args)
        .collect { case ArgValue.Expr(e) => containsDateFunction(e) }
        .exists(identity)
    // Date-to-serial wrappers (for arithmetic)
    case DateToSerial(_) | DateTimeToSerial(_) => true
    // Arithmetic - recursively check operands
    case Add(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Sub(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Mul(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Div(l, r) => containsDateFunction(l) || containsDateFunction(r)
    // Conditionals and logical functions handled via Call args
    // Comparisons
    case Eq(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Neq(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Lt(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Lte(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Gt(l, r) => containsDateFunction(l) || containsDateFunction(r)
    case Gte(l, r) => containsDateFunction(l) || containsDateFunction(r)
    // Error handling
    // Type conversion
    case ToInt(e) => containsDateFunction(e)
    // Default: no date function
    case _ => false

  /**
   * Check if expression contains time-returning functions (NOW).
   *
   * Used to distinguish between Date format (m/d/yy) and DateTime format (m/d/yy h:mm). If true,
   * use DateTime format; otherwise use Date format.
   */
  def containsTimeFunction(expr: TExpr[?]): Boolean = expr match
    case call: Call[?] =>
      call.spec.flags.returnsTime ||
      call.spec.argSpec
        .toValues(call.args)
        .collect { case ArgValue.Expr(e) => containsTimeFunction(e) }
        .exists(identity)
    case DateTimeToSerial(_) => true
    // Arithmetic - recursively check operands
    case Add(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Sub(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Mul(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Div(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    // Conditionals and logical functions handled via Call args
    // Comparisons
    case Eq(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Neq(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Lt(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Lte(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Gt(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    case Gte(l, r) => containsTimeFunction(l) || containsTimeFunction(r)
    // Error handling
    // Type conversion
    case ToInt(e) => containsTimeFunction(e)
    // Default: no time function
    case _ => false

  // Extension methods for ergonomic formula construction

  extension (x: TExpr[BigDecimal])
    /**
     * Arithmetic operators as infix methods.
     *
     * Example: expr1 + expr2
     */
    def +(y: TExpr[BigDecimal]): TExpr[BigDecimal] = Add(x, y)
    def -(y: TExpr[BigDecimal]): TExpr[BigDecimal] = Sub(x, y)
    def *(y: TExpr[BigDecimal]): TExpr[BigDecimal] = Mul(x, y)
    def /(y: TExpr[BigDecimal]): TExpr[BigDecimal] = Div(x, y)

    /**
     * Comparison operators (return Boolean).
     *
     * Example: expr1 < expr2
     */
    def <(y: TExpr[BigDecimal]): TExpr[Boolean] = Lt(x, y)
    def <=(y: TExpr[BigDecimal]): TExpr[Boolean] = Lte(x, y)
    def >(y: TExpr[BigDecimal]): TExpr[Boolean] = Gt(x, y)
    def >=(y: TExpr[BigDecimal]): TExpr[Boolean] = Gte(x, y)

  extension (x: TExpr[Boolean])
    /**
     * Boolean operators as infix methods.
     *
     * Example: expr1 && expr2
     */
    def &&(y: TExpr[Boolean]): TExpr[Boolean] = Call(FunctionSpecs.and, List(x, y))
    def ||(y: TExpr[Boolean]): TExpr[Boolean] = Call(FunctionSpecs.or, List(x, y))
    def unary_! : TExpr[Boolean] = Call(FunctionSpecs.not, x)

  extension [A](x: TExpr[A])
    /**
     * Equality/inequality operators.
     *
     * Example: expr1 === expr2
     */
    def ===(y: TExpr[A]): TExpr[Boolean] = Eq(x, y)
    def !==(y: TExpr[A]): TExpr[Boolean] = Neq(x, y)
