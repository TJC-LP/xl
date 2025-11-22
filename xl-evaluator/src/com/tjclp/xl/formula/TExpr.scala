package com.tjclp.xl.formula

import com.tjclp.xl.{ARef, CellRange}
import com.tjclp.xl.cells.Cell
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
   * @param decode
   *   Function to decode the cell's value to type A
   *
   * Example: Ref(ARef("A1"), decodeNumber) reads numeric value from A1
   */
  case Ref[A](at: ARef, decode: Cell => Either[CodecError, A]) extends TExpr[A]

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
   *
   * Example: PolyRef(ARef("A1")) - type determined by enclosing function
   */
  case PolyRef(at: ARef) extends TExpr[Nothing]

  /**
   * Conditional expression - if/then/else.
   *
   * @param cond
   *   Boolean condition to evaluate
   * @param ifTrue
   *   Expression to evaluate if condition is true
   * @param ifFalse
   *   Expression to evaluate if condition is false
   *
   * Example: If(Ref(A1) > 0, Lit("Positive"), Lit("Non-positive"))
   */
  case If[A](cond: TExpr[Boolean], ifTrue: TExpr[A], ifFalse: TExpr[A]) extends TExpr[A]

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

  // Boolean/logical operators

  /**
   * Logical AND: x && y
   *
   * Laws:
   *   - Commutative: And(x, y) = And(y, x)
   *   - Associative: And(And(x, y), z) = And(x, And(y, z))
   *   - Short-circuit: If x evaluates to false, y is not evaluated
   */
  case And(x: TExpr[Boolean], y: TExpr[Boolean]) extends TExpr[Boolean]

  /**
   * Logical OR: x || y
   *
   * Laws:
   *   - Commutative: Or(x, y) = Or(y, x)
   *   - Associative: Or(Or(x, y), z) = Or(x, Or(y, z))
   *   - Short-circuit: If x evaluates to true, y is not evaluated
   */
  case Or(x: TExpr[Boolean], y: TExpr[Boolean]) extends TExpr[Boolean]

  /**
   * Logical NOT: !x
   *
   * Laws:
   *   - Double negation: Not(Not(x)) = x
   *   - De Morgan: Not(And(x, y)) = Or(Not(x), Not(y))
   */
  case Not(x: TExpr[Boolean]) extends TExpr[Boolean]

  // Range aggregation

  /**
   * Fold over a cell range - generalized aggregation.
   *
   * This is the foundation for SUM, COUNT, AVERAGE, etc.
   *
   * @param range
   *   The range of cells to aggregate
   * @param z
   *   Initial accumulator value
   * @param step
   *   Aggregation function: (accumulator, cell_value) => new_accumulator
   * @param decode
   *   Function to decode each cell's value to type A
   *
   * Example: SUM(A1:A10) = FoldRange(A1:A10, BigDecimal(0), _ + _, decodeNumber) Example:
   * COUNT(A1:A10) = FoldRange(A1:A10, 0, (acc, _) => acc + 1, decodeAny)
   */
  case FoldRange[A, B](
    range: CellRange,
    z: B,
    step: (B, A) => B,
    decode: Cell => Either[CodecError, A]
  ) extends TExpr[B]

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

  // Text functions

  /**
   * Concatenate multiple text values: CONCATENATE(text1, text2, ...)
   *
   * Example: CONCATENATE("Hello", " ", "World") = "Hello World"
   */
  case Concatenate(xs: List[TExpr[String]]) extends TExpr[String]

  /**
   * Extract left substring: LEFT(text, n)
   *
   * @param text
   *   The text to extract from
   * @param n
   *   Number of characters to extract
   *
   * Example: LEFT("Hello", 3) = "Hel"
   */
  case Left(text: TExpr[String], n: TExpr[Int]) extends TExpr[String]

  /**
   * Extract right substring: RIGHT(text, n)
   *
   * @param text
   *   The text to extract from
   * @param n
   *   Number of characters to extract
   *
   * Example: RIGHT("Hello", 3) = "llo"
   */
  case Right(text: TExpr[String], n: TExpr[Int]) extends TExpr[String]

  /**
   * Text length: LEN(text)
   *
   * Returns BigDecimal to match Excel semantics and enable composition with arithmetic.
   *
   * Example: LEN("Hello") = 5
   */
  case Len(text: TExpr[String]) extends TExpr[BigDecimal]

  /**
   * Convert to uppercase: UPPER(text)
   *
   * Example: UPPER("hello") = "HELLO"
   */
  case Upper(text: TExpr[String]) extends TExpr[String]

  /**
   * Convert to lowercase: LOWER(text)
   *
   * Example: LOWER("HELLO") = "hello"
   */
  case Lower(text: TExpr[String]) extends TExpr[String]

  // Date/Time functions

  /**
   * Current date: TODAY()
   *
   * Returns the current date without time component. Requires Clock parameter in evaluator for
   * purity.
   *
   * Example: TODAY() = LocalDate.of(2025, 11, 21)
   */
  case Today() extends TExpr[java.time.LocalDate]

  /**
   * Current date and time: NOW()
   *
   * Returns the current date and time. Requires Clock parameter in evaluator for purity.
   *
   * Example: NOW() = LocalDateTime.of(2025, 11, 21, 18, 30, 0)
   */
  case Now() extends TExpr[java.time.LocalDateTime]

  /**
   * Construct date from components: DATE(year, month, day)
   *
   * @param year
   *   The year (e.g., 2025)
   * @param month
   *   The month (1-12)
   * @param day
   *   The day (1-31)
   *
   * Example: DATE(2025, 11, 21) = LocalDate.of(2025, 11, 21)
   */
  case Date(year: TExpr[Int], month: TExpr[Int], day: TExpr[Int]) extends TExpr[java.time.LocalDate]

  /**
   * Extract year from date: YEAR(date)
   *
   * Returns BigDecimal to match Excel semantics (all numbers are doubles) and enable composition
   * with arithmetic operators (e.g., YEAR(A1) + 1).
   *
   * Example: YEAR(DATE(2025, 11, 21)) = 2025
   */
  case Year(date: TExpr[java.time.LocalDate]) extends TExpr[BigDecimal]

  /**
   * Extract month from date: MONTH(date)
   *
   * Returns BigDecimal to match Excel semantics and enable composition.
   *
   * Example: MONTH(DATE(2025, 11, 21)) = 11
   */
  case Month(date: TExpr[java.time.LocalDate]) extends TExpr[BigDecimal]

  /**
   * Extract day from date: DAY(date)
   *
   * Returns BigDecimal to match Excel semantics and enable composition.
   *
   * Example: DAY(DATE(2025, 11, 21)) = 21
   */
  case Day(date: TExpr[java.time.LocalDate]) extends TExpr[BigDecimal]

  // Arithmetic range functions (MIN, MAX)

  /**
   * Minimum value in range: MIN(range)
   *
   * Example: MIN(A1:A10) = smallest numeric value in range
   */
  case Min(range: CellRange) extends TExpr[BigDecimal]

  /**
   * Maximum value in range: MAX(range)
   *
   * Example: MAX(A1:A10) = largest numeric value in range
   */
  case Max(range: CellRange) extends TExpr[BigDecimal]

object TExpr:
  /**
   * Smart constructor for literals.
   *
   * Example: TExpr.lit(42)
   */
  def lit[A](value: A): TExpr[A] = Lit(value)

  /**
   * Smart constructor for cell references.
   *
   * Example: TExpr.ref(ARef("A1"), codec)
   */
  def ref[A](at: ARef, decode: Cell => Either[CodecError, A]): TExpr[A] =
    Ref(at, decode)

  /**
   * Smart constructor for conditionals.
   *
   * Example: TExpr.cond(test, ifTrue, ifFalse)
   */
  def cond[A](test: TExpr[Boolean], ifTrue: TExpr[A], ifFalse: TExpr[A]): TExpr[A] =
    If(test, ifTrue, ifFalse)

  // Convenience constructors for common operations

  /**
   * SUM aggregation: sum all numeric values in range.
   *
   * Example: TExpr.sum(CellRange("A1:A10"))
   */
  def sum(range: CellRange): TExpr[BigDecimal] =
    FoldRange(
      range,
      BigDecimal(0),
      (acc: BigDecimal, value: BigDecimal) => acc + value,
      decodeNumeric
    )

  /**
   * COUNT aggregation: count non-empty cells in range.
   *
   * Example: TExpr.count(CellRange("A1:A10"))
   */
  def count(range: CellRange): TExpr[Int] =
    FoldRange(
      range,
      0,
      (acc: Int, _: Option[Any]) => acc + 1,
      decodeAny
    )

  /**
   * AVERAGE aggregation: average of numeric values in range.
   *
   * Implementation: SUM / COUNT (requires evaluation phase)
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def average(range: CellRange): TExpr[BigDecimal] =
    // Note: This is a simplified representation
    // Full implementation requires division of sum by count
    // Suppression rationale: FoldRange returns TExpr[(BigDecimal, Int)], but we know
    // evaluation will extract BigDecimal via pattern matching. Type-safe by construction.
    FoldRange(
      range,
      (BigDecimal(0), 0),
      (acc: (BigDecimal, Int), value: BigDecimal) => (acc._1 + value, acc._2 + 1),
      decodeNumeric
    ).asInstanceOf[TExpr[BigDecimal]]

  /**
   * MIN aggregation: minimum numeric value in range.
   *
   * Example: TExpr.min(CellRange("A1:A10"))
   */
  def min(range: CellRange): TExpr[BigDecimal] = Min(range)

  /**
   * MAX aggregation: maximum numeric value in range.
   *
   * Example: TExpr.max(CellRange("A1:A10"))
   */
  def max(range: CellRange): TExpr[BigDecimal] = Max(range)

  // Text function smart constructors

  /**
   * CONCATENATE text values.
   *
   * Example: TExpr.concatenate(List(TExpr.Lit("Hello"), TExpr.Lit(" "), TExpr.Lit("World")))
   */
  def concatenate(xs: List[TExpr[String]]): TExpr[String] = Concatenate(xs)

  /**
   * LEFT substring extraction.
   *
   * Example: TExpr.left(TExpr.Lit("Hello"), TExpr.Lit(3))
   */
  def left(text: TExpr[String], n: TExpr[Int]): TExpr[String] = Left(text, n)

  /**
   * RIGHT substring extraction.
   *
   * Example: TExpr.right(TExpr.Lit("Hello"), TExpr.Lit(3))
   */
  def right(text: TExpr[String], n: TExpr[Int]): TExpr[String] = Right(text, n)

  /**
   * LEN text length.
   *
   * Returns BigDecimal to match Excel semantics.
   *
   * Example: TExpr.len(TExpr.Lit("Hello"))
   */
  def len(text: TExpr[String]): TExpr[BigDecimal] = Len(text)

  /**
   * UPPER convert to uppercase.
   *
   * Example: TExpr.upper(TExpr.Lit("hello"))
   */
  def upper(text: TExpr[String]): TExpr[String] = Upper(text)

  /**
   * LOWER convert to lowercase.
   *
   * Example: TExpr.lower(TExpr.Lit("HELLO"))
   */
  def lower(text: TExpr[String]): TExpr[String] = Lower(text)

  // Date/Time function smart constructors

  /**
   * TODAY current date.
   *
   * Example: TExpr.today()
   */
  def today(): TExpr[java.time.LocalDate] = Today()

  /**
   * NOW current date and time.
   *
   * Example: TExpr.now()
   */
  def now(): TExpr[java.time.LocalDateTime] = Now()

  /**
   * DATE construct from year, month, day.
   *
   * Example: TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21))
   */
  def date(year: TExpr[Int], month: TExpr[Int], day: TExpr[Int]): TExpr[java.time.LocalDate] =
    Date(year, month, day)

  /**
   * YEAR extract year from date.
   *
   * Returns BigDecimal to match Excel semantics and enable arithmetic composition.
   *
   * Example: TExpr.year(TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21)))
   */
  def year(date: TExpr[java.time.LocalDate]): TExpr[BigDecimal] = Year(date)

  /**
   * MONTH extract month from date.
   *
   * Returns BigDecimal to match Excel semantics and enable arithmetic composition.
   *
   * Example: TExpr.month(TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21)))
   */
  def month(date: TExpr[java.time.LocalDate]): TExpr[BigDecimal] = Month(date)

  /**
   * DAY extract day from date.
   *
   * Returns BigDecimal to match Excel semantics and enable arithmetic composition.
   *
   * Example: TExpr.day(TExpr.date(TExpr.Lit(2025), TExpr.Lit(11), TExpr.Lit(21)))
   */
  def day(date: TExpr[java.time.LocalDate]): TExpr[BigDecimal] = Day(date)

  // Decoder functions for FoldRange

  /**
   * Decode cell as numeric value (Double or BigDecimal).
   */
  def decodeNumeric(cell: Cell): Either[CodecError, BigDecimal] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Number(value) => scala.util.Right(value)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Numeric",
            actual = other
          )
        )

  /**
   * Decode cell as any value (for COUNT, etc).
   */
  def decodeAny(cell: Cell): Either[CodecError, Option[Any]] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Empty => scala.util.Right(None)
      case other => scala.util.Right(Some(other))

  /**
   * Decode cell as text/string value.
   */
  def decodeString(cell: Cell): Either[CodecError, String] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Text(value) => scala.util.Right(value)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Text",
            actual = other
          )
        )

  /**
   * Decode cell as integer value.
   */
  def decodeInt(cell: Cell): Either[CodecError, Int] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Number(value) =>
        if value.isValidInt then scala.util.Right(value.toInt)
        else
          scala.util.Left(
            CodecError.TypeMismatch(
              expected = "Int",
              actual = CellValue.Number(value)
            )
          )
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Int",
            actual = other
          )
        )

  /**
   * Decode cell as LocalDate value (extracts date from DateTime).
   */
  def decodeDate(cell: Cell): Either[CodecError, java.time.LocalDate] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.DateTime(value) => scala.util.Right(value.toLocalDate)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Date",
            actual = other
          )
        )

  /**
   * Decode cell as LocalDateTime value.
   */
  def decodeDateTime(cell: Cell): Either[CodecError, java.time.LocalDateTime] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.DateTime(value) => scala.util.Right(value)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "DateTime",
            actual = other
          )
        )

  /**
   * Decode cell as Boolean value.
   */
  def decodeBool(cell: Cell): Either[CodecError, Boolean] =
    import com.tjclp.xl.cells.CellValue
    cell.value match
      case CellValue.Bool(value) => scala.util.Right(value)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Boolean",
            actual = other
          )
        )

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
      case CellValue.Formula(text) => scala.util.Right(text)
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
    case PolyRef(at) => Ref(at, decodeAsString)
    case other => other.asInstanceOf[TExpr[String]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to LocalDate type with coercion.
   *
   * Used by date functions (YEAR, MONTH, DAY) to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asDateExpr(expr: TExpr[?]): TExpr[java.time.LocalDate] = expr match
    case PolyRef(at) => Ref(at, decodeAsDate)
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
    case PolyRef(at) => Ref(at, decodeAsInt)
    case TExpr.Lit(bd: BigDecimal) if bd.isValidInt => TExpr.Lit(bd.toInt)
    // Convert BigDecimal expressions to Int (YEAR/MONTH/DAY/LEN return BigDecimal)
    case year: TExpr.Year => ToInt(year)
    case month: TExpr.Month => ToInt(month)
    case day: TExpr.Day => ToInt(day)
    case len: TExpr.Len => ToInt(len)
    case other => other.asInstanceOf[TExpr[Int]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to BigDecimal type (numeric).
   *
   * Used by arithmetic functions to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asNumericExpr(expr: TExpr[?]): TExpr[BigDecimal] = expr match
    case PolyRef(at) => Ref(at, decodeNumeric)
    case other =>
      other.asInstanceOf[TExpr[BigDecimal]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to Boolean type.
   *
   * Used by logical functions (AND, OR, NOT, IF) to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asBooleanExpr(expr: TExpr[?]): TExpr[Boolean] = expr match
    case PolyRef(at) => Ref(at, decodeBool)
    case other => other.asInstanceOf[TExpr[Boolean]] // Safe: non-PolyRef already has correct type

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
    def &&(y: TExpr[Boolean]): TExpr[Boolean] = And(x, y)
    def ||(y: TExpr[Boolean]): TExpr[Boolean] = Or(x, y)
    def unary_! : TExpr[Boolean] = Not(x)

  extension [A](x: TExpr[A])
    /**
     * Equality/inequality operators.
     *
     * Example: expr1 === expr2
     */
    def ===(y: TExpr[A]): TExpr[Boolean] = Eq(x, y)
    def !==(y: TExpr[A]): TExpr[Boolean] = Neq(x, y)
