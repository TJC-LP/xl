package com.tjclp.xl.formula

/**
 * Type class for parsing Excel formula functions.
 *
 * Each parser-based function can provide a FunctionParser instance that encapsulates:
 *   - Function metadata (name, arity)
 *   - Argument validation logic
 *   - TExpr construction
 *
 * This enables:
 *   - Extensible function library for custom functions via given instances
 *   - Centralized function registry for introspection
 *   - Consistent error reporting
 *   - Type-safe construction of TExpr nodes
 *
 * Design follows CellWriter pattern: trait + given instances + companion object registry.
 */
trait FunctionParser[F]:
  /** Function name (uppercase, e.g., "SUM") */
  def name: String

  /** Arity specification for this function */
  def arity: Arity

  /**
   * Parse function arguments into TExpr.
   *
   * @param args
   *   Parsed argument expressions
   * @param pos
   *   Position in input (for error reporting)
   * @return
   *   Either ParseError or constructed TExpr
   */
  def parse(args: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]]

/**
 * Arity specification for function arguments.
 *
 * Captures argument count constraints in a type-safe way.
 */
enum Arity derives CanEqual:
  /** Exactly N arguments required (e.g., NOT takes exactly 1) */
  case Exact(n: Int)

  /** Between min and max arguments inclusive (e.g., IF with optional 4th arg for error value) */
  case Range(min: Int, max: Int)

  /** At least N arguments, no upper bound (e.g., AND, OR, CONCATENATE) */
  case AtLeast(n: Int)

object Arity:
  /** Zero arguments (TODAY, NOW) */
  def none: Arity = Exact(0)

  /** Exactly one argument (NOT, LEN, UPPER, etc.) */
  def one: Arity = Exact(1)

  /** Exactly two arguments (LEFT, RIGHT) */
  def two: Arity = Exact(2)

  /** Exactly three arguments (IF, DATE) */
  def three: Arity = Exact(3)

  /** At least one argument (AND, OR, CONCATENATE) */
  def atLeastOne: Arity = AtLeast(1)

  extension (arity: Arity)
    /**
     * Validate argument count matches arity.
     *
     * @return
     *   Either error message or unit
     */
    def validate(argCount: Int, functionName: String, pos: Int): Either[ParseError, Unit] =
      arity match
        case Exact(n) if argCount != n =>
          val expected =
            if n == 0 then "0 arguments"
            else if n == 1 then "1 argument"
            else s"$n arguments"
          scala.util.Left(
            ParseError.InvalidArguments(functionName, pos, expected, s"$argCount arguments")
          )
        case Range(min, max) if argCount < min || argCount > max =>
          scala.util.Left(
            ParseError.InvalidArguments(
              functionName,
              pos,
              s"between $min and $max arguments",
              s"$argCount arguments"
            )
          )
        case AtLeast(min) if argCount < min =>
          val expected =
            if min == 1 then "at least 1 argument" else s"at least $min arguments"
          scala.util.Left(
            ParseError.InvalidArguments(functionName, pos, expected, s"$argCount arguments")
          )
        case _ => scala.util.Right(())

/**
 * FunctionParser companion object and given instances.
 *
 * @note
 *   Suppression rationale:
 *   - AsInstanceOf: Type casts restore GADT type parameters lost during runtime parsing. Each cast
 *     is safe based on parser validation context.
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
object FunctionParser:
  /**
   * Summon a parser for function type F.
   *
   * Example: FunctionParser[TExpr.Min]
   */
  def apply[F](using p: FunctionParser[F]): FunctionParser[F] = p

  /**
   * Registry of all available functions (populated by given instances).
   *
   * Key: Function name (uppercase) Value: FunctionParser instance (type-erased)
   */
  private lazy val registry: Map[String, FunctionParser[?]] = buildRegistry()

  /**
   * Lookup function parser by name (case-insensitive).
   *
   * @param name
   *   Function name (e.g., "sum", "SUM", "Sum")
   * @return
   *   Option[FunctionParser] if found
   */
  def lookup(name: String): Option[FunctionParser[?]] =
    registry.get(name.toUpperCase)

  /**
   * All registered function names (sorted alphabetically).
   *
   * Useful for introspection, documentation generation, and autocomplete.
   */
  def allFunctions: List[String] =
    registry.keys.toList.sorted

  /**
   * Check if function name is registered.
   */
  def isKnown(name: String): Boolean =
    registry.contains(name.toUpperCase)

  /**
   * Build registry from all given instances.
   *
   * Collects all FunctionParser instances in scope at initialization time.
   */
  private def buildRegistry(): Map[String, FunctionParser[?]] =
    Map.empty
  // No built-in parser instances. FunctionSpecs covers built-ins.
