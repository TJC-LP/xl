package com.tjclp.xl.formula

import com.tjclp.xl.XLError

/**
 * Formula parsing errors with detailed diagnostics.
 *
 * All error types include position information (line, column) to help users locate the issue in
 * their formula string.
 *
 * These errors are pure data - no exceptions thrown.
 */
enum ParseError derives CanEqual:
  /**
   * Unexpected character encountered during parsing.
   *
   * @param char
   *   The unexpected character
   * @param position
   *   Character offset in input string (0-based)
   * @param context
   *   Surrounding context (e.g., "in arithmetic expression")
   *
   * Example: "=SUM(A1@B2)" → UnexpectedChar('@', 7, "expected operator or ')'")
   */
  case UnexpectedChar(char: Char, position: Int, context: String)

  /**
   * Unexpected end of input (incomplete formula).
   *
   * @param position
   *   Character offset where more input was expected
   * @param expected
   *   Description of what was expected
   *
   * Example: "=SUM(A1" → UnexpectedEOF(7, "expected ')'")
   */
  case UnexpectedEOF(position: Int, expected: String)

  /**
   * Invalid cell reference format.
   *
   * @param input
   *   The invalid reference string
   * @param position
   *   Character offset where reference started
   * @param reason
   *   Why the reference is invalid
   *
   * Example: "=A99999999" → InvalidCellRef("A99999999", 1, "row out of range")
   */
  case InvalidCellRef(input: String, position: Int, reason: String)

  /**
   * Invalid number literal format.
   *
   * @param input
   *   The invalid number string
   * @param position
   *   Character offset where number started
   * @param reason
   *   Why the number is invalid
   *
   * Example: "=1.2.3" → InvalidNumber("1.2.3", 1, "multiple decimal points")
   */
  case InvalidNumber(input: String, position: Int, reason: String)

  /**
   * Unbalanced parentheses or brackets.
   *
   * @param position
   *   Character offset of the mismatched delimiter
   * @param delimiter
   *   The mismatched character ('(', ')', '[', ']')
   * @param expected
   *   What was expected instead
   *
   * Example: "=SUM(A1:B2]" → UnbalancedDelimiter(10, ']', "expected ')'")
   */
  case UnbalancedDelimiter(position: Int, delimiter: Char, expected: String)

  /**
   * Unknown function name.
   *
   * @param name
   *   The unrecognized function name
   * @param position
   *   Character offset where function name started
   * @param suggestions
   *   Similar function names (for helpful error messages)
   *
   * Example: "=SUMM(A1:B2)" → UnknownFunction("SUMM", 1, List("SUM", "SUMIF"))
   */
  case UnknownFunction(name: String, position: Int, suggestions: List[String])

  /**
   * Invalid function arguments (wrong count or types).
   *
   * @param function
   *   The function name
   * @param position
   *   Character offset where arguments started
   * @param expected
   *   Description of expected arguments
   * @param actual
   *   Description of actual arguments provided
   *
   * Example: "=IF(A1)" → InvalidArguments("IF", 4, "3 arguments", "1 argument")
   */
  case InvalidArguments(function: String, position: Int, expected: String, actual: String)

  /**
   * Invalid operator usage or precedence.
   *
   * @param operator
   *   The problematic operator
   * @param position
   *   Character offset of the operator
   * @param reason
   *   Why the operator usage is invalid
   *
   * Example: "=A1 + * B2" → InvalidOperator("*", 6, "unexpected operator after '+'")
   */
  case InvalidOperator(operator: String, position: Int, reason: String)

  /**
   * Empty formula (only whitespace or just "=").
   *
   * Example: "=" → EmptyFormula
   */
  case EmptyFormula

  /**
   * Formula string too long (Excel limit: 8192 characters).
   *
   * @param length
   *   The actual length
   * @param maxLength
   *   The maximum allowed length
   *
   * Example: 10000-char formula → FormulaT
   *
   * ooLong(10000, 8192)
   */
  case FormulaTooLong(length: Int, maxLength: Int)

  /**
   * Generic parse error with custom message.
   *
   * Used for errors that don't fit other categories.
   *
   * @param message
   *   Human-readable error description
   * @param position
   *   Optional position in input string
   */
  case GenericError(message: String, position: Option[Int])

object ParseError:
  /**
   * Convert ParseError to XLError for integration with existing error handling.
   */
  def toXLError(error: ParseError, formula: String): XLError =
    val message = error match
      case UnexpectedChar(char, pos, ctx) =>
        s"Unexpected character '$char' at position $pos: $ctx"
      case UnexpectedEOF(pos, expected) =>
        s"Unexpected end of formula at position $pos: $expected"
      case InvalidCellRef(input, pos, reason) =>
        s"Invalid cell reference '$input' at position $pos: $reason"
      case InvalidNumber(input, pos, reason) =>
        s"Invalid number '$input' at position $pos: $reason"
      case UnbalancedDelimiter(pos, delim, expected) =>
        s"Unbalanced delimiter '$delim' at position $pos: $expected"
      case UnknownFunction(name, pos, suggestions) =>
        val suggest =
          if suggestions.isEmpty then ""
          else s" Did you mean: ${suggestions.mkString(", ")}?"
        s"Unknown function '$name' at position $pos$suggest"
      case InvalidArguments(func, pos, expected, actual) =>
        s"Invalid arguments for $func at position $pos: expected $expected, got $actual"
      case InvalidOperator(op, pos, reason) =>
        s"Invalid operator '$op' at position $pos: $reason"
      case EmptyFormula =>
        "Formula is empty"
      case FormulaTooLong(len, max) =>
        s"Formula too long: $len characters (max $max)"
      case GenericError(msg, posOpt) =>
        posOpt.fold(msg)(pos => s"$msg at position $pos")

    XLError.FormulaError(formula, message)

  /**
   * Format error with visual pointer to error location.
   *
   * Example output:
   * {{{
   * =SUM(A1@B2)
   *        ^
   * Unexpected character '@' at position 7: expected operator or ')'
   * }}}
   */
  def formatWithContext(error: ParseError, formula: String): String =
    val position = error match
      case UnexpectedChar(_, pos, _) => Some(pos)
      case UnexpectedEOF(pos, _) => Some(pos)
      case InvalidCellRef(_, pos, _) => Some(pos)
      case InvalidNumber(_, pos, _) => Some(pos)
      case UnbalancedDelimiter(pos, _, _) => Some(pos)
      case UnknownFunction(_, pos, _) => Some(pos)
      case InvalidArguments(_, pos, _, _) => Some(pos)
      case InvalidOperator(_, pos, _) => Some(pos)
      case EmptyFormula => None
      case FormulaTooLong(_, _) => None
      case GenericError(_, pos) => pos

    val message = toXLError(error, formula).message

    position match
      case Some(pos) if pos >= 0 && pos < formula.length =>
        val pointer = " " * pos + "^"
        s"$formula\n$pointer\n$message"
      case _ =>
        s"$formula\n$message"

  /**
   * Create a GenericError with position.
   */
  def generic(message: String, position: Int): ParseError =
    GenericError(message, Some(position))

  /**
   * Create a GenericError without position.
   */
  def generic(message: String): ParseError =
    GenericError(message, None)
