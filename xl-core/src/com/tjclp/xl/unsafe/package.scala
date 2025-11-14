package com.tjclp.xl.unsafe

import com.tjclp.xl.error.{XLError, XLResult}

/**
 * Unsafe operations for XLResult.
 *
 * WARNING: These methods throw exceptions and violate purity guarantees. They are provided as an
 * escape hatch for:
 *   - Demo code and examples
 *   - REPL exploration
 *   - Situations where you can prove the operation cannot fail
 *   - Legacy code migration
 *
 * Import explicitly to opt-in to unsafe behavior:
 * {{{
 * import com.tjclp.xl.unsafe.*
 *
 * val sheet = emptySheet.put(ref"A1" -> "Hello").unsafe  // Explicit unsafe usage
 * }}}
 *
 * For production code, prefer explicit Either handling:
 * {{{
 * emptySheet.put(ref"A1" -> "Hello") match
 *   case Right(sheet) => processSheet(sheet)
 *   case Left(err) => handleError(err)
 * }}}
 *
 * @see
 *   com.tjclp.xl.error.XLResult for safe error handling patterns
 * @since 0.2.0
 */
extension [A](result: XLResult[A])
  /**
   * Unwrap result, throwing exception if Left.
   *
   * DANGER: This method throws IllegalStateException and breaks purity guarantees. Only use when
   * you can prove the operation cannot fail or in non-production code (demos, tests, REPL).
   *
   * Requires explicit import: `import com.tjclp.xl.unsafe.*`
   *
   * Example:
   * {{{
   * import com.tjclp.xl.unsafe.*
   *
   * val sheet = Sheet("Data").unsafe  // Throws if sheet creation fails
   *   .put(ref"A1" -> "Title").unsafe
   *   .put(ref"A2" -> "Data").unsafe
   * }}}
   *
   * @return
   *   The wrapped value if Right
   * @throws IllegalStateException
   *   if result is Left
   */
  def unsafe: A = result match
    case Right(value) => value
    case Left(err) => throw new IllegalStateException(s"XLResult.unsafe failed: ${err.message}")

  /**
   * Unwrap with fallback value.
   *
   * Provides a default value if the result is a Left, enabling graceful degradation without
   * exceptions.
   *
   * This method is safer than .unsafe because it doesn't throw, but still requires explicit import
   * to make the partial error handling visible.
   *
   * Example:
   * {{{
   * import com.tjclp.xl.unsafe.*
   *
   * val sheet = baseSheet.put(patch).getOrElse(baseSheet)  // Fallback to original
   * }}}
   *
   * @param default
   *   Fallback value if result is Left (lazy evaluation)
   * @return
   *   The wrapped value if Right, otherwise default
   */
  def getOrElse(default: => A): A = result.toOption.getOrElse(default)
