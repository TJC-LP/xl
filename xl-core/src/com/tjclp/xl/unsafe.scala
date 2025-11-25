package com.tjclp.xl

import com.tjclp.xl.error.{XLError, XLException, XLResult}

/**
 * Unsafe operations for XLResult (Scala 3 object style).
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
 * @note
 *   Exceptions thrown are [[XLException]] which wrap [[XLError]] for structured error data.
 * @see
 *   com.tjclp.xl.error.XLResult for safe error handling patterns
 * @since 0.3.0
 *   (migrated from package object to top-level object)
 */
object unsafe:
  extension [A](result: XLResult[A])
    /**
     * Unwrap result, throwing [[XLException]] if Left.
     *
     * DANGER: This method throws XLException and breaks purity guarantees. Only use when you can
     * prove the operation cannot fail or in non-production code (demos, tests, REPL).
     *
     * Requires explicit import: `import com.tjclp.xl.unsafe.*`
     *
     * Example:
     * {{{
     * import com.tjclp.xl.unsafe.*
     *
     * val sheet = Sheet("Data").unsafe  // Throws XLException if sheet creation fails
     *   .put(ref"A1" -> "Title").unsafe
     *   .put(ref"A2" -> "Data").unsafe
     * }}}
     *
     * The thrown [[XLException]] preserves the underlying [[XLError]] for programmatic error
     * recovery:
     * {{{
     * try {
     *   sheet.put("InvalidRef", "Value").unsafe
     * } catch {
     *   case ex: XLException => println(ex.error)  // Access structured error
     * }
     * }}}
     *
     * @return
     *   The wrapped value if Right
     * @throws XLException
     *   if result is Left (wraps the XLError)
     */
    def unsafe: A = result match
      case Right(value) => value
      case Left(err) => throw XLException(err)

    /**
     * Unwrap with fallback value.
     *
     * Provides a default value if the result is a Left, enabling graceful degradation without
     * exceptions.
     *
     * This method is safer than .unsafe because it doesn't throw, but still requires explicit
     * import to make the partial error handling visible.
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

  // Identity extension for Sheet - allows .unsafe chaining to work when put returns Sheet directly
  // This is a backwards compatibility measure: code written for XLResult[Sheet] continues to work
  extension (sheet: com.tjclp.xl.sheets.Sheet)
    /**
     * Identity operation for backwards compatibility.
     *
     * When Sheet.put returns Sheet directly (infallible), calling .unsafe is a no-op. This allows
     * existing code to work unchanged after the API evolved to return Sheet for validated refs.
     */
    @annotation.targetName("unsafeSheet")
    def unsafe: com.tjclp.xl.sheets.Sheet = sheet
