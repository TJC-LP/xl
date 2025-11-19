package com.tjclp.xl.errors

/**
 * Exception wrapper for [[XLError]] providing structured errors data in exception-based code.
 *
 * This class bridges the pure Either-based errors model with exception-based APIs (like Easy Mode).
 * The underlying [[XLError]] is preserved for programmatic errors recovery and structured logging.
 *
 * '''Example: Catching and inspecting errors'''
 * {{{
 * import com.tjclp.xl.*
 *
 * try {
 *   sheets.put("InvalidRef", "Value")  // Easy Mode throws XLException
 * } catch {
 *   case ex: XLException =>
 *     ex.errors match {
 *       case XLError.InvalidCellRef(ref, reason) => println(s"Bad ref: \$ref - \$reason")
 *       case other => println(s"Error: \${other.message}")
 *     }
 * }
 * }}}
 *
 * @param error
 *   The underlying structured errors
 */
final class XLException(val error: XLError) extends RuntimeException(error.message)
